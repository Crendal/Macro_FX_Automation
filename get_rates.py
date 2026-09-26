import requests
import pandas as pd
from datetime import date
from pathlib import Path
import time

from fx_common import (DEFAULT_ECOS_FILE, XLSX_TYPES, ask_save_file, get_ecos_key,
                       last_dir, remember_path)

# ========= 설정 =========
# API 키는 코드에 넣지 않습니다. ECOS_API_KEY 환경변수 또는 .env 파일에서 읽고,
# 둘 다 없으면 실행 시 입력창을 띄워 .env에 저장합니다.
ECOS_KEY = None  # main()에서 채워짐
BASE = "https://ecos.bok.or.kr/api"
START = "2020-01-01"
END   = date.today().strftime("%Y-%m-%d")   
ROUND_DEC = 3
MAX_ROWS = 100000

STAT_CYCLES = {
    "722Y001": ["M"],
    "817Y002": ["D"],
    "721Y001": ["M"],
}

CYCLE_SUFFIX = {"D":"[D]","M":"[M]","Q":"[Q]","A":"[A]"}
ASCENDING_SHEETS = {"817Y002"}

# ========= 저장 경로 팝업 =========
def get_save_path():
    """마지막으로 저장했던 폴더에서 시작하는 저장 창."""
    p = ask_save_file("엑셀 저장 경로를 선택하세요", XLSX_TYPES,
                      initialfile=DEFAULT_ECOS_FILE, initialdir=last_dir())
    return str(p) if p else ""

# ========= 유틸 =========
def ecos_get(url, max_retry=3, sleep=0.5, verbose=False):
    for i in range(max_retry):
        if verbose:
            print("[GET]", url)
        r = requests.get(url, timeout=30)
        try:
            r.raise_for_status()
            js = r.json()
        except Exception:
            if i == max_retry - 1:
                raise
            time.sleep(sleep); continue
        if isinstance(js, dict) and "RESULT" in js:
            code = js["RESULT"].get("CODE")
            if code != "INFO-000":
                if i == max_retry - 1:
                    raise RuntimeError(f"ECOS API Error: {code} - {js['RESULT'].get('MESSAGE')}")
                time.sleep(sleep); continue
        return js
    raise RuntimeError("ECOS GET failed after retries")

def _fmt_date_for_cycle(dt_str, cyc):
    ts = pd.to_datetime(dt_str, errors="coerce")
    if pd.isna(ts): return None
    if cyc == "D": return ts.strftime("%Y%m%d")
    if cyc == "M": return ts.strftime("%Y%m")
    if cyc in ("Q","A"): return ts.strftime("%Y")
    return ts.strftime("%Y%m%d")

def _normalize_time(s):
    s = str(s)
    if len(s) == 8 and s.isdigit():
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    if len(s) == 6 and s.isdigit():
        return pd.to_datetime(s + "01", format="%Y%m%d", errors="coerce")
    if len(s) == 4 and s.isdigit():
        return pd.to_datetime(s + "-01-01", errors="coerce")
    if "Q" in s:
        try:
            y, q = s.split("Q")
            month = {1:1,2:4,3:7,4:10}[int(q)]
            return pd.Timestamp(year=int(y), month=month, day=1)
        except:
            return pd.NaT
    return pd.to_datetime(s, errors="coerce")

def list_items(stat_code, lang="kr", verbose=False):
    url = f"{BASE}/StatisticItemList/{ECOS_KEY}/json/{lang}/1/9999/{stat_code}"
    js = ecos_get(url, verbose=verbose)
    rows = js.get("StatisticItemList", {}).get("row", [])
    if not rows and lang == "kr":
        url = f"{BASE}/StatisticItemList/{ECOS_KEY}/json/en/1/9999/{stat_code}"
        js = ecos_get(url, verbose=verbose)
        rows = js.get("StatisticItemList", {}).get("row", [])
    if not rows:
        return pd.DataFrame(columns=["code","name","unit"])
    df = pd.DataFrame(rows)
    code_col = "ITEM_CODE1" if "ITEM_CODE1" in df.columns else "ITEM_CODE"
    name_col = "ITEM_NAME1" if "ITEM_NAME1" in df.columns else "ITEM_NAME"
    unit_col = "UNIT_NAME" if "UNIT_NAME" in df.columns else None
    out = pd.DataFrame({
        "code": df[code_col].astype(str),
        "name": df[name_col].astype(str),
        "unit": df[unit_col].astype(str) if unit_col else ""
    }).drop_duplicates()
    return out

def fetch_series(stat_code, item_code, cycle, start, end, lang="kr", verbose=False):
    s = _fmt_date_for_cycle(start, cycle)
    e = _fmt_date_for_cycle(end, cycle)
    if not s or not e:
        raise ValueError("날짜 포맷 변환 실패")
    url = f"{BASE}/StatisticSearch/{ECOS_KEY}/json/{lang}/1/{MAX_ROWS}/{stat_code}/{cycle}/{s}/{e}/{item_code}"
    js = ecos_get(url, verbose=verbose)
    rows = js.get("StatisticSearch", {}).get("row", [])
    if not rows:
        return pd.DataFrame(columns=["date","value","unit"])
    df = pd.DataFrame(rows)
    unit_col = "UNIT_NAME" if "UNIT_NAME" in df.columns else None
    df["date"]  = df["TIME"].apply(_normalize_time)
    df["value"] = pd.to_numeric(df["DATA_VALUE"], errors="coerce")
    df["unit"]  = df[unit_col] if unit_col else ""
    return df[["date","value","unit"]].dropna(subset=["date"]).sort_values("date")

def build_sheet_one(stat_code, cycles, start, end, round_dec=3, verbose=False):
    items = list_items(stat_code, "kr", verbose=verbose)
    if items.empty:
        return None
    frames = []
    for cyc in cycles:
        for _, r in items.iterrows():
            ser = fetch_series(stat_code, r["code"], cyc, start, end, "kr", verbose=verbose)
            if ser.empty:
                continue
            col = f"{r['name']} {CYCLE_SUFFIX.get(cyc, '')}"
            tmp = ser.rename(columns={"value": col})[["date", col]]
            frames.append(tmp)
    if not frames:
        return None
    wide = None
    for df in frames:
        wide = df if wide is None else wide.merge(df, on="date", how="outer")
    wide = wide.sort_values("date", ascending=False).set_index("date").round(round_dec)
    wide.index = wide.index.strftime("%Y-%m-%d")
    return wide

def expand_monthly_rate(monthly_series, target_dates):
    monthly_series.index = pd.to_datetime(monthly_series.index)
    target_dates = pd.to_datetime(target_dates)
    expanded = pd.Series(index=target_dates, dtype=float)
    for dt in monthly_series.index:
        mask = (target_dates.year == dt.year) & (target_dates.month == dt.month)
        expanded.loc[mask] = monthly_series.loc[dt]
    return expanded

def build_weekly(df: pd.DataFrame) -> pd.DataFrame:
    df.index = pd.to_datetime(df.index)
    weekly = df.resample("W-FRI").last()
    weekly = weekly.sort_index(ascending=False)
    weekly.index = weekly.index.strftime("%Y-%m-%d")
    return weekly

def build_key_rates(daily_722: pd.DataFrame, daily_817: pd.DataFrame) -> pd.DataFrame:
    base_rate = daily_722.iloc[:, 0].copy()
    base_rate.name = "한국은행 기준금리"
    base_expanded = expand_monthly_rate(base_rate, daily_817.index)

    column_map = {
        "CD금리": "CD(91일) [D]",
        "국고채(1년)": "국고채(1년) [D]",
        "국고채(2년)": "국고채(2년) [D]",
        "국고채(3년)": "국고채(3년) [D]",
        "국고채(5년)": "국고채(5년) [D]",
        "국고채(10년)": "국고채(10년) [D]",
        "국고채(20년)": "국고채(20년) [D]",
        "국고채(30년)": "국고채(30년) [D]",
        "국고채(50년)": "국고채(50년) [D]",
    }

    selected = {}
    for label, exact_col in column_map.items():
        if exact_col in daily_817.columns:
            selected[label] = daily_817[exact_col]
        else:
            print(f"[WARN] 열 '{exact_col}' ({label}) 추출 실패")

    df_key = pd.DataFrame(selected)
    df_key.insert(0, "한국은행 기준금리", base_expanded)
    df_key.index = df_key.index.strftime("%Y-%m-%d")
    return df_key

# ✅ 캡처 기반 채권 Weekly 테이블 생성
def build_weekly_bond_table(df_817):
    df_817.index = pd.to_datetime(df_817.index)

    bond_cols = [
        "CP(91일) [D]",
        "CD(91일) [D]",
        "국고채(3년) [D]",
        "국고채(5년) [D]",
        "국고채(10년) [D]",
        "국고채(20년) [D]",
        "국고채(30년) [D]",
        "국민주택채권1종(5년) [D]",
        "산금채(1년) [D]",
        "회사채(3년, AA-) [D]",
        "회사채(3년, BBB-) [D]",
    ]

    valid_cols = [col for col in bond_cols if col in df_817.columns]
    weekly = df_817[valid_cols].resample("W-FRI").last()
    weekly = weekly.sort_index(ascending=False)  # 최신 날짜부터 정렬
    weekly.insert(0, "date", weekly.index.strftime("%Y-%m-%d"))
    weekly = weekly.round(3)

    return weekly

# ========= 메인 =========
def main(verbose=False):
    global ECOS_KEY
    ECOS_KEY = get_ecos_key(prompt_if_missing=True)
    if not ECOS_KEY:
        print("❌ ECOS API 키가 없습니다. ECOS_API_KEY 환경변수나 .env 파일을 설정하세요.")
        return

    save_path = get_save_path()
    if not save_path:
        print("❌ 저장 취소됨. 프로그램 종료.")
        return

    wrote_any = False
    data_cache = {}

    with pd.ExcelWriter(save_path, engine="openpyxl") as xw:
        for stat_code, cycles in STAT_CYCLES.items():
            try:
                wide = build_sheet_one(stat_code, cycles, START, END, ROUND_DEC, verbose=verbose)
                sheet = stat_code
                if wide is None:
                    print(f"[SKIP] {sheet}: 데이터 없음")
                    continue
                # 일별 금리(817Y002)는 오래된 날짜 -> 최신 순으로 저장
                # (예전 파일 하단의 수동 재정렬 코드를 여기로 옮김)
                out = wide.sort_index(ascending=True) if stat_code in ASCENDING_SHEETS else wide
                out.to_excel(xw, sheet_name=sheet)
                data_cache[stat_code] = wide
                wrote_any = True
                print(f"[OK] {sheet}: {wide.shape[0]} rows x {wide.shape[1]} cols")

                # ✅ 주간 시트 추가
                if stat_code == "817Y002":
                    weekly = build_weekly(wide)
                    weekly.to_excel(xw, sheet_name="817Y002_weekly")
                    print(f"[OK] 817Y002_weekly: {weekly.shape[0]} rows")

            except Exception as e:
                print(f"[ERROR] {stat_code} - {e}")

        # ✅ Key Rates 시트 생성
        try:
            if "722Y001" in data_cache and "817Y002" in data_cache:
                key_rates = build_key_rates(data_cache["722Y001"], data_cache["817Y002"])
                key_rates.to_excel(xw, sheet_name="Key Rates")
                print(f"[OK] Key Rates: {key_rates.shape[0]} rows x {key_rates.shape[1]} cols")

                key_rates_weekly = build_weekly(key_rates)
                key_rates_weekly.to_excel(xw, sheet_name="Key Rates Weekly")
                print(f"[OK] Key Rates Weekly: {key_rates_weekly.shape[0]} rows x {key_rates_weekly.shape[1]} cols")
        except Exception as e:
            print(f"[ERROR] Key Rates 생성 실패 - {e}")

        # ✅ M열부터 채권 Weekly 테이블 삽입
        try:
            if "817Y002" in data_cache:
                bond_weekly = build_weekly_bond_table(data_cache["817Y002"])
                bond_weekly.to_excel(xw, sheet_name="Key Rates Weekly", startcol=12, index=False)
                print(f"[OK] 채권 Weekly 테이블 삽입: {bond_weekly.shape[0]} rows x {bond_weekly.shape[1]} cols")
        except Exception as e:
            print(f"[ERROR] 채권 Weekly 테이블 생성 실패 - {e}")

        if not wrote_any:
            pd.DataFrame({"msg":["no data written"]}).to_excel(xw, sheet_name="EMPTY")

    remember_path("ecos_file", Path(save_path))   # fx_weekly_updater가 이 경로를 자동으로 사용
    print(f"\n완료: {Path(save_path).resolve()}")

if __name__ == "__main__":
    main(verbose=False)

