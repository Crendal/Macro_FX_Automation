# Weekly FX Report Automation

한국은행 ECOS 금리 · Yahoo Finance 환율 · 한국자금중개(KMBCO) IRS/CRS/스왑포인트를 수집해
엑셀 워크북을 증분 업데이트하고, 주간 FX 리포트 PPT의 표·히트맵·차트를 자동으로 교체합니다.

## 구성

| 파일 | 역할 |
|---|---|
| `get_rates.py` | ECOS 금리(기준금리·국고채·회사채 등) → `ecos_rates.xlsx` |
| `fx_weekly_updater.py` | 환율/5분봉/IRS·CRS/스왑포인트/KTB 증분 업데이트, 지표·차트 PNG 생성 |
| `ppt_weekly_updater.py` | 워크북 + 차트 PNG → 주간 PPT 표·히트맵·이미지 교체 (Windows + PowerPoint 필요) |
| `fx_common.py` | 경로 탐색, 파일 선택 창, 로컬 설정, API 키 로딩 공통 모듈 |
| `samples/` | 참고용 워크북·PPT 템플릿 |

## 설치

```bash
pip install -r requirements.txt
cp .env.example .env      # 그리고 ECOS_API_KEY 입력 (없으면 get_rates.py 첫 실행 때 입력창이 뜸)
```

Selenium은 Chrome이 설치되어 있으면 드라이버를 자동으로 받습니다 (Selenium 4.6+).

## 사용 순서

```bash
python get_rates.py              # 저장 위치 선택 창 → ecos_rates.xlsx
python fx_weekly_updater.py      # 워크북 자동 탐색 (없으면 선택 창)
python ppt_weekly_updater.py     # 업데이트할 PPT 선택 창
```

처음 시작할 때는 `samples/FX_Weekly_Automation_Light.xlsx`를 원하는 작업 폴더로 복사해서 쓰면 됩니다
(`samples/` 파일을 직접 갱신하면 커밋 diff가 커지니 권장하지 않습니다).

## 경로 규칙

어느 폴더에서 실행해도 동일하게 동작합니다 (현재 작업 디렉터리에 의존하지 않음).

1. 명령행 인자 (`--workbook`, `--pptx`, `--ecos-file`, `--assets`, `--output`)
2. 마지막으로 사용한 경로 (`.fx_local.json`, git 제외)
3. 스크립트 폴더의 기본 파일명 (`FX_Weekly_Automation_Light.xlsx`, `ecos_rates.xlsx`)
4. 파일 선택 창

- `get_rates.py`로 저장한 `ecos_rates.xlsx` 위치는 기억되어 `fx_weekly_updater.py`가 그대로 사용합니다.
- 차트 PNG는 워크북과 같은 폴더의 `fx_chart_images/`에 생성되고, PPT 업데이터도 거기서 찾습니다.
- 경로를 다시 고르고 싶으면 `--ask` 옵션, 또는 `.fx_local.json` 삭제.

## 주요 옵션

```bash
python fx_weekly_updater.py --asof 2026-09-25 --skip-selenium
python ppt_weekly_updater.py --save-as           # 원본은 두고 다른 이름으로 저장
python ppt_weekly_updater.py --pptx deck.pptx --output 261002_Weekly_FX.pptx
```

## 참고

- PPT 업데이트는 슬라이드별 도형 인덱스(`IMAGE_MAP`, `SLIDE3_TABLE_MAP` 등)에 의존합니다.
  템플릿에서 도형을 추가/삭제하면 인덱스를 다시 맞춰야 합니다.
- 데이터 출처별 이용약관을 확인하고 사용하세요.
