# -*- coding: utf-8 -*-
"""Shared helpers: path resolution, file dialogs, local settings, ECOS key.

Path resolution order (every script)
------------------------------------
1. Command-line argument (--workbook, --pptx, ...)
2. Last path remembered in .fx_local.json (only if the file still exists)
3. Default file name in the script folder
4. File dialog (tkinter). Falls back to console input() if no GUI is available.

Nothing here depends on the current working directory, so the scripts behave
the same whether they are run from VS Code, a terminal, or a double-click.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Iterable, Optional, Sequence, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
LOCAL_SETTINGS = SCRIPT_DIR / ".fx_local.json"   # git-ignored
ENV_FILE = SCRIPT_DIR / ".env"                   # git-ignored

DEFAULT_WORKBOOK = "FX_Weekly_Automation_Light.xlsx"
DEFAULT_ECOS_FILE = "ecos_rates.xlsx"
CHART_DIR_NAME = "fx_chart_images"

XLSX_TYPES = [("Excel files", "*.xlsx"), ("All files", "*.*")]
PPTX_TYPES = [("PowerPoint files", "*.pptx"), ("All files", "*.*")]


# ------------------------- local settings -------------------------
def load_settings() -> dict:
    try:
        return json.loads(LOCAL_SETTINGS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_setting(key: str, value) -> None:
    data = load_settings()
    data[key] = str(value)
    try:
        LOCAL_SETTINGS.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass  # remembering paths is a convenience, never a hard failure


def remember_path(key: str, path: Path) -> None:
    path = Path(path).resolve()
    save_setting(key, path)
    save_setting("last_dir", path.parent)


def remembered_path(key: str) -> Optional[Path]:
    val = load_settings().get(key)
    if not val:
        return None
    p = Path(val)
    return p if p.exists() else None


def last_dir() -> Path:
    val = load_settings().get("last_dir")
    if val and Path(val).is_dir():
        return Path(val)
    return SCRIPT_DIR


# ------------------------- .env / ECOS key -------------------------
def load_dotenv(path: Path = ENV_FILE) -> None:
    """Minimal .env reader (KEY=VALUE lines). Existing env vars win."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def get_ecos_key(cli_key: Optional[str] = None, prompt_if_missing: bool = False) -> Optional[str]:
    """--ecos-api-key -> ECOS_API_KEY env var -> .env file -> (optional) prompt."""
    if cli_key and cli_key.strip():
        return cli_key.strip()
    load_dotenv()
    key = os.getenv("ECOS_API_KEY", "").strip()
    if key:
        return key
    if not prompt_if_missing:
        return None

    key = _ask_string("ECOS API Key", "한국은행 ECOS API 키를 입력하세요\n(.env 파일에 저장되어 다음부터는 묻지 않습니다)")
    if key:
        key = key.strip()
        try:
            with ENV_FILE.open("a", encoding="utf-8") as f:
                f.write(f"\nECOS_API_KEY={key}\n")
            print(f"[INFO] ECOS 키를 {ENV_FILE} 에 저장했습니다.")
        except Exception as exc:
            print(f"[WARN] .env 저장 실패: {exc}")
        os.environ["ECOS_API_KEY"] = key
        return key
    return None


# ------------------------- dialogs -------------------------
def _tk_root():
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def _ask_string(title: str, prompt: str) -> Optional[str]:
    try:
        from tkinter import simpledialog
        root = _tk_root()
        try:
            return simpledialog.askstring(title, prompt, parent=root)
        finally:
            root.destroy()
    except Exception:
        try:
            return input(f"{prompt}: ").strip() or None
        except EOFError:
            return None


def ask_open_file(title: str, filetypes: Sequence[Tuple[str, str]], initialdir: Optional[Path] = None) -> Optional[Path]:
    try:
        from tkinter.filedialog import askopenfilename
        root = _tk_root()
        try:
            p = askopenfilename(title=title, filetypes=list(filetypes), initialdir=str(initialdir or last_dir()))
        finally:
            root.destroy()
    except Exception:
        p = input(f"{title} (파일 경로 입력): ").strip().strip('"')
    return Path(p).resolve() if p else None


def ask_save_file(title: str, filetypes: Sequence[Tuple[str, str]], initialfile: str,
                  initialdir: Optional[Path] = None) -> Optional[Path]:
    try:
        from tkinter.filedialog import asksaveasfilename
        root = _tk_root()
        try:
            p = asksaveasfilename(title=title, filetypes=list(filetypes), initialfile=initialfile,
                                  initialdir=str(initialdir or last_dir()),
                                  defaultextension=Path(initialfile).suffix or None)
        finally:
            root.destroy()
    except Exception:
        p = input(f"{title} (저장 경로 입력): ").strip().strip('"')
    return Path(p).resolve() if p else None


def ask_directory(title: str, initialdir: Optional[Path] = None) -> Optional[Path]:
    try:
        from tkinter.filedialog import askdirectory
        root = _tk_root()
        try:
            p = askdirectory(title=title, initialdir=str(initialdir or last_dir()))
        finally:
            root.destroy()
    except Exception:
        p = input(f"{title} (폴더 경로 입력): ").strip().strip('"')
    return Path(p).resolve() if p else None


# ------------------------- resolver -------------------------
def resolve_file(
    cli_value: Optional[str],
    *,
    key: str,
    title: str,
    filetypes: Sequence[Tuple[str, str]],
    default_names: Iterable[str] = (),
    search_dirs: Iterable[Path] = (),
    force_dialog: bool = False,
    must_exist: bool = True,
) -> Path:
    """Resolve an input file: CLI -> remembered -> default names -> dialog."""
    if cli_value:
        p = Path(cli_value).expanduser().resolve()
        if must_exist and not p.exists():
            raise FileNotFoundError(p)
        remember_path(key, p)
        return p

    if not force_dialog:
        p = remembered_path(key)
        if p is not None:
            return p
        dirs = list(search_dirs) or [SCRIPT_DIR]
        for d in dirs:
            for name in default_names:
                cand = Path(d) / name
                if cand.exists():
                    remember_path(key, cand)
                    return cand.resolve()

    p = ask_open_file(title, filetypes)
    if p is None:
        raise SystemExit(f"❌ 선택 취소됨: {title}")
    remember_path(key, p)
    return p
