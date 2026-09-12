"""
자동 지오코더 v2.17
- VWorld 주소 API 기반 도로명/지번주소 좌표 변환
- xlsx/xls/csv/tsv 입력 -> xlsx 출력
- 시도/시군구 자동 보완
- 도로명+지번주소 교차검증
- 100건 자동저장 및 중지 후 이어하기
"""
from __future__ import annotations

import os
import sys
import json
import math
import time
import urllib3
import requests
import tempfile
import threading
import pandas as pd

from tqdm import tqdm
from pathlib import Path
from datetime import datetime
from pyproj import Transformer
from tkinter import ttk
from tkinter import font as tkfont
from tkinter import Tk, Canvas, StringVar, BooleanVar, DoubleVar, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


APP_NAME = "자동 지오코더"
APP_VERSION = "2.17"
API_URL = "https://api.vworld.kr/req/address"
API_CRS = "EPSG:4326"
RESULT_COLUMNS = ["X좌표", "Y좌표", "사용주소", "검증여부"]
INVALID_VALUES = {"", "미기재", "없음", "NULL", "null", "None", "nan", "-"}
COMMON_CRS = ["EPSG:4326", "EPSG:5179", "EPSG:5181", "EPSG:5186", "EPSG:3857"]

# 시도명이 정식명칭/약칭으로 이미 주소에 들어 있는지 확인하기 위한 별칭
SIDO_ALIASES = {
    "서울특별시": {"서울특별시", "서울"},
    "부산광역시": {"부산광역시", "부산"},
    "대구광역시": {"대구광역시", "대구"},
    "인천광역시": {"인천광역시", "인천"},
    "광주광역시": {"광주광역시", "광주"},
    "대전광역시": {"대전광역시", "대전"},
    "울산광역시": {"울산광역시", "울산"},
    "세종특별자치시": {"세종특별자치시", "세종"},
    "경기도": {"경기도", "경기"},
    "강원특별자치도": {"강원특별자치도", "강원도", "강원"},
    "충청북도": {"충청북도", "충북"},
    "충청남도": {"충청남도", "충남"},
    "전북특별자치도": {"전북특별자치도", "전라북도", "전북"},
    "전라남도": {"전라남도", "전남"},
    "경상북도": {"경상북도", "경북"},
    "경상남도": {"경상남도", "경남"},
    "제주특별자치도": {"제주특별자치도", "제주도", "제주"},
}


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config() -> dict:
    path = app_dir() / "config.json"
    default = {
        "api_key": os.getenv("VWORLD_API_KEY", ""),
        "ssl_verify": True,
        "request_interval": 0.05,
        "request_timeout": 15,
        "max_retry": 3,
        "save_every": 100,
        "validation_distance_m": 1000,
        "default_sido": "",
        "default_sigungu": "",
    }
    if not path.exists():
        return default
    try:
        with path.open("r", encoding="utf-8") as f:
            user_cfg = json.load(f)
        default.update(user_cfg)
    except Exception:
        pass
    return default


CONFIG = load_config()


def is_valid_value(value) -> bool:
    if pd.isna(value):
        return False
    return str(value).strip() not in INVALID_VALUES


def clean_address(value) -> str | None:
    if not is_valid_value(value):
        return None
    return " ".join(str(value).strip().split())


def normalize_space(value: str | None) -> str:
    return " ".join((value or "").strip().split())


def sido_aliases(sido: str) -> set[str]:
    sido = normalize_space(sido)
    if not sido:
        return set()
    if sido in SIDO_ALIASES:
        return SIDO_ALIASES[sido]
    for canonical, aliases in SIDO_ALIASES.items():
        if sido in aliases:
            return aliases | {canonical}
    # 알 수 없는 시도명도 정확 문자열은 중복 방지
    return {sido}


def apply_region_prefix(address: str | None, sido: str, sigungu: str) -> str | None:
    """주소에 시도/시군구가 빠져 있을 때만 앞에 보완한다."""
    address = clean_address(address)
    if not address:
        return None

    sido = normalize_space(sido)
    sigungu = normalize_space(sigungu)
    result = address

    if sigungu and sigungu not in result:
        result = f"{sigungu} {result}"

    aliases = sido_aliases(sido)
    if sido and not any(alias and alias in result for alias in aliases):
        result = f"{sido} {result}"

    return normalize_space(result)


def _pick_region_column(columns, keyword: str) -> str | None:
    """시도/시군구 컬럼 자동 탐지. 정확명/명칭 컬럼을 우선하고 코드 컬럼은 제외한다."""
    cols = [str(c) for c in columns]
    normalized = {c: c.replace(" ", "") for c in cols}

    for target in (keyword, f"{keyword}명"):
        for c in cols:
            if normalized[c] == target:
                return c

    candidates = [
        c for c in cols
        if keyword in normalized[c]
        and "코드" not in normalized[c]
        and "code" not in normalized[c].lower()
    ]
    return candidates[0] if candidates else None


def _pick_place_column(columns) -> str | None:
    """장소명 컬럼 자동 탐지. 정확히 '장소명'인 컬럼을 가장 먼저 사용한다."""
    cols = [str(c) for c in columns]
    normalized = {c: c.replace(" ", "") for c in cols}
    for c in cols:
        if normalized[c] == "장소명":
            return c
    candidates = [c for c in cols if "장소명" in normalized[c]]
    return candidates[0] if candidates else None


def _short_place_name(value, limit: int = 6) -> str:
    if not is_valid_value(value):
        return "장소명없음"
    text = normalize_space(str(value))
    return text if len(text) <= limit else text[:limit] + "..."


def detect_columns(columns) -> tuple[str | None, str | None, str | None, str | None]:
    road_candidates = [str(c) for c in columns if "도로명주소" in str(c).replace(" ", "")]
    parcel_candidates = [str(c) for c in columns if "지번주소" in str(c).replace(" ", "")]
    return (
        road_candidates[0] if road_candidates else None,
        parcel_candidates[0] if parcel_candidates else None,
        _pick_region_column(columns, "시도"),
        _pick_region_column(columns, "시군구"),
    )


def detect_address_columns(columns) -> tuple[str | None, str | None]:
    road, parcel, _, _ = detect_columns(columns)
    return road, parcel


def read_delimited(path: Path, sep: str) -> tuple[pd.DataFrame, str]:
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return pd.read_csv(path, sep=sep, encoding=enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError("문자 인코딩을 자동 판별하지 못했습니다. UTF-8 또는 CP949 파일인지 확인해 주세요.")


def load_input_file(path: Path) -> tuple[pd.DataFrame, str, str | None, str | None, str]:
    ext = path.suffix.lower()
    if ext in {".xlsx", ".xls"}:
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        xls = pd.ExcelFile(path, engine=engine)
        selected_sheet = None
        selected_road = None
        selected_parcel = None
        for sheet in xls.sheet_names:
            header = pd.read_excel(path, sheet_name=sheet, nrows=0, engine=engine)
            road, parcel = detect_address_columns(header.columns)
            if road or parcel:
                selected_sheet, selected_road, selected_parcel = sheet, road, parcel
                break
        if selected_sheet is None:
            selected_sheet = xls.sheet_names[0]
            df = pd.read_excel(path, sheet_name=selected_sheet, engine=engine)
            road, parcel = detect_address_columns(df.columns)
        else:
            df = pd.read_excel(path, sheet_name=selected_sheet, engine=engine)
            road, parcel = selected_road, selected_parcel
        return df, selected_sheet, road, parcel, f"Excel 시트: {selected_sheet}"

    if ext == ".csv":
        df, enc = read_delimited(path, ",")
        road, parcel = detect_address_columns(df.columns)
        return df, "수집결과", road, parcel, f"CSV 인코딩: {enc}"

    if ext == ".tsv":
        df, enc = read_delimited(path, "\t")
        road, parcel = detect_address_columns(df.columns)
        return df, "수집결과", road, parcel, f"TSV 인코딩: {enc}"

    raise ValueError("지원하지 않는 파일 형식입니다. xlsx, xls, csv, tsv만 사용할 수 있습니다.")


def haversine_m(lon1: float, lat1: float, lon2: float, lat2: float) -> float:
    r = 6371008.8
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def source_signature(path: Path) -> dict:
    stat = path.stat()
    return {
        "path": str(path.resolve()),
        "size": int(stat.st_size),
        "mtime_ns": int(stat.st_mtime_ns),
    }


def resume_path_for(output_path: Path) -> Path:
    return output_path.with_name(output_path.name + ".resume.json")


def load_resume_state(output_path: Path) -> dict | None:
    state_path = resume_path_for(output_path)
    if not state_path.exists():
        return None
    try:
        with state_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_resume_state(output_path: Path, state: dict) -> None:
    state_path = resume_path_for(output_path)
    tmp = state_path.with_suffix(state_path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, state_path)


def remove_resume_state(output_path: Path) -> None:
    path = resume_path_for(output_path)
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def contiguous_processed_count(df: pd.DataFrame) -> int:
    if "검증여부" not in df.columns:
        return 0
    count = 0
    for value in df["검증여부"]:
        if is_valid_value(value):
            count += 1
        else:
            break
    return count


def choose_korean_ui_font(root: Tk) -> str:
    """Windows에 설치된 한글 UI 폰트를 선택한다."""
    installed = set(tkfont.families(root))
    preferred = [
        "Pretendard",
        "Noto Sans KR",
        "맑은 고딕",
        "Malgun Gothic",
        "나눔스퀘어라운드",
        "NanumSquareRound",
        "나눔고딕",
        "NanumGothic",
    ]
    for name in preferred:
        if name in installed:
            return name
    return "TkDefaultFont"


class RoundedButton(Canvas):
    """Tkinter Canvas 기반의 둥근 모서리 버튼."""

    def __init__(
        self,
        parent,
        text: str,
        command=None,
        *,
        width: int = 118,
        height: int = 38,
        radius: int = 11,
        fill: str = "#D5A2B4",
        hover_fill: str = "#C88EA3",
        foreground: str = "#241B1F",
        border: str = "#A77989",
        canvas_bg: str = "#E8E1E4",
        font=None,
        state: str = "normal",
    ):
        super().__init__(
            parent,
            width=width,
            height=height,
            highlightthickness=0,
            bd=0,
            relief="flat",
            bg=canvas_bg,
            cursor="hand2" if state == "normal" else "arrow",
        )
        self._button_text = text
        self._command = command
        self._width = width
        self._height = height
        self._radius = radius
        self._fill = fill
        self._hover_fill = hover_fill
        self._foreground = foreground
        self._border = border
        self._font = font
        self._state = state
        self._hovered = False
        self._draw()
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonRelease-1>", self._on_click)

    @staticmethod
    def _rounded_points(x1, y1, x2, y2, r):
        return [
            x1 + r, y1, x2 - r, y1,
            x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r,
            x1, y1 + r, x1, y1,
        ]

    def _draw(self):
        self.delete("all")
        disabled = self._state != "normal"
        fill = "#CFC5C9" if disabled else (self._hover_fill if self._hovered else self._fill)
        fg = "#81767A" if disabled else self._foreground
        border = "#B9AEB2" if disabled else self._border
        self.create_polygon(
            self._rounded_points(1, 1, self._width - 1, self._height - 1, self._radius),
            smooth=True,
            splinesteps=24,
            fill=fill,
            outline=border,
            width=1,
        )
        self.create_text(
            self._width / 2,
            self._height / 2,
            text=self._button_text,
            fill=fg,
            font=self._font,
        )
        self.configure(cursor="hand2" if not disabled else "arrow")

    def _on_enter(self, _event=None):
        if self._state == "normal":
            self._hovered = True
            self._draw()

    def _on_leave(self, _event=None):
        self._hovered = False
        self._draw()

    def _on_click(self, _event=None):
        if self._state == "normal" and callable(self._command):
            self._command()

    def configure(self, cnf=None, **kwargs):
        if cnf and isinstance(cnf, dict):
            kwargs.update(cnf)
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            self._hovered = False
            self._draw()
        if kwargs:
            return super().configure(**kwargs)
        return None

    config = configure


class VWorldGeocoder:
    def __init__(self, api_key: str, ssl_verify: bool, timeout: int, max_retry: int, request_interval: float):
        self.api_key = api_key.strip()
        self.ssl_verify = bool(ssl_verify)
        self.timeout = int(timeout)
        self.max_retry = int(max_retry)
        self.request_interval = float(request_interval)
        self.cache: dict[tuple[str, str], dict] = {}
        self.session = requests.Session()
        if not self.ssl_verify:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def get_coord(self, address: str | None, address_type: str) -> dict:
        if not address:
            return {"success": False, "status": "주소없음", "x": None, "y": None, "error": ""}
        key = (address_type, address)
        if key in self.cache:
            return self.cache[key]

        params = {
            "service": "address",
            "version": "2.0",
            "request": "GetCoord",
            "key": self.api_key,
            "format": "json",
            "errorFormat": "json",
            "type": address_type,
            "address": address,
            "refine": "true",
            "simple": "false",
            "crs": API_CRS,
        }
        last_error = ""
        result = None
        for attempt in range(1, self.max_retry + 1):
            try:
                response = self.session.get(API_URL, params=params, timeout=self.timeout, verify=self.ssl_verify)
                response.raise_for_status()
                data = response.json()
                response_data = data.get("response", {})
                status = response_data.get("status", "")
                if status == "OK":
                    point = response_data.get("result", {}).get("point", {})
                    x, y = point.get("x"), point.get("y")
                    if x is not None and y is not None:
                        result = {"success": True, "status": "성공", "x": float(x), "y": float(y), "error": ""}
                    else:
                        result = {"success": False, "status": "EMPTY_POINT", "x": None, "y": None, "error": "좌표 응답이 비어 있습니다."}
                    break
                if status == "NOT_FOUND":
                    result = {"success": False, "status": "NOT_FOUND", "x": None, "y": None, "error": ""}
                    break
                result = {
                    "success": False,
                    "status": "API_ERROR",
                    "x": None,
                    "y": None,
                    "error": str(response_data.get("error", {})),
                }
                break
            except requests.exceptions.RequestException as e:
                last_error = str(e)
                if attempt < self.max_retry:
                    time.sleep(1)
            except ValueError as e:
                last_error = f"JSON 파싱 오류: {e}"
                break
            except Exception as e:
                last_error = str(e)
                break

        if result is None:
            result = {"success": False, "status": "REQUEST_ERROR", "x": None, "y": None, "error": last_error}
        self.cache[key] = result
        time.sleep(self.request_interval)
        return result


class TkTqdmWriter:
    def __init__(self, callback):
        self.callback = callback

    def write(self, text):
        text = text.replace("\r", "").strip()
        if text:
            self.callback(text)

    def flush(self):
        pass


class AutoGeocoderApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1000x800")
        self.root.minsize(920, 740)

        self.input_path = StringVar()
        self.output_path = StringVar()
        self.crs_var = StringVar(value="EPSG:4326")
        self.validation_var = DoubleVar(value=float(CONFIG.get("validation_distance_m", 1000)))
        self.sido_var = StringVar(value=str(CONFIG.get("default_sido", "")))
        self.sigungu_var = StringVar(value=str(CONFIG.get("default_sigungu", "")))
        self.ssl_var = BooleanVar(value=bool(CONFIG.get("ssl_verify", True)))
        self.detected_var = StringVar(value="입력 파일을 선택하면 주소 컬럼을 자동으로 찾습니다.")
        self.status_var = StringVar(value="대기 중")
        self.progress_text_var = StringVar(value="지오코딩: 0% | 0/0건")

        self.is_running = False
        self.cancel_event = threading.Event()
        self.last_output: Path | None = None
        self._input_meta = None
        self.ui_font = choose_korean_ui_font(root)

        self.setup_style()
        self.build_ui()

    def setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # - 전체 배경: 아주 연한 블러시 화이트
        # - 카드/입력 영역: 따뜻한 화이트
        # - 타이틀/시작 버튼: 더 깊고 선명한 딥 로즈핑크
        # - 일반 버튼: 거의 흰색에 가까운 아주 연한 베이비 핑크
        bg = "#FFF7FA"
        card = "#FFFBFC"
        field = "#FFFFFF"
        text = "#2D2430"
        subtext = "#655861"
        header = "#5F1835"
        header_sub = "#FFF7FA"
        pink_light = "#FFEAF0"
        pink_light_hover = "#FFE0E9"
        pink_mid = "#C77994"
        pink_dark = "#B83F68"
        pink_dark_hover = "#A9345B"
        border = "#F3CCD8"

        self.palette = {
            "bg": bg,
            "card": card,
            "field": field,
            "text": text,
            "subtext": subtext,
            "header": header,
            "pink_light": pink_light,
            "pink_light_hover": pink_light_hover,
            "pink_mid": pink_mid,
            "pink_dark": pink_dark,
            "pink_dark_hover": pink_dark_hover,
            "border": border,
        }

        self.root.configure(bg=bg)
        style.configure("Root.TFrame", background=bg)
        style.configure("Card.TFrame", background=card)
        style.configure("Header.TFrame", background=header)
        style.configure("HeaderTitle.TLabel", background=header, foreground="#FFFFFF", font=(self.ui_font, 20, "bold"))
        style.configure("HeaderSub.TLabel", background=header, foreground=header_sub, font=(self.ui_font, 10, "bold"))
        style.configure("Title.TLabel", background=card, foreground=text, font=(self.ui_font, 10, "bold"))
        style.configure("Body.TLabel", background=card, foreground=subtext, font=(self.ui_font, 9))
        style.configure("Helper.TLabel", background=card, foreground=text, font=(self.ui_font, 10, "bold"))
        style.configure("ProgressInfo.TLabel", background=card, foreground=text, font=(self.ui_font, 10, "bold"))
        style.configure("Status.TLabel", background=card, foreground="#A43F66", font=(self.ui_font, 10, "bold"))
        style.configure("TCombobox", padding=6, font=(self.ui_font, 10), fieldbackground=field, background=field, foreground=text)
        style.configure("TEntry", padding=7, font=(self.ui_font, 10), fieldbackground=field, foreground=text, bordercolor=border, lightcolor=border, darkcolor=border)
        style.configure("TCheckbutton", background=card, foreground=text, font=(self.ui_font, 9))
        style.map("TCheckbutton", background=[("active", card)], foreground=[("!disabled", text)])
        style.configure("Helper.TCheckbutton", background=card, foreground=text, font=(self.ui_font, 10, "bold"))
        style.map("Helper.TCheckbutton", background=[("active", card)], foreground=[("!disabled", text)])
        style.configure(
            "Pink.Horizontal.TProgressbar",
            troughcolor="#BFB5B9",
            background=pink_mid,
            lightcolor=pink_mid,
            darkcolor=pink_mid,
            bordercolor="#AFA3A8",
            thickness=14,
        )

    def _make_button(self, parent, text: str, command, *, accent=False, width=118, state="normal"):
        p = self.palette
        if accent:
            return RoundedButton(
                parent,
                text=text,
                command=command,
                width=width,
                height=40,
                radius=11,
                fill=p["pink_dark"],
                hover_fill=p["pink_dark_hover"],
                foreground="#FFFFFF",
                border="#9F3157",
                canvas_bg=p["card"],
                font=(self.ui_font, 10, "bold"),
                state=state,
            )
        return RoundedButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=38,
            radius=10,
            fill=p["pink_light"],
            hover_fill=p["pink_light_hover"],
            foreground=p["text"],
            border=p["border"],
            canvas_bg=p["card"],
            font=(self.ui_font, 10, "bold"),
            state=state,
        )

    def build_ui(self):
        outer = ttk.Frame(self.root, style="Root.TFrame", padding=0)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer, style="Header.TFrame", padding=(26, 19))
        header.pack(fill="x")
        ttk.Label(header, text="자동 지오코더", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="주소 파일을 선택하면 도로명·지번주소를 자동 탐지하고 VWorld 좌표를 추가합니다.",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(5, 0))

        body = ttk.Frame(outer, style="Root.TFrame", padding=(20, 16, 20, 20))
        body.pack(fill="both", expand=True)

        control = ttk.Frame(body, style="Card.TFrame", padding=18)
        control.pack(fill="x")
        control.columnconfigure(1, weight=1)

        ttk.Label(control, text="입력 파일", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(control, textvariable=self.input_path, state="readonly").grid(row=0, column=1, sticky="ew", pady=6)
        self.file_btn = self._make_button(control, "파일 선택", self.select_input, width=108)
        self.file_btn.grid(row=0, column=2, padx=(10, 0), pady=6)

        ttk.Label(control, text="출력 파일", style="Title.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(control, textvariable=self.output_path, state="readonly").grid(row=1, column=1, sticky="ew", pady=6)
        self.output_btn = self._make_button(control, "저장 위치", self.select_output, width=108)
        self.output_btn.grid(row=1, column=2, padx=(10, 0), pady=6)

        region = ttk.Frame(control, style="Card.TFrame")
        region.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        ttk.Label(region, text="시도", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.sido_entry = ttk.Entry(region, textvariable=self.sido_var, width=20)
        self.sido_entry.grid(row=0, column=1, padx=(8, 22), sticky="w")
        ttk.Label(region, text="시군구", style="Title.TLabel").grid(row=0, column=2, sticky="w")
        self.sigungu_entry = ttk.Entry(region, textvariable=self.sigungu_var, width=20)
        self.sigungu_entry.grid(row=0, column=3, padx=(8, 18), sticky="w")
        ttk.Label(region, text="입력 파일의 시도·시군구 컬럼을 우선 사용하며, 빈 셀만 이 값을 사용합니다.", style="Helper.TLabel").grid(row=0, column=4, sticky="w")

        options = ttk.Frame(control, style="Card.TFrame")
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        options.columnconfigure(5, weight=1)
        ttk.Label(options, text="출력 좌표계", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.crs_combo = ttk.Combobox(options, textvariable=self.crs_var, values=COMMON_CRS, width=16)
        self.crs_combo.grid(row=0, column=1, padx=(8, 22), sticky="w")
        ttk.Label(options, text="검증 허용거리(m)", style="Title.TLabel").grid(row=0, column=2, sticky="w")
        self.validation_entry = ttk.Entry(options, textvariable=self.validation_var, width=9)
        self.validation_entry.grid(row=0, column=3, padx=(8, 22), sticky="w")
        self.ssl_check = ttk.Checkbutton(options, text="SSL 인증서 검증", variable=self.ssl_var, style="Helper.TCheckbutton")
        self.ssl_check.grid(row=0, column=4, sticky="w")

        ttk.Label(control, textvariable=self.detected_var, style="Helper.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 4))

        actions = ttk.Frame(control, style="Card.TFrame")
        actions.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        actions.columnconfigure(0, weight=1)
        left_actions = ttk.Frame(actions, style="Card.TFrame")
        left_actions.grid(row=0, column=0, sticky="w")
        self.start_btn = self._make_button(left_actions, "지오코딩 시작", self.start_processing, accent=True, width=132)
        self.start_btn.pack(side="left")
        self.cancel_btn = self._make_button(left_actions, "중지", self.cancel_processing, width=82, state="disabled")
        self.cancel_btn.pack(side="left", padx=(8, 0))
        self.open_btn = self._make_button(left_actions, "결과 폴더 열기", self.open_output_folder, width=128, state="disabled")
        self.open_btn.pack(side="left", padx=(8, 0))
        ttk.Label(actions, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=1, sticky="e")

        terminal_card = ttk.Frame(body, style="Card.TFrame", padding=14)
        terminal_card.pack(fill="both", expand=True, pady=(14, 0))
        terminal_header = ttk.Frame(terminal_card, style="Card.TFrame")
        terminal_header.pack(fill="x")
        ttk.Label(terminal_header, text="작업 상황", style="Title.TLabel").pack(side="left")
        ttk.Label(terminal_header, textvariable=self.progress_text_var, style="ProgressInfo.TLabel").pack(side="right")
        self.progress = ttk.Progressbar(terminal_card, style="Pink.Horizontal.TProgressbar", mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(10, 8))

        # CMD 영역은 사용자가 기존 디자인을 선호하여 그대로 유지
        self.log_box = ScrolledText(
            terminal_card,
            height=19,
            bg="#101820",
            fg="#F1E9EC",
            insertbackground="white",
            selectbackground="#7A5060",
            relief="flat",
            font=("Consolas", 9),
            padx=10,
            pady=10,
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

        self.log(f"{APP_NAME} v{APP_VERSION} 준비 완료")
        self.log(f"UI 폰트: {self.ui_font}")
        self.log("지원 형식: XLSX / XLS / CSV / TSV → 결과는 XLSX로 저장")
        self.log("중지 후 같은 입력/출력 파일로 다시 시작하면 자동으로 이어서 처리합니다.")

    def log(self, message: str):
        def _append():
            now = datetime.now().strftime("%H:%M:%S")
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"[{now}] {message}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, _append)

    def set_tqdm_text(self, text: str):
        self.root.after(0, lambda: self.progress_text_var.set(text))

    def set_progress(self, current: int, total: int):
        pct = (current / total * 100) if total else 0
        self.root.after(0, lambda: self.progress.configure(value=pct))

    def select_input(self):
        path = filedialog.askopenfilename(
            title="주소 파일 선택",
            filetypes=[("지원 파일", "*.xlsx *.xls *.csv *.tsv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("TSV", "*.tsv"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        input_path = Path(path)
        self.input_path.set(str(input_path))
        self.output_path.set(str(self.suggest_output_path(input_path)))
        self._inspect_input(input_path)

    def suggest_output_path(self, input_path: Path) -> Path:
        candidate = input_path.with_name(f"{input_path.stem}_좌표추가.xlsx")
        # 미완료 체크포인트가 있으면 반드시 같은 결과 파일을 제안하여 자동 이어하기
        if resume_path_for(candidate).exists():
            return candidate
        if not candidate.exists():
            return candidate
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return input_path.with_name(f"{input_path.stem}_좌표추가_{stamp}.xlsx")

    def select_output(self):
        initial = self.output_path.get() or str(Path.home() / "Desktop" / "좌표추가.xlsx")
        path = filedialog.asksaveasfilename(
            title="결과 파일 저장 위치", initialfile=Path(initial).name, initialdir=str(Path(initial).parent),
            defaultextension=".xlsx", filetypes=[("Excel 통합문서", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _inspect_input(self, path: Path):
        try:
            df, sheet, road_col, parcel_col, extra = load_input_file(path)
            self._input_meta = (sheet, road_col, parcel_col)
            found = []
            if road_col:
                found.append(f"도로명주소 → [{road_col}]")
            if parcel_col:
                found.append(f"지번주소 → [{parcel_col}]")
            sido_col = _pick_region_column(df.columns, "시도")
            sigungu_col = _pick_region_column(df.columns, "시군구")
            place_col = _pick_place_column(df.columns)
            if sido_col:
                found.append(f"시도 → [{sido_col}]")
            if sigungu_col:
                found.append(f"시군구 → [{sigungu_col}]")
            if place_col:
                found.append(f"장소명 → [{place_col}]")
            if road_col or parcel_col:
                self.detected_var.set(f"자동 탐지: {' / '.join(found)} | {len(df):,}건 | {extra}")
                self.log(f"입력 파일 확인: {path.name}")
                self.log(f"컬럼 탐지: {' / '.join(found)}")
                suggested = Path(self.output_path.get()) if self.output_path.get() else None
                if suggested and resume_path_for(suggested).exists():
                    self.log("이전 미완료 작업을 발견했습니다. 시작하면 저장 지점 다음부터 자동으로 이어집니다.")
            else:
                self.detected_var.set("주소 컬럼을 찾지 못했습니다. 컬럼명에 '도로명주소' 또는 '지번주소'가 포함되어야 합니다.")
                self.log("주소 컬럼 탐지 실패")
        except Exception as e:
            self._input_meta = None
            self.detected_var.set("파일을 읽지 못했습니다.")
            self.log(f"입력 파일 확인 오류: {e}")
            messagebox.showerror(APP_NAME, f"파일을 확인하는 중 오류가 발생했습니다.\n\n{e}")

    def start_processing(self):
        if self.is_running:
            return
        input_text = self.input_path.get().strip()
        output_text = self.output_path.get().strip()
        if not input_text:
            messagebox.showwarning(APP_NAME, "먼저 입력 파일을 선택해 주세요.")
            return
        if not output_text:
            messagebox.showwarning(APP_NAME, "출력 파일 위치를 지정해 주세요.")
            return

        api_key = str(CONFIG.get("api_key", "")).strip() or os.getenv("VWORLD_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror(APP_NAME, "VWorld API Key가 설정되어 있지 않습니다.\n\n프로그램(EXE)과 같은 폴더의 config.json에 api_key를 입력해 주세요.")
            return

        try:
            if Path(input_text).resolve() == Path(output_text).resolve():
                messagebox.showwarning(APP_NAME, "입력 파일과 출력 파일은 서로 다른 경로로 지정해 주세요.")
                return
        except Exception:
            pass

        crs = self.crs_var.get().strip().upper()
        if not crs.startswith("EPSG:"):
            messagebox.showwarning(APP_NAME, "좌표계는 'EPSG:4326'과 같은 형식으로 입력해 주세요.")
            return
        try:
            threshold = float(self.validation_var.get())
            if threshold < 0:
                raise ValueError
            if crs != API_CRS:
                Transformer.from_crs(API_CRS, crs, always_xy=True)
        except Exception:
            messagebox.showerror(APP_NAME, "좌표계 또는 검증 허용거리 값이 올바르지 않습니다.")
            return

        sido = normalize_space(self.sido_var.get())
        sigungu = normalize_space(self.sigungu_var.get())

        self.cancel_event.clear()
        self.is_running = True
        self._set_running_ui(True)
        self.progress.configure(value=0)
        self.status_var.set("처리 중")
        self.progress_text_var.set("지오코딩 준비 중...")

        worker = threading.Thread(
            target=self._process_worker,
            args=(Path(input_text), Path(output_text), crs, threshold, api_key, bool(self.ssl_var.get()), sido, sigungu),
            daemon=True,
        )
        worker.start()

    def cancel_processing(self):
        if self.is_running:
            self.cancel_event.set()
            self.log("중지 요청을 받았습니다. 현재 처리 중인 주소까지 마친 뒤 저장합니다.")
            self.status_var.set("중지 요청")

    def _set_running_ui(self, running: bool):
        state_normal = "disabled" if running else "normal"
        for widget in (self.file_btn, self.output_btn, self.start_btn):
            widget.configure(state=state_normal)
        self.crs_combo.configure(state="disabled" if running else "normal")
        self.validation_entry.configure(state="disabled" if running else "normal")
        self.sido_entry.configure(state="disabled" if running else "normal")
        self.sigungu_entry.configure(state="disabled" if running else "normal")
        self.ssl_check.configure(state="disabled" if running else "normal")
        self.cancel_btn.configure(state="normal" if running else "disabled")

    @staticmethod
    def _resume_settings(crs: str, threshold_m: float, sido: str, sigungu: str) -> dict:
        return {"crs": crs, "threshold_m": float(threshold_m), "sido": sido, "sigungu": sigungu}

    def _make_resume_state(self, input_path: Path, sheet_name: str, completed: int, total: int, crs: str, threshold_m: float, sido: str, sigungu: str) -> dict:
        return {
            "app": APP_NAME,
            "version": APP_VERSION,
            "source": source_signature(input_path),
            "sheet_name": sheet_name,
            "completed": int(completed),
            "total": int(total),
            "settings": self._resume_settings(crs, threshold_m, sido, sigungu),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }

    def _load_resume_dataframe(self, input_path: Path, output_path: Path, base_df: pd.DataFrame, sheet_name: str, crs: str, threshold_m: float, sido: str, sigungu: str) -> tuple[pd.DataFrame | None, int]:
        state = load_resume_state(output_path)
        if not state or not output_path.exists():
            return None, 0
        try:
            if state.get("source") != source_signature(input_path):
                self.log("이전 복구정보가 현재 입력 파일과 다르므로 새 작업으로 시작합니다.")
                return None, 0
            if state.get("settings") != self._resume_settings(crs, threshold_m, sido, sigungu):
                self.log("이전 작업과 좌표계/지역/검증거리 설정이 달라 새 작업으로 시작합니다.")
                return None, 0
            resumed = pd.read_excel(output_path, sheet_name=str(sheet_name)[:31], engine="openpyxl")
            if len(resumed) != len(base_df):
                self.log("이전 결과 파일의 행 수가 달라 새 작업으로 시작합니다.")
                return None, 0
            base_columns = [c for c in base_df.columns if c not in RESULT_COLUMNS]
            if any(c not in resumed.columns for c in base_columns):
                self.log("이전 결과 파일의 컬럼 구조가 달라 새 작업으로 시작합니다.")
                return None, 0
            completed = contiguous_processed_count(resumed)
            if completed <= 0 or completed >= len(resumed):
                return None, 0
            return resumed, completed
        except Exception as e:
            self.log(f"이전 작업 복구 확인 실패: {e} | 새 작업으로 시작합니다.")
            return None, 0

    @staticmethod
    def _stats_from_dataframe(df: pd.DataFrame, completed: int) -> dict:
        stats = {"성공": 0, "검증완료": 0, "검증불일치": 0, "부분검증": 0, "검증대상아님": 0, "실패": 0}
        if completed <= 0:
            return stats
        part = df.iloc[:completed]
        for status in ("검증완료", "검증불일치", "부분검증", "검증대상아님", "실패"):
            stats[status] = int((part["검증여부"] == status).sum())
        stats["성공"] = int((part["X좌표"].notna() & part["Y좌표"].notna()).sum())
        return stats

    def _process_worker(self, input_path: Path, output_path: Path, crs: str, threshold_m: float, api_key: str, ssl_verify: bool, sido: str, sigungu: str):
        try:
            base_df, sheet_name, road_col, parcel_col, extra = load_input_file(input_path)
            if not road_col and not parcel_col:
                raise ValueError("주소 컬럼을 찾을 수 없습니다. 컬럼명에 '도로명주소' 또는 '지번주소'가 포함되어야 합니다.")

            resumed_df, completed = self._load_resume_dataframe(input_path, output_path, base_df, sheet_name, crs, threshold_m, sido, sigungu)
            if resumed_df is not None:
                df = resumed_df
                self.log(f"이전 작업 이어하기 | {completed:,}/{len(df):,}건까지 저장된 결과를 확인했습니다.")
            else:
                df = base_df.copy()
                for col in RESULT_COLUMNS:
                    if col in df.columns:
                        df = df.drop(columns=[col])
                for col in RESULT_COLUMNS:
                    df[col] = None if col in {"X좌표", "Y좌표"} else ""
                completed = 0

            total = len(df)
            save_every = max(1, int(CONFIG.get("save_every", 100)))
            geocoder = VWorldGeocoder(
                api_key=api_key, ssl_verify=ssl_verify,
                timeout=int(CONFIG.get("request_timeout", 15)),
                max_retry=int(CONFIG.get("max_retry", 3)),
                request_interval=float(CONFIG.get("request_interval", 0.05)),
            )
            transformer = None if crs == API_CRS else Transformer.from_crs(API_CRS, crs, always_xy=True)

            self.log("=" * 68)
            self.log(f"지오코딩 시작 | {input_path.name}")
            self.log(f"대상 건수: {total:,}건 | 출력 좌표계: {crs} | 자동저장: {save_every:,}건")
            sido_col = _pick_region_column(df.columns, "시도")
            sigungu_col = _pick_region_column(df.columns, "시군구")
            place_col = _pick_place_column(df.columns)
            self.log(f"주소 컬럼: ROAD={road_col or '-'} / PARCEL={parcel_col or '-'}")
            self.log(f"지역 컬럼: 시도={sido_col or '-'} / 시군구={sigungu_col or '-'} / 장소명={place_col or '-'}")
            self.log(f"화면 기본 지역값(빈 셀 fallback): 시도={sido or '(없음)'} / 시군구={sigungu or '(없음)'}")
            self.log(f"교차검증 기준: {threshold_m:g}m")
            if not sido_col and not sigungu_col and not sido and not sigungu:
                self.log("안내: 지역 컬럼과 화면 기본 지역값이 모두 비어 있습니다. '봉산동 296' 같은 축약 주소는 다른 지역으로 매칭될 수 있습니다.")
            if not ssl_verify:
                self.log("주의: SSL 인증서 검증이 꺼져 있습니다.")

            stats = self._stats_from_dataframe(df, completed)
            tqdm_writer = TkTqdmWriter(self.set_tqdm_text)
            pbar = tqdm(
                total=total, initial=completed, desc="지오코딩", unit="건", file=tqdm_writer,
                dynamic_ncols=False, mininterval=0.2,
                bar_format="{l_bar}{bar:28}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]",
            )
            self.set_progress(completed, total)
            processed = completed

            for pos in range(completed, total):
                if self.cancel_event.is_set():
                    break

                idx = df.index[pos]
                row = df.iloc[pos]
                place_label = _short_place_name(row.get(place_col)) if place_col else "장소명없음"
                road_raw = clean_address(row.get(road_col)) if road_col else None
                parcel_raw = clean_address(row.get(parcel_col)) if parcel_col else None

                # 행별 지역 컬럼을 우선 사용하고, 빈 값만 GUI 기본값으로 보완한다.
                row_sido = clean_address(row.get(sido_col)) if sido_col else None
                row_sigungu = clean_address(row.get(sigungu_col)) if sigungu_col else None
                effective_sido = row_sido or sido
                effective_sigungu = row_sigungu or sigungu

                road = apply_region_prefix(road_raw, effective_sido, effective_sigungu)
                parcel = apply_region_prefix(parcel_raw, effective_sido, effective_sigungu)

                x_out = y_out = None
                used = ""
                verified = "실패"
                road_result = geocoder.get_coord(road, "ROAD") if road else None
                parcel_result = geocoder.get_coord(parcel, "PARCEL") if parcel else None
                road_ok = bool(road_result and road_result.get("success"))
                parcel_ok = bool(parcel_result and parcel_result.get("success"))

                if road and parcel:
                    if road_ok and parcel_ok:
                        dist = haversine_m(road_result["x"], road_result["y"], parcel_result["x"], parcel_result["y"])
                        lon, lat = road_result["x"], road_result["y"]  # 도로명주소 우선
                        x_out, y_out = transformer.transform(lon, lat) if transformer else (lon, lat)
                        used = "도로명+지번주소"
                        if dist <= threshold_m:
                            verified = "검증완료"
                            stats["검증완료"] += 1
                        else:
                            verified = "검증불일치"
                            stats["검증불일치"] += 1
                            self.log(f"{place_label} | 행 {pos + 2:,} | 두 좌표 간 약 {dist:.1f}m | ROAD='{road}' | PARCEL='{parcel}'")
                        stats["성공"] += 1
                    elif road_ok:
                        lon, lat = road_result["x"], road_result["y"]
                        x_out, y_out = transformer.transform(lon, lat) if transformer else (lon, lat)
                        used, verified = "도로명주소", "부분검증"
                        stats["성공"] += 1
                        stats["부분검증"] += 1
                    elif parcel_ok:
                        lon, lat = parcel_result["x"], parcel_result["y"]
                        x_out, y_out = transformer.transform(lon, lat) if transformer else (lon, lat)
                        used, verified = "지번주소", "부분검증"
                        stats["성공"] += 1
                        stats["부분검증"] += 1
                    else:
                        stats["실패"] += 1
                elif road:
                    if road_ok:
                        lon, lat = road_result["x"], road_result["y"]
                        x_out, y_out = transformer.transform(lon, lat) if transformer else (lon, lat)
                        used, verified = "도로명주소", "검증대상아님"
                        stats["성공"] += 1
                        stats["검증대상아님"] += 1
                    else:
                        stats["실패"] += 1
                elif parcel:
                    if parcel_ok:
                        lon, lat = parcel_result["x"], parcel_result["y"]
                        x_out, y_out = transformer.transform(lon, lat) if transformer else (lon, lat)
                        used, verified = "지번주소", "검증대상아님"
                        stats["성공"] += 1
                        stats["검증대상아님"] += 1
                    else:
                        stats["실패"] += 1
                else:
                    stats["실패"] += 1

                df.at[idx, "X좌표"] = x_out
                df.at[idx, "Y좌표"] = y_out
                df.at[idx, "사용주소"] = used
                df.at[idx, "검증여부"] = verified

                processed = pos + 1
                pbar.update(1)
                self.set_progress(processed, total)

                if processed % save_every == 0:
                    self.atomic_save(df, output_path, sheet_name)
                    save_resume_state(output_path, self._make_resume_state(input_path, sheet_name, processed, total, crs, threshold_m, sido, sigungu))
                    self.log(f"중간 저장 완료 | {processed:,}/{total:,}건 | 성공 {stats['성공']:,} | 실패 {stats['실패']:,} | 검증완료 {stats['검증완료']:,}")

            pbar.close()
            self.atomic_save(df, output_path, sheet_name)
            self.last_output = output_path

            if self.cancel_event.is_set():
                save_resume_state(output_path, self._make_resume_state(input_path, sheet_name, processed, total, crs, threshold_m, sido, sigungu))
                self.log(f"작업 중지 | {processed:,}/{total:,}건까지 저장 완료 | 다음 시작 시 {processed + 1:,}번째부터 이어집니다.")
                self._finish_ui("중지됨", success=False, cancelled=True)
                return

            remove_resume_state(output_path)
            self.log("-" * 68)
            self.log(f"작업 완료 | 전체 {total:,}건")
            self.log(f"좌표 성공: {stats['성공']:,}건 | 실패: {stats['실패']:,}건")
            self.log(f"검증완료 {stats['검증완료']:,} | 검증불일치 {stats['검증불일치']:,} | 부분검증 {stats['부분검증']:,} | 검증대상아님 {stats['검증대상아님']:,}")
            self.log(f"결과 파일: {output_path}")
            self.log("=" * 68)
            self._finish_ui("완료", success=True)

        except PermissionError:
            self.log("저장 실패: 결과 파일이 Excel에서 열려 있는지 확인해 주세요.")
            self._finish_ui("오류", success=False, error="결과 파일이 열려 있어 저장할 수 없습니다. Excel에서 닫은 뒤 다시 실행해 주세요.")
        except Exception as e:
            self.log(f"오류: {e}")
            self._finish_ui("오류", success=False, error=str(e))

    @staticmethod
    def atomic_save(df: pd.DataFrame, output_path: Path, sheet_name: str):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        safe_sheet = str(sheet_name)[:31] if sheet_name else "수집결과"
        fd, temp_name = tempfile.mkstemp(prefix="autogeocoder_", suffix=".xlsx", dir=str(output_path.parent))
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            df.to_excel(temp_path, sheet_name=safe_sheet, index=False, engine="openpyxl")
            os.replace(temp_path, output_path)
        finally:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass

    def _finish_ui(self, status: str, success: bool, cancelled: bool = False, error: str | None = None):
        def _done():
            self.is_running = False
            self._set_running_ui(False)
            self.status_var.set(status)
            self.cancel_btn.configure(state="disabled")
            if self.last_output and self.last_output.exists():
                self.open_btn.configure(state="normal")
            if success:
                self.progress.configure(value=100)
                messagebox.showinfo(APP_NAME, f"지오코딩이 완료되었습니다.\n\n{self.last_output}")
            elif error:
                messagebox.showerror(APP_NAME, f"작업 중 오류가 발생했습니다.\n\n{error}")
            elif cancelled:
                messagebox.showinfo(APP_NAME, "작업을 중지했습니다. 현재까지의 결과와 이어하기 정보가 저장되었습니다.\n\n다시 '지오코딩 시작'을 누르면 이어서 처리합니다.")
        self.root.after(0, _done)

    def open_output_folder(self):
        if not self.last_output:
            return
        folder = self.last_output.parent
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror(APP_NAME, f"폴더를 열 수 없습니다.\n\n{e}")


def resource_path(name: str) -> Path:
    """PyInstaller onefile/소스 실행 모두에서 번들 리소스 경로를 반환한다."""
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return base / name


def apply_window_icon(root: Tk) -> None:
    """빌드 시 icon.ico가 포함되어 있으면 Tk 창 아이콘에도 적용한다."""
    icon_path = resource_path("icon.ico")
    if not icon_path.exists():
        return
    try:
        root.iconbitmap(default=str(icon_path))
    except Exception:
        pass


def enable_windows_dpi_awareness():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def main():
    enable_windows_dpi_awareness()
    root = Tk()
    apply_window_icon(root)
    AutoGeocoderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
