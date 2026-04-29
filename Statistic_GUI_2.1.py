# -*- coding: utf-8 -*-
"""
异常台风路径统计标记 GUI

说明
1. 本脚本统一按 UTF-8 编写，界面文本、注释、TXT 输出全部避免乱码。
2. 当前默认读取工作区内的两个参考文件，后续你只需要修改下面的“固定路径配置”即可切换正式输入。
3. 本 GUI 聚焦“单条路径、人工协作标记、自动统计、持续存档”的工作流，不包含目标识别算法。
"""

from __future__ import annotations

import math
import os
import re
import shutil
import sqlite3
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.collections import LineCollection
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from tkinter import BOTH, END, HORIZONTAL, LEFT, RIGHT, VERTICAL, X, Y, IntVar, StringVar, Tk
from tkinter import messagebox, ttk

import cartopy.crs as ccrs
import cartopy.feature as cfeature
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "Arial Unicode MS", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False


# ============================================================================
# 固定路径配置
# 下面这几条路径都使用完整路径，后续换正式数据时只改这里即可。
# ============================================================================
WORKSPACE_ROOT = Path(r"D:\code_repository\Statistic_GUI")

# 1) 台风绘图数据 CSV：轨迹点级别输入，当前使用你指定的实验数据。
TRACK_CSV_PATH = Path(
    r"C:\Users\LE\PyCharmMiscProject\岁岁知相思\Kaggle\Stastic-GUI\data实验\49-24年_合并数据集_第5版本_绘图数据-标记-0319H.csv"
)

# 2) 台风名单 / 属性表 XLSX：当前使用你指定的实验数据表。
ATTRIBUTE_XLSX_PATH = Path(
    r"C:\Users\LE\PyCharmMiscProject\岁岁知相思\Kaggle\Stastic-GUI\data实验\49-24年_合并数据集_第5版本_绘图数据-标记-0319H.xlsx"
)

# 3) 输出根目录：所有 PNG / CSV / TXT / XLSX / DB / 日志都在这里。
OUTPUT_ROOT = WORKSPACE_ROOT / "Statistic_GUI_Output"

# 4) 数据库存档路径：用于保存持续工作的进度。
DATABASE_PATH = OUTPUT_ROOT / "db" / "statistic_gui_progress.db"

# 5) 操作历史日志：每次增删段、完成、重生成都会写一条。
HISTORY_LOG_PATH = OUTPUT_ROOT / "db" / "operation_history.txt"


# ============================================================================
# 可调显示 / 统计配置
# ============================================================================
APP_TITLE = "异常台风路径统计标记 GUI v2.0"
MAP_EXTENT = (100.0, 180.0, 0.0, 60.0)
DEFAULT_TIMESTEP_HOURS = 3.0
TRACK_POINT_PICK_RADIUS = 5
MARKED_LINE_PICK_RADIUS = 8
FILTER_SLOT_COUNT = 4
FILTER_DISABLED_LABEL = "不启用"
DEFAULT_FILTER_FIELDS = ["类别", "程度", "PREDICTION", "Ruler"]
FILTER_FIELD_EXCLUDES = {
    "SID",
    "SEASON",
    "NAME",
    "BASIN",
    "SID_YEAR",
    "TOTAL_POINTS",
    "TOTAL_DURATION_HOURS",
    "CATALOG_ORDER",
    "统计",
}
BASE_COLUMNS = ["SID", "SEASON", "NAME", "BASIN", "LAT", "LON", "USA_WIND", "ISO_TIME"]
OUTPUT_ENCODING = "utf-8"
CSV_OUTPUT_ENCODING = "utf-8-sig"
EXPORT_PNG_DPI = 600
BASE_TRACK_LINEWIDTH = 1.9
BASE_TRACK_ALPHA = 0.68
HIGHLIGHT_LINEWIDTH = 4.9
HIGHLIGHT_HALO_WIDTH = 7.1
PREVIEW_LINEWIDTH = 4.8
PREVIEW_HALO_WIDTH = 6.8

PLOT_OUTPUT_DIR = OUTPUT_ROOT / "marked_plots"
SEGMENT_OUTPUT_DIR = OUTPUT_ROOT / "segment_csv"
TEXT_OUTPUT_DIR = OUTPUT_ROOT / "track_txt"
TRACK_XLSX_OUTPUT_DIR = OUTPUT_ROOT / "track_xlsx"
SUMMARY_OUTPUT_DIR = OUTPUT_ROOT / "summary"
MARKED_TRACKS_XLSX_PATH = SUMMARY_OUTPUT_DIR / "marked_tracks_statistics.xlsx"
COMPLETED_TRACKS_XLSX_PATH = SUMMARY_OUTPUT_DIR / "completed_marked_tracks_statistics.xlsx"

SEGMENT_COLOR_POOL = [
    "#d73027",
    "#fc8d59",
    "#fee08b",
    "#91bfdb",
    "#4575b4",
    "#7b3294",
    "#008837",
    "#c51b7d",
    "#4d9221",
    "#f46d43",
]


@dataclass
class SegmentRecord:
    segment_order: int
    start_idx: int
    end_idx: int
    start_time: str
    end_time: str
    point_count: int
    duration_hours: float
    point_ratio: float
    time_ratio: float


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sanitize_filename(text: object) -> str:
    value = str(text or "").strip()
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_") or "UNKNOWN"


def safe_percent(value: Optional[float]) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    return f"{value * 100:.2f}%"


def safe_hours(value: Optional[float]) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    return f"{value:.1f} h"


def safe_int_text(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    try:
        return str(int(value))
    except Exception:
        return str(value)


def extract_year_from_sid(sid: object, fallback: Optional[int] = None) -> Optional[int]:
    sid_str = str(sid or "").strip()
    matches = re.findall(r"(\d{4})", sid_str)
    for item in matches:
        year = int(item)
        if 1850 <= year <= 2100:
            return year
    return fallback


def build_intensity_mapper() -> Tuple[List[float], List[str], List[str]]:
    bins = [0, 34, 64, 83, 96, 113, 137, np.inf]
    labels = ["TD", "TS", "C1", "C2", "C3", "C4", "C5"]
    colors = ["#5ebaff", "#00cc00", "#ffff00", "#ffcc00", "#ff0000", "#ff00ff", "#cc00cc"]
    return bins, labels, colors


def get_intensity_colors(wind_series: Sequence[float], bins: Sequence[float], colors: Sequence[str]) -> List[str]:
    arr = np.array(pd.to_numeric(pd.Series(wind_series), errors="coerce"))
    arr = np.where(np.isnan(arr), -1, arr)
    idx = np.digitize(arr, bins, right=False) - 1
    idx = np.where(arr < 0, 0, idx)
    idx = np.clip(idx, 0, len(colors) - 1)
    return [colors[int(i)] for i in idx]


def read_csv_utf8(path: Path) -> pd.DataFrame:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8"):
        try:
            return pd.read_csv(path, dtype=str, low_memory=False, encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc

    # 最后仍按 UTF-8 处理，但将坏字节替换掉，避免 GUI 在加载阶段直接崩掉。
    try:
        return pd.read_csv(path, dtype=str, low_memory=False, encoding="utf-8", encoding_errors="replace")
    except Exception as exc:
        if last_error is not None:
            raise ValueError(f"CSV 读取失败，UTF-8 解码异常：{last_error}") from exc
        raise


def normalize_track_dataframe(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"找不到轨迹 CSV：{path}")

    df = read_csv_utf8(path)
    missing_columns = [col for col in BASE_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"轨迹 CSV 缺少必要字段：{missing_columns}")

    df = df.copy()
    for col in ["SID", "NAME", "BASIN"]:
        df[col] = df[col].astype(str).str.strip()
    df["SEASON"] = pd.to_numeric(df["SEASON"], errors="coerce").astype("Int64")
    df["LAT"] = pd.to_numeric(df["LAT"], errors="coerce")
    df["LON"] = pd.to_numeric(df["LON"], errors="coerce")
    df["USA_WIND"] = pd.to_numeric(df["USA_WIND"], errors="coerce")
    df["ISO_TIME"] = pd.to_datetime(df["ISO_TIME"], errors="coerce")
    df["SID_YEAR"] = df["SID"].apply(lambda x: extract_year_from_sid(x, None))
    df["SID_YEAR"] = df["SID_YEAR"].fillna(df["SEASON"]).astype("Int64")

    df = df.dropna(subset=["SID", "LAT", "LON"]).copy()
    df = df[(df["LAT"] >= -90) & (df["LAT"] <= 90) & (df["LON"] >= -180) & (df["LON"] <= 180)].copy()
    df["ROW_ORDER"] = np.arange(len(df), dtype=int)
    df = df.sort_values(by=["SID", "ISO_TIME", "ROW_ORDER"], kind="stable").reset_index(drop=True)
    return df


def load_attribute_dataframe(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=["SID"])
    df = pd.read_excel(path)
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    if "SID" not in df.columns:
        raise ValueError(f"属性 XLSX 缺少 SID 字段：{path}")

    # 统一出 GUI 里优先使用的筛选别名，避免 Excel 列名变化时主逻辑跟着改。
    alias_candidates = {
        "程度": ["程度", "程度-旧"],
        "PREDICTION": ["PREDICTION", "PREDICTION-过去-5", "PREDICTION-过去-4"],
        "Ruler": ["Ruler", "判据", "准则判断-0317版本"],
    }
    for alias, candidates in alias_candidates.items():
        if alias in df.columns:
            continue
        for candidate in candidates:
            if candidate in df.columns:
                df[alias] = df[candidate]
                break

    df["SID"] = df["SID"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["SID"], keep="first").reset_index(drop=True)
    df["CATALOG_ORDER"] = np.arange(1, len(df) + 1, dtype=int)
    return df


def compute_duration_hours(times: pd.Series, point_count: int) -> float:
    valid = pd.to_datetime(times, errors="coerce").dropna()
    if len(valid) >= 2:
        diffs = valid.diff().dt.total_seconds().div(3600.0).fillna(0.0)
        positive = diffs[diffs > 0]
        total = float(positive.sum())
        if total > 0:
            return total
    return max(point_count - 1, 0) * DEFAULT_TIMESTEP_HOURS


def build_track_catalog(track_df: pd.DataFrame, attr_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for sid, group in track_df.groupby("SID", sort=False):
        group = group.sort_values(by=["ISO_TIME", "ROW_ORDER"], kind="stable").reset_index(drop=True)
        season = group["SEASON"].dropna().iloc[0] if group["SEASON"].dropna().shape[0] else pd.NA
        name = str(group["NAME"].iloc[0]).strip() if "NAME" in group.columns else sid
        basin = str(group["BASIN"].iloc[0]).strip() if "BASIN" in group.columns else "WP"
        total_points = int(len(group))
        total_hours = compute_duration_hours(group["ISO_TIME"], total_points)
        rows.append(
            {
                "SID": sid,
                "SEASON": season,
                "NAME": name,
                "BASIN": basin,
                "TOTAL_POINTS": total_points,
                "TOTAL_DURATION_HOURS": total_hours,
            }
        )
    catalog = pd.DataFrame(rows)

    if attr_df is not None and not attr_df.empty:
        # 只合并轨迹目录里不存在的附加字段，避免 SEASON / NAME / BASIN 等基础字段重名冲突。
        extra_cols = [col for col in attr_df.columns if col != "SID" and col not in catalog.columns]
        merged = catalog.merge(attr_df[["SID"] + extra_cols], on="SID", how="left")
        if "CATALOG_ORDER" in merged.columns:
            max_order = int(pd.to_numeric(merged["CATALOG_ORDER"], errors="coerce").max() or 0)
            fallback_orders = np.arange(max_order + 1, max_order + 1 + len(merged))
            merged["CATALOG_ORDER"] = pd.to_numeric(merged["CATALOG_ORDER"], errors="coerce")
            merged["CATALOG_ORDER"] = merged["CATALOG_ORDER"].fillna(pd.Series(fallback_orders, index=merged.index))
        else:
            merged["CATALOG_ORDER"] = np.arange(1, len(merged) + 1)
        catalog = merged
    else:
        catalog["CATALOG_ORDER"] = np.arange(1, len(catalog) + 1)

    catalog["SEASON"] = pd.to_numeric(catalog["SEASON"], errors="coerce").astype("Int64")
    catalog = catalog.sort_values(by=["CATALOG_ORDER", "SID"], kind="stable").reset_index(drop=True)
    return catalog


def format_timestamp(value: object) -> str:
    if pd.isna(value):
        return "-"
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def build_segment_metrics(track_df: pd.DataFrame, start_idx: int, end_idx: int) -> SegmentRecord:
    s_idx = int(min(start_idx, end_idx))
    e_idx = int(max(start_idx, end_idx))
    segment_df = track_df.iloc[s_idx : e_idx + 1].copy()
    point_count = int(len(segment_df))
    total_points = int(len(track_df))
    total_hours = float(track_df["TRACK_TOTAL_HOURS"].iloc[0])
    duration_hours = compute_duration_hours(segment_df["ISO_TIME"], point_count)
    point_ratio = (point_count / total_points) if total_points else 0.0
    time_ratio = (duration_hours / total_hours) if total_hours > 0 else point_ratio
    start_time = format_timestamp(segment_df["ISO_TIME"].iloc[0])
    end_time = format_timestamp(segment_df["ISO_TIME"].iloc[-1])
    return SegmentRecord(
        segment_order=0,
        start_idx=s_idx,
        end_idx=e_idx,
        start_time=start_time,
        end_time=end_time,
        point_count=point_count,
        duration_hours=duration_hours,
        point_ratio=point_ratio,
        time_ratio=time_ratio,
    )


def segment_color(order: int) -> str:
    return SEGMENT_COLOR_POOL[(max(order, 1) - 1) % len(SEGMENT_COLOR_POOL)]


def overlap_exists(segments: Sequence[SegmentRecord], start_idx: int, end_idx: int) -> bool:
    s_idx = min(start_idx, end_idx)
    e_idx = max(start_idx, end_idx)
    for seg in segments:
        if not (e_idx < seg.start_idx or s_idx > seg.end_idx):
            return True
    return False


def point_in_segment(index: int, segments: Sequence[SegmentRecord]) -> Optional[SegmentRecord]:
    for seg in segments:
        if seg.start_idx <= index <= seg.end_idx:
            return seg
    return None


def build_output_stub(track_meta: pd.Series) -> str:
    sid = sanitize_filename(track_meta.get("SID"))
    name = sanitize_filename(track_meta.get("NAME"))
    season = sanitize_filename(track_meta.get("SEASON"))
    return f"{sid}_{name}_{season}"


def ensure_output_layout() -> None:
    for folder in [
        OUTPUT_ROOT,
        OUTPUT_ROOT / "db",
        PLOT_OUTPUT_DIR,
        SEGMENT_OUTPUT_DIR,
        TEXT_OUTPUT_DIR,
        TRACK_XLSX_OUTPUT_DIR,
        SUMMARY_OUTPUT_DIR,
    ]:
        ensure_dir(folder)


class ProgressRepository:
    def __init__(self, db_path: Path, log_path: Path):
        ensure_output_layout()
        self.db_path = db_path
        self.log_path = log_path
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.ensure_schema()

    def close(self) -> None:
        try:
            self.conn.close()
        except Exception:
            pass

    def ensure_schema(self) -> None:
        with self.conn:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS tracks (
                    sid TEXT PRIMARY KEY,
                    catalog_order INTEGER,
                    season INTEGER,
                    name TEXT,
                    basin TEXT,
                    total_points INTEGER,
                    total_duration_hours REAL,
                    completed INTEGER DEFAULT 0,
                    last_edit_at TEXT,
                    segment_count INTEGER DEFAULT 0,
                    total_anomaly_ratio REAL DEFAULT 0,
                    max_segment_ratio REAL DEFAULT 0
                )
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS segments (
                    sid TEXT NOT NULL,
                    segment_order INTEGER NOT NULL,
                    start_idx INTEGER NOT NULL,
                    end_idx INTEGER NOT NULL,
                    start_time TEXT,
                    end_time TEXT,
                    point_count INTEGER,
                    duration_hours REAL,
                    point_ratio REAL,
                    time_ratio REAL,
                    created_at TEXT,
                    PRIMARY KEY (sid, segment_order)
                )
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS operations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sid TEXT,
                    action TEXT,
                    detail TEXT,
                    created_at TEXT
                )
                """
            )

    def sync_tracks(self, catalog_df: pd.DataFrame) -> None:
        with self.conn:
            for _, row in catalog_df.iterrows():
                self.conn.execute(
                    """
                    INSERT INTO tracks (
                        sid, catalog_order, season, name, basin,
                        total_points, total_duration_hours,
                        completed, last_edit_at, segment_count,
                        total_anomaly_ratio, max_segment_ratio
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, NULL, 0, 0, 0)
                    ON CONFLICT(sid) DO UPDATE SET
                        catalog_order = excluded.catalog_order,
                        season = excluded.season,
                        name = excluded.name,
                        basin = excluded.basin,
                        total_points = excluded.total_points,
                        total_duration_hours = excluded.total_duration_hours
                    """,
                    (
                        str(row["SID"]),
                        int(row["CATALOG_ORDER"]),
                        None if pd.isna(row["SEASON"]) else int(row["SEASON"]),
                        str(row["NAME"]),
                        str(row["BASIN"]),
                        int(row["TOTAL_POINTS"]),
                        float(row["TOTAL_DURATION_HOURS"]),
                    ),
                )

    def get_track_state(self, sid: str) -> Dict[str, object]:
        row = self.conn.execute("SELECT * FROM tracks WHERE sid = ?", (sid,)).fetchone()
        return dict(row) if row else {}

    def get_all_track_states(self) -> pd.DataFrame:
        rows = self.conn.execute(
            """
            SELECT sid, catalog_order, season, name, basin, total_points,
                   total_duration_hours, completed, last_edit_at, segment_count,
                   total_anomaly_ratio, max_segment_ratio
            FROM tracks
            ORDER BY catalog_order, sid
            """
        ).fetchall()
        return pd.DataFrame([dict(row) for row in rows])

    def get_all_segments(self) -> pd.DataFrame:
        rows = self.conn.execute(
            """
            SELECT sid, segment_order, start_idx, end_idx, start_time, end_time,
                   point_count, duration_hours, point_ratio, time_ratio
            FROM segments
            ORDER BY sid, segment_order
            """
        ).fetchall()
        return pd.DataFrame([dict(row) for row in rows])

    def load_segments(self, sid: str) -> List[SegmentRecord]:
        rows = self.conn.execute(
            """
            SELECT segment_order, start_idx, end_idx, start_time, end_time,
                   point_count, duration_hours, point_ratio, time_ratio
            FROM segments
            WHERE sid = ?
            ORDER BY segment_order
            """,
            (sid,),
        ).fetchall()
        segments = []
        for row in rows:
            segments.append(
                SegmentRecord(
                    segment_order=int(row["segment_order"]),
                    start_idx=int(row["start_idx"]),
                    end_idx=int(row["end_idx"]),
                    start_time=str(row["start_time"]),
                    end_time=str(row["end_time"]),
                    point_count=int(row["point_count"]),
                    duration_hours=float(row["duration_hours"]),
                    point_ratio=float(row["point_ratio"]),
                    time_ratio=float(row["time_ratio"]),
                )
            )
        return segments

    def save_track_segments(
        self,
        track_meta: pd.Series,
        segments: Sequence[SegmentRecord],
        completed: bool,
        action: str,
        detail: str,
    ) -> None:
        sid = str(track_meta["SID"])
        total_ratio = float(sum(seg.time_ratio for seg in segments))
        max_ratio = float(max([seg.time_ratio for seg in segments], default=0.0))
        last_edit = now_text()
        with self.conn:
            self.conn.execute("DELETE FROM segments WHERE sid = ?", (sid,))
            for order, seg in enumerate(sorted(segments, key=lambda item: (item.start_idx, item.end_idx)), start=1):
                self.conn.execute(
                    """
                    INSERT INTO segments (
                        sid, segment_order, start_idx, end_idx, start_time, end_time,
                        point_count, duration_hours, point_ratio, time_ratio, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        sid,
                        int(order),
                        int(seg.start_idx),
                        int(seg.end_idx),
                        str(seg.start_time),
                        str(seg.end_time),
                        int(seg.point_count),
                        float(seg.duration_hours),
                        float(seg.point_ratio),
                        float(seg.time_ratio),
                        last_edit,
                    ),
                )
            self.conn.execute(
                """
                UPDATE tracks
                SET completed = ?, last_edit_at = ?, segment_count = ?,
                    total_anomaly_ratio = ?, max_segment_ratio = ?
                WHERE sid = ?
                """,
                (
                    1 if completed else 0,
                    last_edit,
                    len(segments),
                    total_ratio,
                    max_ratio,
                    sid,
                ),
            )
            self.conn.execute(
                """
                INSERT INTO operations (sid, action, detail, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (sid, action, detail, last_edit),
            )
        self.append_history_log(sid, action, detail, last_edit)

    def append_history_log(self, sid: str, action: str, detail: str, created_at: str) -> None:
        ensure_dir(self.log_path.parent)
        with self.log_path.open("a", encoding=OUTPUT_ENCODING) as fp:
            fp.write(f"[{created_at}] SID={sid} | ACTION={action} | DETAIL={detail}\n")


class TyphoonStatisticApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1780x980")

        self.intensity_bins, self.intensity_labels, self.intensity_colors = build_intensity_mapper()

        self.repo: Optional[ProgressRepository] = None
        self.track_df = pd.DataFrame()
        self.attr_df = pd.DataFrame()
        self.catalog_df = pd.DataFrame()
        self.available_filter_columns: List[str] = []
        self.filter_rows: List[Dict[str, object]] = []

        self.filtered_sids: List[str] = []
        self.listbox_sids: List[str] = []

        self.current_sid: Optional[str] = None
        self.current_track_df = pd.DataFrame()
        self.current_track_meta = pd.Series(dtype=object)
        self.current_segments: List[SegmentRecord] = []
        self.current_completed = False
        self.selected_segment_order: Optional[int] = None
        self.selected_point_index: Optional[int] = None
        self.active_start_idx: Optional[int] = None
        self.current_cursor_idx: Optional[int] = None
        self.context_point_index: Optional[int] = None

        self.track_scatter = None
        self.hover_annotation = None
        self.hover_marker = None
        self.hover_point_index: Optional[int] = None
        self.segment_line_artists: List[Line2D] = []
        self.view_extents: Dict[str, Optional[Tuple[float, float, float, float]]] = {"basemap": None, "focus": None}
        self.last_rendered_view_mode = "basemap"
        self.last_rendered_sid: Optional[str] = None
        self.segment_list_collapsed = False

        self.view_mode_var = StringVar(value="basemap")
        self.track_points_visible = True
        self.track_points_toggle_text = StringVar(value="隐藏轨迹点")
        self.season_filter_enabled_var = IntVar(value=0)
        self.completed_only_var = IntVar(value=0)
        self.season_range_text_var = StringVar(value="全部年份")
        self.season_min_year: Optional[int] = None
        self.season_max_year: Optional[int] = None
        self.available_seasons: List[int] = []
        self.season_selected_values: List[str] = []
        self.search_sid_var = StringVar()
        self._timeline_internal_update = False
        self.timeline_var = IntVar(value=0)

        self.status_var = StringVar(value="准备加载固定路径数据...")
        self.timeline_info_var = StringVar(value="时间轴未加载")
        self.current_segment_hint_var = StringVar(value="当前没有正在预览的异常段。")

        self.basic_info_var = StringVar(value="-")
        self.track_stats_var = StringVar(value="-")
        self.point_info_var = StringVar(value="未选中轨迹点")
        self.segment_info_var = StringVar(value="未选中异常段")
        self.db_state_var = StringVar(value="-")

        self._build_ui()
        self._build_context_menu()
        self._load_fixed_input()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=X)
        ttk.Label(top, text=APP_TITLE, font=("Microsoft YaHei", 15, "bold")).pack(side=LEFT)
        ttk.Label(top, textvariable=self.status_var).pack(side=RIGHT)

        container = ttk.Panedwindow(self.root, orient=HORIZONTAL)
        container.pack(fill=BOTH, expand=True, padx=8, pady=(0, 8))

        self.filter_panel = ttk.Frame(container, padding=8)
        self.center_panel = ttk.Frame(container, padding=8)
        self.info_panel = ttk.Frame(container, padding=8)
        container.add(self.filter_panel, weight=18)
        container.add(self.center_panel, weight=57)
        container.add(self.info_panel, weight=25)

        self._build_filter_panel()
        self._build_center_panel()
        self._build_info_panel()

    def _build_filter_panel(self) -> None:
        block = ttk.LabelFrame(self.filter_panel, text="数据筛选")
        block.pack(fill=BOTH, expand=True)
        tkmod = __import__("tkinter")

        controls_wrap = ttk.Frame(block)
        controls_wrap.pack(fill=X, padx=6, pady=(6, 4))
        self.filter_controls_canvas = tkmod.Canvas(controls_wrap, height=320, highlightthickness=0)
        self.filter_controls_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        filter_scroll = ttk.Scrollbar(controls_wrap, orient=VERTICAL, command=self.filter_controls_canvas.yview)
        filter_scroll.pack(side=RIGHT, fill=Y)
        self.filter_controls_canvas.configure(yscrollcommand=filter_scroll.set)
        self.filter_controls_inner = ttk.Frame(self.filter_controls_canvas)
        self.filter_controls_window = self.filter_controls_canvas.create_window(
            (0, 0),
            window=self.filter_controls_inner,
            anchor="nw",
        )
        self.filter_controls_inner.bind("<Configure>", self._on_filter_inner_configure)
        self.filter_controls_canvas.bind("<Configure>", self._on_filter_canvas_configure)
        self.filter_controls_canvas.bind("<MouseWheel>", lambda event: self._scroll_canvas_with_mousewheel(event, self.filter_controls_canvas))
        self.filter_controls_inner.bind("<MouseWheel>", lambda event: self._scroll_canvas_with_mousewheel(event, self.filter_controls_canvas))

        ttk.Label(self.filter_controls_inner, text="SID 搜索").pack(anchor="w", pady=(2, 2))
        search_row = ttk.Frame(self.filter_controls_inner)
        search_row.pack(fill=X)
        sid_entry = ttk.Entry(search_row, textvariable=self.search_sid_var)
        sid_entry.pack(side=LEFT, fill=X, expand=True)
        sid_entry.bind("<Return>", lambda _event: self.search_by_sid())
        ttk.Button(search_row, text="搜索", command=self.search_by_sid).pack(side=LEFT, padx=(6, 0))

        ttk.Label(self.filter_controls_inner, text="SEASON").pack(anchor="w", pady=(12, 2))
        season_row = ttk.Frame(self.filter_controls_inner)
        season_row.pack(fill=X)
        ttk.Checkbutton(
            season_row,
            text="启用",
            variable=self.season_filter_enabled_var,
            command=self.apply_filters,
        ).pack(side=LEFT)
        ttk.Button(season_row, text="选择值", command=self.open_season_value_selector).pack(side=LEFT, padx=(6, 0))
        ttk.Button(season_row, text="区间", command=self.open_season_range_dialog).pack(side=LEFT, padx=(6, 0))
        ttk.Label(self.filter_controls_inner, textvariable=self.season_range_text_var, foreground="#475569").pack(
            anchor="w",
            pady=(4, 0),
        )

        ttk.Checkbutton(
            self.filter_controls_inner,
            text="只看已完成标记",
            variable=self.completed_only_var,
            command=self.apply_filters,
        ).pack(anchor="w", pady=(12, 0))

        ttk.Label(self.filter_controls_inner, text="自定义筛选").pack(anchor="w", pady=(12, 2))
        for slot_index in range(FILTER_SLOT_COUNT):
            row_frame = ttk.Frame(self.filter_controls_inner)
            row_frame.pack(fill=X, pady=(4 if slot_index else 2, 0))

            enabled_var = IntVar(value=1 if slot_index < len(DEFAULT_FILTER_FIELDS) else 0)
            field_var = StringVar(value=FILTER_DISABLED_LABEL)
            display_var = StringVar(value="全部")

            top_row = ttk.Frame(row_frame)
            top_row.pack(fill=X)
            ttk.Checkbutton(
                top_row,
                text="启用",
                variable=enabled_var,
                command=self.apply_filters,
            ).pack(side=LEFT)

            field_combo = ttk.Combobox(top_row, textvariable=field_var, state="readonly", width=12)
            field_combo.pack(side=LEFT, fill=X, expand=True, padx=(6, 0))
            field_combo.bind(
                "<<ComboboxSelected>>",
                lambda _event, idx=slot_index: self._on_filter_field_changed(idx),
            )
            ttk.Button(
                top_row,
                text="选择值",
                command=lambda idx=slot_index: self.open_filter_value_selector(idx),
            ).pack(side=LEFT, padx=(6, 0))

            ttk.Label(row_frame, textvariable=display_var, foreground="#475569").pack(anchor="w", pady=(4, 0))

            self.filter_rows.append(
                {
                    "enabled_var": enabled_var,
                    "field_var": field_var,
                    "display_var": display_var,
                    "field_combo": field_combo,
                    "selected_values": [],
                    "available_values": [],
                }
            )

        nav_row = ttk.Frame(self.filter_controls_inner)
        nav_row.pack(fill=X, pady=(14, 2))
        ttk.Button(nav_row, text="上一条", command=self.go_prev_track).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(nav_row, text="下一条", command=self.go_next_track).pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(nav_row, text="重置筛选", command=self.reset_filters).pack(side=LEFT, fill=X, expand=True)

        ttk.Label(block, text="路径列表").pack(anchor="w", padx=6, pady=(8, 2))
        list_frame = ttk.Frame(block)
        list_frame.pack(fill=BOTH, expand=True, padx=6, pady=(0, 8))
        self.track_listbox = tkmod.Listbox(list_frame, exportselection=False)
        self.track_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.track_listbox.bind("<<ListboxSelect>>", self.on_track_selected)
        scroll = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.track_listbox.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.track_listbox.config(yscrollcommand=scroll.set)

    def _build_center_panel(self) -> None:
        view_bar = ttk.Frame(self.center_panel)
        view_bar.pack(fill=X, pady=(0, 6))
        ttk.Radiobutton(
            view_bar,
            text="底图视图",
            variable=self.view_mode_var,
            value="basemap",
            command=self.render_current_track,
        ).pack(side=LEFT)
        ttk.Radiobutton(
            view_bar,
            text="聚焦放大视图",
            variable=self.view_mode_var,
            value="focus",
            command=self.render_current_track,
        ).pack(side=LEFT, padx=(10, 0))
        ttk.Button(
            view_bar,
            textvariable=self.track_points_toggle_text,
            command=self.toggle_track_points_visibility,
        ).pack(side=LEFT, padx=(12, 0))

        self.figure = Figure(figsize=(10.5, 7.2))
        self.ax = self.figure.add_subplot(111, projection=ccrs.PlateCarree())
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.center_panel)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=True)
        self.canvas.mpl_connect("pick_event", self.on_pick_event)
        self.canvas.mpl_connect("motion_notify_event", self.on_motion_event)
        self.canvas.mpl_connect("scroll_event", self.on_scroll_zoom)

        self.toolbar = NavigationToolbar2Tk(self.canvas, self.center_panel)
        self.toolbar.update()

        time_block = ttk.LabelFrame(self.center_panel, text="时间轴与局部异常段操作")
        time_block.pack(fill=X, pady=(8, 0))
        tkmod = __import__("tkinter")
        self.timeline_scale = tkmod.Scale(
            time_block,
            orient=HORIZONTAL,
            from_=0,
            to=0,
            resolution=1,
            variable=self.timeline_var,
            showvalue=False,
            command=self.on_timeline_changed,
        )
        self.timeline_scale.pack(fill=X, padx=8, pady=(8, 2))
        ttk.Label(time_block, textvariable=self.timeline_info_var).pack(anchor="w", padx=8)
        ttk.Label(time_block, textvariable=self.current_segment_hint_var, foreground="#9a3412").pack(
            anchor="w", padx=8, pady=(2, 6)
        )

        action_row = ttk.Frame(time_block)
        action_row.pack(fill=X, padx=8, pady=(0, 8))
        ttk.Button(action_row, text="将当前点设为起始点", command=self.set_current_point_as_start).pack(side=LEFT)
        ttk.Button(action_row, text="局部标记", command=self.commit_active_segment).pack(side=LEFT, padx=(6, 0))
        ttk.Button(action_row, text="取消/重新选择这一段", command=self.clear_active_segment).pack(side=LEFT, padx=(6, 0))
        ttk.Button(action_row, text="删除选中段", command=self.delete_selected_segment).pack(side=LEFT, padx=(6, 0))

    def _build_info_panel(self) -> None:
        status_block = ttk.LabelFrame(self.info_panel, text="工作状态")
        status_block.pack(fill=X, pady=(0, 8))
        ttk.Label(status_block, textvariable=self.db_state_var, justify=LEFT).pack(anchor="w", padx=8, pady=8)

        basic_block = ttk.LabelFrame(self.info_panel, text="基本信息")
        basic_block.pack(fill=X, pady=(0, 8))
        ttk.Label(basic_block, textvariable=self.basic_info_var, justify=LEFT).pack(anchor="w", padx=8, pady=8)

        stats_block = ttk.LabelFrame(self.info_panel, text="整体统计")
        stats_block.pack(fill=X, pady=(0, 8))
        ttk.Label(stats_block, textvariable=self.track_stats_var, justify=LEFT).pack(anchor="w", padx=8, pady=8)

        point_block = ttk.LabelFrame(self.info_panel, text="选中点信息")
        point_block.pack(fill=X, pady=(0, 8))
        ttk.Label(point_block, textvariable=self.point_info_var, justify=LEFT).pack(anchor="w", padx=8, pady=8)

        segment_block = ttk.LabelFrame(self.info_panel, text="选中异常段信息")
        segment_block.pack(fill=X, pady=(0, 8))
        ttk.Label(segment_block, textvariable=self.segment_info_var, justify=LEFT).pack(anchor="w", padx=8, pady=8)

        segment_list_block = ttk.LabelFrame(self.info_panel, text="已标记异常段")
        segment_list_block.pack(fill=BOTH, expand=True, pady=(0, 8))
        toggle_row = ttk.Frame(segment_list_block)
        toggle_row.pack(fill=X, padx=8, pady=(8, 0))
        self.segment_toggle_button = ttk.Button(
            toggle_row,
            text="收起已标记异常段",
            command=self.toggle_segment_list_panel,
        )
        self.segment_toggle_button.pack(anchor="w")

        frame = ttk.Frame(segment_list_block)
        frame.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.segment_list_container = frame
        tkmod = __import__("tkinter")
        self.segment_listbox = tkmod.Listbox(frame, exportselection=False, height=10)
        self.segment_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.segment_listbox.bind("<<ListboxSelect>>", self.on_segment_list_selected)
        seg_scroll = ttk.Scrollbar(frame, orient=VERTICAL, command=self.segment_listbox.yview)
        seg_scroll.pack(side=RIGHT, fill=Y)
        self.segment_listbox.config(yscrollcommand=seg_scroll.set)

        output_block = ttk.LabelFrame(self.info_panel, text="输出与存档")
        output_block.pack(fill=BOTH, pady=(0, 0))
        tkmod = __import__("tkinter")
        output_canvas_wrap = ttk.Frame(output_block)
        output_canvas_wrap.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.output_canvas = tkmod.Canvas(output_canvas_wrap, height=150, highlightthickness=0)
        self.output_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        output_scroll = ttk.Scrollbar(output_canvas_wrap, orient=VERTICAL, command=self.output_canvas.yview)
        output_scroll.pack(side=RIGHT, fill=Y)
        self.output_canvas.configure(yscrollcommand=output_scroll.set)
        self.output_inner_frame = ttk.Frame(self.output_canvas)
        self.output_canvas_window = self.output_canvas.create_window((0, 0), window=self.output_inner_frame, anchor="nw")
        self.output_inner_frame.bind("<Configure>", self._on_output_inner_configure)
        self.output_canvas.bind("<Configure>", self._on_output_canvas_configure)
        self.output_canvas.bind("<MouseWheel>", lambda event: self._scroll_canvas_with_mousewheel(event, self.output_canvas))
        self.output_inner_frame.bind("<MouseWheel>", lambda event: self._scroll_canvas_with_mousewheel(event, self.output_canvas))

        ttk.Button(self.output_inner_frame, text="保存并标记完成", command=self.mark_current_track_completed).pack(
            fill=X, pady=(0, 4)
        )
        ttk.Button(self.output_inner_frame, text="标记为未完成", command=self.mark_current_track_incomplete).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="重新生成输出", command=self.regenerate_outputs).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="重生成全部已标记 PNG", command=self.regenerate_all_marked_pngs).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="清理空段 CSV 目录", command=self.cleanup_segment_output_dirs).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="打开 PNG 输出文件夹", command=self.open_plot_folder).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="打开段 CSV 文件夹", command=self.open_segment_folder).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="打开 TXT 输出文件夹", command=self.open_text_folder).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="打开单路径 XLSX 文件夹", command=self.open_xlsx_folder).pack(
            fill=X, pady=4
        )
        ttk.Button(self.output_inner_frame, text="打开汇总 XLSX 文件夹", command=self.open_summary_folder).pack(
            fill=X, pady=(4, 0)
        )

    def _build_context_menu(self) -> None:
        tkmod = __import__("tkinter")
        self.context_menu = tkmod.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="将这个轨迹点作为起始点", command=self.context_set_start)
        self.context_menu.add_command(label="将这个轨迹点作为结束点并局部标记", command=self.context_set_end_and_commit)
        self.context_menu.add_command(label="仅将这个轨迹点设为结束点", command=self.context_set_end_only)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="取消-重新选择这一段", command=self.clear_active_segment)
        self.context_menu.add_command(label="选中该点所在已标记段", command=self.context_select_existing_segment)

    def _load_fixed_input(self) -> None:
        try:
            self.track_df = normalize_track_dataframe(TRACK_CSV_PATH)
            self.attr_df = load_attribute_dataframe(ATTRIBUTE_XLSX_PATH)
            self.catalog_df = build_track_catalog(self.track_df, self.attr_df)
            self.repo = ProgressRepository(DATABASE_PATH, HISTORY_LOG_PATH)
            self.repo.sync_tracks(self.catalog_df)
            self._prepare_track_total_hours()
            self._init_filter_options()
            self.cleanup_segment_output_dirs()
            self.apply_filters()
            self.status_var.set(f"已加载 {len(self.catalog_df)} 条路径。")
            if self.filtered_sids:
                self.load_track_by_sid(self.filtered_sids[0])
        except Exception as exc:
            self.status_var.set("固定路径数据加载失败")
            tb_lines = traceback.format_exc().strip().splitlines()
            last_trace = tb_lines[-1] if tb_lines else "-"
            messagebox.showerror(
                "加载失败",
                f"固定路径读取失败：\n{type(exc).__name__}: {exc}\n\n最后一行追踪：\n{last_trace}",
            )

    def _prepare_track_total_hours(self) -> None:
        hours_map = {}
        for _, row in self.catalog_df.iterrows():
            hours_map[str(row["SID"])] = float(row["TOTAL_DURATION_HOURS"])
        self.track_df["TRACK_TOTAL_HOURS"] = self.track_df["SID"].map(hours_map).astype(float)

    def _init_filter_options(self) -> None:
        self.available_seasons = sorted([int(v) for v in pd.Series(self.catalog_df["SEASON"]).dropna().unique()])
        if self.available_seasons:
            self.season_min_year = int(self.available_seasons[0])
            self.season_max_year = int(self.available_seasons[-1])
            self.season_selected_values = [str(year) for year in self.available_seasons]
        else:
            self.season_min_year = None
            self.season_max_year = None
            self.season_selected_values = []
        self._update_season_range_label()
        self.season_filter_enabled_var.set(0)
        self.completed_only_var.set(0)

        self.available_filter_columns = self._get_available_filter_columns()
        default_fields = self._get_default_filter_fields()
        field_choices = [FILTER_DISABLED_LABEL] + self.available_filter_columns

        for idx, row in enumerate(self.filter_rows):
            field_combo = row["field_combo"]
            field_var = row["field_var"]
            field_combo["values"] = field_choices
            field_var.set(default_fields[idx] if idx < len(default_fields) else FILTER_DISABLED_LABEL)
            self._refresh_filter_row(idx, preserve_value=False)

    def _update_season_range_label(self) -> None:
        available_values = [str(year) for year in self.available_seasons]
        selected_values = [value for value in self.season_selected_values if value in available_values]
        self.season_selected_values = selected_values if selected_values else list(available_values)
        if not available_values:
            self.season_range_text_var.set("全部年份")
            return
        self.season_min_year = int(self.season_selected_values[0])
        self.season_max_year = int(self.season_selected_values[-1])
        self.season_range_text_var.set(self._summarize_filter_values(self.season_selected_values, available_values))

    def _get_available_filter_columns(self) -> List[str]:
        available = []
        for col in self.catalog_df.columns:
            if col in FILTER_FIELD_EXCLUDES:
                continue
            series = self.catalog_df[col]
            non_empty = [str(v).strip() for v in series.dropna().tolist() if str(v).strip()]
            if not non_empty:
                continue
            available.append(col)

        preferred = [col for col in DEFAULT_FILTER_FIELDS if col in available]
        remaining = [col for col in available if col not in preferred]
        return preferred + remaining

    def _get_default_filter_fields(self) -> List[str]:
        selected = []
        for col in DEFAULT_FILTER_FIELDS:
            if col in self.available_filter_columns and col not in selected:
                selected.append(col)
        for col in self.available_filter_columns:
            if len(selected) >= FILTER_SLOT_COUNT:
                break
            if col not in selected:
                selected.append(col)
        while len(selected) < FILTER_SLOT_COUNT:
            selected.append(FILTER_DISABLED_LABEL)
        return selected

    def _refresh_filter_row(self, slot_index: int, preserve_value: bool = True) -> None:
        row = self.filter_rows[slot_index]
        field_var = row["field_var"]
        display_var = row["display_var"]

        selected_field = field_var.get().strip()
        previous_values = list(row.get("selected_values", []))

        if not selected_field or selected_field == FILTER_DISABLED_LABEL or selected_field not in self.catalog_df.columns:
            row["available_values"] = []
            row["selected_values"] = []
            display_var.set("全部")
            return

        items = [
            str(v).strip()
            for v in self.catalog_df[selected_field].dropna().tolist()
            if str(v).strip()
        ]
        values = sorted(dict.fromkeys(items).keys())
        row["available_values"] = values
        if preserve_value and previous_values:
            selected_values = [value for value in previous_values if value in values]
            row["selected_values"] = selected_values if selected_values else list(values)
        else:
            row["selected_values"] = list(values)
        display_var.set(self._summarize_filter_values(row["selected_values"], values))

    @staticmethod
    def _summarize_filter_values(selected_values: Sequence[str], available_values: Sequence[str]) -> str:
        if not available_values:
            return "无可选值"
        if not selected_values or len(selected_values) >= len(available_values):
            return "全部"
        if len(selected_values) <= 2:
            return " / ".join(selected_values)
        return f"已选 {len(selected_values)} 项"

    def _on_filter_field_changed(self, slot_index: int) -> None:
        self._refresh_filter_row(slot_index, preserve_value=False)
        self.apply_filters()

    def _open_multi_value_selector_dialog(
        self,
        title: str,
        available_values: Sequence[str],
        selected_values: Sequence[str],
        apply_callback,
    ) -> None:
        if not available_values:
            messagebox.showinfo("无可选值", "当前没有可用于筛选的取值。")
            return

        tkmod = __import__("tkinter")
        dialog = tkmod.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("360x420")

        selected_set = set(selected_values or available_values)
        all_var = IntVar(value=1 if len(selected_set) >= len(available_values) else 0)
        item_vars = []

        outer = ttk.Frame(dialog, padding=10)
        outer.pack(fill=BOTH, expand=True)
        ttk.Checkbutton(outer, text="全选", variable=all_var).pack(anchor="w")

        canvas_wrap = ttk.Frame(outer)
        canvas_wrap.pack(fill=BOTH, expand=True, pady=(8, 0))
        canvas = tkmod.Canvas(canvas_wrap, highlightthickness=0)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(canvas_wrap, orient=VERTICAL, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_inner_configure(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        inner.bind("<Configure>", on_inner_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        for value in available_values:
            var = IntVar(value=1 if value in selected_set else 0)
            ttk.Checkbutton(inner, text=value, variable=var).pack(anchor="w", fill=X)
            item_vars.append((value, var))

        updating_all_flag = {"value": False}

        def toggle_all() -> None:
            if updating_all_flag["value"]:
                return
            checked = int(all_var.get())
            updating_all_flag["value"] = True
            for _value, var in item_vars:
                var.set(checked)
            updating_all_flag["value"] = False

        def sync_all_var() -> None:
            if updating_all_flag["value"]:
                return
            selected_count = sum(int(var.get()) for _value, var in item_vars)
            updating_all_flag["value"] = True
            all_var.set(1 if selected_count == len(item_vars) else 0)
            updating_all_flag["value"] = False

        all_var.trace_add("write", lambda *_args: toggle_all())
        for _value, var in item_vars:
            var.trace_add("write", lambda *_args: sync_all_var())

        def apply_values() -> None:
            new_selected_values = [value for value, var in item_vars if int(var.get()) == 1]
            if not new_selected_values:
                messagebox.showwarning("未选择取值", "至少保留一个筛选值。")
                return
            apply_callback(new_selected_values)
            dialog.destroy()

        action_row = ttk.Frame(outer)
        action_row.pack(fill=X, pady=(10, 0))
        ttk.Button(action_row, text="确定", command=apply_values).pack(side=LEFT)
        ttk.Button(action_row, text="取消", command=dialog.destroy).pack(side=LEFT, padx=(6, 0))

    def open_season_value_selector(self) -> None:
        available_values = [str(year) for year in self.available_seasons]
        if not available_values:
            messagebox.showinfo("无可选年份", "当前数据里没有可用于筛选的 SEASON 年份。")
            return

        def apply_values(selected_values: Sequence[str]) -> None:
            ordered_selected = [value for value in available_values if value in selected_values]
            self.season_selected_values = ordered_selected
            if ordered_selected:
                self.season_min_year = int(ordered_selected[0])
                self.season_max_year = int(ordered_selected[-1])
            self.season_filter_enabled_var.set(1)
            self._update_season_range_label()
            self.apply_filters()

        self._open_multi_value_selector_dialog(
            "筛选值: SEASON",
            available_values,
            self.season_selected_values,
            apply_values,
        )

    def open_season_range_dialog(self) -> None:
        if not self.available_seasons:
            messagebox.showinfo("无可选年份", "当前数据里没有可用于筛选的 SEASON 年份。")
            return

        tkmod = __import__("tkinter")
        dialog = tkmod.Toplevel(self.root)
        dialog.title("设置 SEASON 区间")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        start_var = StringVar(value=str(self.season_min_year or self.available_seasons[0]))
        end_var = StringVar(value=str(self.season_max_year or self.available_seasons[-1]))
        year_values = [str(year) for year in self.available_seasons]

        body = ttk.Frame(dialog, padding=12)
        body.pack(fill=BOTH, expand=True)
        ttk.Label(body, text="起始年份").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="结束年份").grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Combobox(body, textvariable=start_var, values=year_values, state="readonly", width=12).grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(6, 0),
        )
        ttk.Combobox(body, textvariable=end_var, values=year_values, state="readonly", width=12).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(10, 0),
            pady=(6, 0),
        )

        def apply_range() -> None:
            start_year = int(start_var.get())
            end_year = int(end_var.get())
            if start_year > end_year:
                messagebox.showwarning("区间无效", "SEASON 起始年份不能大于结束年份。")
                return
            self.season_min_year = start_year
            self.season_max_year = end_year
            self.season_selected_values = [
                str(year) for year in self.available_seasons if start_year <= int(year) <= end_year
            ]
            self.season_filter_enabled_var.set(1)
            self._update_season_range_label()
            dialog.destroy()
            self.apply_filters()

        action_row = ttk.Frame(body)
        action_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(action_row, text="确定", command=apply_range).pack(side=LEFT)
        ttk.Button(action_row, text="取消", command=dialog.destroy).pack(side=LEFT, padx=(6, 0))

    def open_filter_value_selector(self, slot_index: int) -> None:
        row = self.filter_rows[slot_index]
        field_name = row["field_var"].get().strip()
        available_values = list(row.get("available_values", []))
        if not field_name or field_name == FILTER_DISABLED_LABEL:
            messagebox.showinfo("未选择字段", "请先选择一个自定义筛选字段。")
            return
        if not available_values:
            messagebox.showinfo("无可选值", f"{field_name} 当前没有可用于筛选的取值。")
            return

        def apply_values(selected_values: Sequence[str]) -> None:
            row["selected_values"] = selected_values
            row["display_var"].set(self._summarize_filter_values(selected_values, available_values))
            row["enabled_var"].set(1)
            self.apply_filters()

        self._open_multi_value_selector_dialog(
            f"筛选值: {field_name}",
            available_values,
            row.get("selected_values", []),
            apply_values,
        )

    def apply_filters(self) -> None:
        if self.catalog_df.empty:
            return

        current_sid = self.current_sid
        df = self.catalog_df.copy()

        if int(self.season_filter_enabled_var.get()) == 1:
            available_season_values = [str(year) for year in self.available_seasons]
            selected_season_values = [value for value in self.season_selected_values if value in available_season_values]
            if selected_season_values and len(selected_season_values) < len(available_season_values):
                season_series = pd.to_numeric(df["SEASON"], errors="coerce").astype("Int64").astype(str)
                df = df[season_series.isin(selected_season_values)]

        for row in self.filter_rows:
            enabled = int(row["enabled_var"].get()) == 1
            field_name = row["field_var"].get().strip()
            selected_values = list(row.get("selected_values", []))
            available_values = list(row.get("available_values", []))
            if not enabled:
                continue
            if field_name and field_name != FILTER_DISABLED_LABEL and field_name in df.columns:
                if selected_values and len(selected_values) < len(available_values):
                    df = df[df[field_name].astype(str).str.strip().isin(selected_values)]

        if int(self.completed_only_var.get()) == 1 and self.repo is not None:
            state_df = self.repo.get_all_track_states()
            if not state_df.empty:
                completed_sids = state_df[
                    (state_df["completed"].fillna(0).astype(int) == 1)
                    & (state_df["segment_count"].fillna(0).astype(int) > 0)
                ]["sid"].astype(str)
                df = df[df["SID"].astype(str).isin(completed_sids.tolist())]
            else:
                df = df.iloc[0:0].copy()

        sid_query = self.search_sid_var.get().strip()
        if sid_query:
            sid_query_lower = sid_query.lower()
            mask = df["SID"].astype(str).str.lower().str.contains(sid_query_lower)
            if "NAME" in df.columns:
                mask |= df["NAME"].astype(str).str.lower().str.contains(sid_query_lower)
            df = df[mask]

        df = df.sort_values(by=["CATALOG_ORDER", "SID"], kind="stable").reset_index(drop=True)
        self.filtered_sids = [str(v) for v in df["SID"].tolist()]
        self._refresh_track_listbox(df)

        if current_sid and current_sid in self.filtered_sids:
            self.select_track_in_listbox(current_sid)
        elif self.filtered_sids:
            self.select_track_in_listbox(self.filtered_sids[0])
        else:
            self.track_listbox.selection_clear(0, END)
            self.clear_track_view("当前筛选条件下没有路径。")

    def _refresh_track_listbox(self, df: pd.DataFrame) -> None:
        self.track_listbox.delete(0, END)
        self.listbox_sids = []
        state_df = self.repo.get_all_track_states() if self.repo else pd.DataFrame()
        state_map = {str(row["sid"]): row for _, row in state_df.iterrows()} if not state_df.empty else {}

        for _, row in df.iterrows():
            sid = str(row["SID"])
            state = state_map.get(sid, {})
            completed = int(state.get("completed", 0) or 0)
            seg_count = int(state.get("segment_count", 0) or 0)
            prefix = "✓" if completed else ("●" if seg_count > 0 else "○")
            label = f"{prefix} {sid} | {row['NAME']} | {safe_int_text(row['SEASON'])}"
            self.track_listbox.insert(END, label)
            self.listbox_sids.append(sid)
            idx = self.track_listbox.size() - 1
            if completed:
                self.track_listbox.itemconfig(idx, fg="#166534")
            elif seg_count > 0:
                self.track_listbox.itemconfig(idx, fg="#b45309")

    def search_by_sid(self) -> None:
        self.apply_filters()

    def reset_filters(self) -> None:
        self.search_sid_var.set("")
        self.season_filter_enabled_var.set(0)
        self.completed_only_var.set(0)
        if self.available_seasons:
            self.season_min_year = int(self.available_seasons[0])
            self.season_max_year = int(self.available_seasons[-1])
            self.season_selected_values = [str(year) for year in self.available_seasons]
        else:
            self.season_min_year = None
            self.season_max_year = None
            self.season_selected_values = []
        self._update_season_range_label()
        default_fields = self._get_default_filter_fields()
        for idx, row in enumerate(self.filter_rows):
            row["enabled_var"].set(1 if idx < len(DEFAULT_FILTER_FIELDS) else 0)
            row["field_var"].set(default_fields[idx] if idx < len(default_fields) else FILTER_DISABLED_LABEL)
            self._refresh_filter_row(idx, preserve_value=False)
        self.apply_filters()

    def select_track_in_listbox(self, sid: str) -> None:
        if sid not in self.listbox_sids:
            return
        idx = self.listbox_sids.index(sid)
        self.track_listbox.selection_clear(0, END)
        self.track_listbox.selection_set(idx)
        self.track_listbox.see(idx)
        self.load_track_by_sid(sid)

    def on_track_selected(self, _event=None) -> None:
        selection = self.track_listbox.curselection()
        if not selection:
            return
        sid = self.listbox_sids[selection[0]]
        self.load_track_by_sid(sid)

    def go_prev_track(self) -> None:
        if not self.filtered_sids:
            return
        if self.current_sid not in self.filtered_sids:
            self.load_track_by_sid(self.filtered_sids[0])
            return
        idx = self.filtered_sids.index(self.current_sid)
        target = self.filtered_sids[(idx - 1) % len(self.filtered_sids)]
        self.select_track_in_listbox(target)

    def go_next_track(self) -> None:
        if not self.filtered_sids:
            return
        if self.current_sid not in self.filtered_sids:
            self.load_track_by_sid(self.filtered_sids[0])
            return
        idx = self.filtered_sids.index(self.current_sid)
        target = self.filtered_sids[(idx + 1) % len(self.filtered_sids)]
        self.select_track_in_listbox(target)

    def load_track_by_sid(self, sid: str) -> None:
        if not sid:
            return
        group = self.track_df[self.track_df["SID"] == sid].copy()
        if group.empty:
            messagebox.showwarning("找不到路径", f"没有找到 SID={sid} 的轨迹点。")
            return

        group = group.sort_values(by=["ISO_TIME", "ROW_ORDER"], kind="stable").reset_index(drop=True)
        meta = self.catalog_df[self.catalog_df["SID"] == sid]
        if meta.empty:
            return
        self.current_sid = sid
        self.current_track_df = group
        self.current_track_meta = meta.iloc[0]
        self.current_segments = self.repo.load_segments(sid) if self.repo else []
        state = self.repo.get_track_state(sid) if self.repo else {}
        self.current_completed = bool(int(state.get("completed", 0) or 0))
        self.selected_segment_order = None
        self.selected_point_index = 0 if not group.empty else None
        self.active_start_idx = None
        self.current_cursor_idx = 0 if not group.empty else None
        self.context_point_index = None
        self.hover_point_index = None
        self.view_extents = {"basemap": None, "focus": None}
        self.last_rendered_sid = None

        self._update_timeline_bounds()
        self._refresh_segment_listbox()
        self._refresh_info_panels()
        self.render_current_track()
        self.status_var.set(f"已定位到 {sid}，可开始人工标记。")

    def clear_track_view(self, message: str) -> None:
        self.current_sid = None
        self.current_track_df = pd.DataFrame()
        self.current_track_meta = pd.Series(dtype=object)
        self.current_segments = []
        self.current_completed = False
        self.selected_segment_order = None
        self.selected_point_index = None
        self.active_start_idx = None
        self.current_cursor_idx = None
        self.timeline_info_var.set("时间轴未加载")
        self.current_segment_hint_var.set(message)
        self.basic_info_var.set("-")
        self.track_stats_var.set("-")
        self.point_info_var.set("未选中轨迹点")
        self.segment_info_var.set("未选中异常段")
        self.db_state_var.set("-")
        self.segment_listbox.delete(0, END)
        self.ax.clear()
        self.ax.set_title(message)
        self.canvas.draw_idle()

    def _update_timeline_bounds(self) -> None:
        point_count = max(len(self.current_track_df) - 1, 0)
        self._timeline_internal_update = True
        self.timeline_scale.configure(from_=0, to=point_count)
        self.timeline_var.set(0 if point_count == 0 else min(self.timeline_var.get(), point_count))
        self._timeline_internal_update = False
        self._update_timeline_info()

    def _update_timeline_info(self) -> None:
        if self.current_track_df.empty:
            self.timeline_info_var.set("时间轴未加载")
            return
        idx = int(self.timeline_var.get())
        idx = max(0, min(idx, len(self.current_track_df) - 1))
        row = self.current_track_df.iloc[idx]
        start_text = "-"
        if self.active_start_idx is not None:
            start_text = f"{self.active_start_idx} | {format_timestamp(self.current_track_df.iloc[self.active_start_idx]['ISO_TIME'])}"
        end_text = f"{idx} | {format_timestamp(row['ISO_TIME'])}"
        self.timeline_info_var.set(f"起始点锁定：{start_text}    当前结束点/游标：{end_text}")

        if self.active_start_idx is not None:
            preview = build_segment_metrics(self.current_track_df, self.active_start_idx, idx)
            self.current_segment_hint_var.set(
                f"正在预览：[{preview.start_idx}-{preview.end_idx}] {preview.start_time} -> {preview.end_time} | "
                f"{preview.point_count} 点 | {safe_hours(preview.duration_hours)} | {safe_percent(preview.time_ratio)}"
            )
        else:
            self.current_segment_hint_var.set("当前没有正在预览的异常段。")

    def _draw_basemap(self, ax, focus_extent: Optional[Tuple[float, float, float, float]] = None) -> None:
        ax.coastlines(resolution="110m")
        ax.add_feature(cfeature.BORDERS, linewidth=0.4)
        ax.add_feature(cfeature.LAND, facecolor="#f1f5f9", alpha=0.8)
        ax.add_feature(cfeature.OCEAN, facecolor="#dbeafe", alpha=0.8)
        ax.gridlines(draw_labels=False, linewidth=0.3, color="#94a3b8", alpha=0.7)
        if self.view_mode_var.get() == "focus" and focus_extent is not None:
            ax.set_extent(focus_extent, crs=ccrs.PlateCarree())
        else:
            ax.set_extent(MAP_EXTENT, crs=ccrs.PlateCarree())

    @staticmethod
    def _normalize_extent(extent: Tuple[float, float, float, float]) -> Tuple[float, float, float, float]:
        lon0, lon1, lat0, lat1 = extent
        lon_min, lon_max = sorted([float(lon0), float(lon1)])
        lat_min, lat_max = sorted([float(lat0), float(lat1)])
        lon_min = max(MAP_EXTENT[0], lon_min)
        lon_max = min(MAP_EXTENT[1], lon_max)
        lat_min = max(MAP_EXTENT[2], lat_min)
        lat_max = min(MAP_EXTENT[3], lat_max)

        if lon_max - lon_min < 1.0:
            center = (lon_min + lon_max) / 2.0
            lon_min = max(MAP_EXTENT[0], center - 0.5)
            lon_max = min(MAP_EXTENT[1], center + 0.5)
        if lat_max - lat_min < 1.0:
            center = (lat_min + lat_max) / 2.0
            lat_min = max(MAP_EXTENT[2], center - 0.5)
            lat_max = min(MAP_EXTENT[3], center + 0.5)
        return lon_min, lon_max, lat_min, lat_max

    def _current_axes_extent(self) -> Tuple[float, float, float, float]:
        return self._normalize_extent((*self.ax.get_xlim(), *self.ax.get_ylim()))

    def _remember_current_extent(self) -> None:
        if self.last_rendered_sid == self.current_sid and self.ax.has_data():
            self.view_extents[self.last_rendered_view_mode] = self._current_axes_extent()

    def _default_extent_for_mode(self, focus_extent: Tuple[float, float, float, float]) -> Tuple[float, float, float, float]:
        if self.view_mode_var.get() == "focus":
            return self._normalize_extent(focus_extent)
        return self._normalize_extent(MAP_EXTENT)

    def _active_extent_for_render(self, focus_extent: Tuple[float, float, float, float]) -> Tuple[float, float, float, float]:
        stored = self.view_extents.get(self.view_mode_var.get())
        if stored is not None:
            return self._normalize_extent(stored)
        return self._default_extent_for_mode(focus_extent)

    def on_scroll_zoom(self, event) -> None:
        if event.inaxes != self.ax or self.current_track_df.empty:
            return
        if event.xdata is None or event.ydata is None:
            return

        x0, x1 = self.ax.get_xlim()
        y0, y1 = self.ax.get_ylim()
        scale = 0.85 if getattr(event, "button", None) == "up" else 1.18
        new_w = (x1 - x0) * scale
        new_h = (y1 - y0) * scale
        rel_x = (event.xdata - x0) / (x1 - x0) if (x1 - x0) else 0.5
        rel_y = (event.ydata - y0) / (y1 - y0) if (y1 - y0) else 0.5
        left = event.xdata - new_w * rel_x
        right = left + new_w
        bottom = event.ydata - new_h * rel_y
        top = bottom + new_h
        extent = self._normalize_extent((left, right, bottom, top))
        self.ax.set_extent(extent, crs=ccrs.PlateCarree())
        self.view_extents[self.view_mode_var.get()] = extent
        self.canvas.draw_idle()

    def toggle_segment_list_panel(self) -> None:
        self.segment_list_collapsed = not self.segment_list_collapsed
        if self.segment_list_collapsed:
            self.segment_list_container.pack_forget()
            self.segment_toggle_button.configure(text="展开已标记异常段")
        else:
            self.segment_list_container.pack(fill=BOTH, expand=True, padx=8, pady=8)
            self.segment_toggle_button.configure(text="收起已标记异常段")

    def toggle_track_points_visibility(self) -> None:
        self.track_points_visible = not self.track_points_visible
        self.track_points_toggle_text.set("隐藏轨迹点" if self.track_points_visible else "显示轨迹点")
        self.render_current_track()

    def _on_output_inner_configure(self, _event=None) -> None:
        if hasattr(self, "output_canvas"):
            self.output_canvas.configure(scrollregion=self.output_canvas.bbox("all"))

    def _on_output_canvas_configure(self, event) -> None:
        if hasattr(self, "output_canvas") and hasattr(self, "output_canvas_window"):
            self.output_canvas.itemconfigure(self.output_canvas_window, width=event.width)

    def _on_filter_inner_configure(self, _event=None) -> None:
        if hasattr(self, "filter_controls_canvas"):
            self.filter_controls_canvas.configure(scrollregion=self.filter_controls_canvas.bbox("all"))

    def _on_filter_canvas_configure(self, event) -> None:
        if hasattr(self, "filter_controls_canvas") and hasattr(self, "filter_controls_window"):
            self.filter_controls_canvas.itemconfigure(self.filter_controls_window, width=event.width)

    @staticmethod
    def _scroll_canvas_with_mousewheel(event, canvas) -> None:
        if canvas is None:
            return
        delta = getattr(event, "delta", 0)
        if delta == 0:
            return
        canvas.yview_scroll(int(-delta / 120), "units")

    def _draw_highlight_path(
        self,
        ax,
        lons: Sequence[float],
        lats: Sequence[float],
        color: str,
        *,
        zorder: float,
        linewidth: float,
        halo_width: float,
        halo_color: str = "white",
        alpha: float = 0.96,
        picker: Optional[float] = None,
    ) -> Optional[Line2D]:
        if len(lons) < 2 or len(lats) < 2:
            return None
        ax.plot(
            lons,
            lats,
            color=halo_color,
            linewidth=halo_width,
            alpha=0.92,
            transform=ccrs.PlateCarree(),
            zorder=zorder,
        )
        line = ax.plot(
            lons,
            lats,
            color=color,
            linewidth=linewidth,
            alpha=alpha,
            transform=ccrs.PlateCarree(),
            zorder=zorder + 0.1,
            picker=picker,
        )[0]
        return line

    def _get_focus_extent(self) -> Tuple[float, float, float, float]:
        if self.current_track_df.empty:
            return MAP_EXTENT
        lons = self.current_track_df["LON"].to_numpy()
        lats = self.current_track_df["LAT"].to_numpy()
        lon_pad = max((lons.max() - lons.min()) * 0.2, 4.0)
        lat_pad = max((lats.max() - lats.min()) * 0.2, 3.0)
        lon_min = max(MAP_EXTENT[0], float(lons.min() - lon_pad))
        lon_max = min(MAP_EXTENT[1], float(lons.max() + lon_pad))
        lat_min = max(MAP_EXTENT[2], float(lats.min() - lat_pad))
        lat_max = min(MAP_EXTENT[3], float(lats.max() + lat_pad))
        return lon_min, lon_max, lat_min, lat_max

    def render_current_track(self) -> None:
        self._remember_current_extent()
        self.ax.clear()
        self.segment_line_artists = []
        self.track_scatter = None

        if self.current_track_df.empty:
            self.ax.set_title("没有可显示的路径")
            self.canvas.draw_idle()
            return

        focus_extent = self._get_focus_extent()
        self._draw_basemap(self.ax, focus_extent=focus_extent)
        target_extent = self._active_extent_for_render(focus_extent)

        lons = self.current_track_df["LON"].to_numpy()
        lats = self.current_track_df["LAT"].to_numpy()
        winds = self.current_track_df["USA_WIND"].to_numpy()
        point_colors = get_intensity_colors(winds, self.intensity_bins, self.intensity_colors)
        points = np.array([lons, lats]).T.reshape(-1, 1, 2)
        if len(points) >= 2:
            segments = np.concatenate([points[:-1], points[1:]], axis=1)
            seg_colors = point_colors[:-1] if len(point_colors) >= 2 else point_colors
            lc = LineCollection(
                segments,
                colors=seg_colors,
                linewidths=BASE_TRACK_LINEWIDTH,
                transform=ccrs.PlateCarree(),
                alpha=BASE_TRACK_ALPHA,
            )
            self.ax.add_collection(lc)

        self.track_scatter = self.ax.scatter(
            lons,
            lats,
            s=28,
            c=point_colors,
            edgecolors="#0f172a" if self.track_points_visible else "none",
            linewidths=0.4 if self.track_points_visible else 0.0,
            transform=ccrs.PlateCarree(),
            zorder=5,
            picker=TRACK_POINT_PICK_RADIUS,
            alpha=1.0 if self.track_points_visible else 0.0,
        )
        self.track_scatter.my_track_df = self.current_track_df.reset_index(drop=True)
        self.track_scatter.my_artist_type = "track_points"

        for seg in self.current_segments:
            seg_df = self.current_track_df.iloc[seg.start_idx : seg.end_idx + 1]
            color = segment_color(seg.segment_order)
            line = self._draw_highlight_path(
                self.ax,
                seg_df["LON"],
                seg_df["LAT"],
                color,
                zorder=7,
                linewidth=HIGHLIGHT_LINEWIDTH,
                halo_width=HIGHLIGHT_HALO_WIDTH,
                picker=MARKED_LINE_PICK_RADIUS,
            )
            if line is None:
                continue
            line.my_artist_type = "saved_segment"
            line.segment_order = seg.segment_order
            self.segment_line_artists.append(line)
            if self.track_points_visible:
                self.ax.scatter(
                    seg_df["LON"],
                    seg_df["LAT"],
                    s=42,
                    c=color,
                    edgecolors="white",
                    linewidths=0.6,
                    transform=ccrs.PlateCarree(),
                    zorder=8,
                )

        if self.selected_segment_order is not None:
            selected = self.get_segment_by_order(self.selected_segment_order)
            if selected:
                seg_df = self.current_track_df.iloc[selected.start_idx : selected.end_idx + 1]
                self._draw_highlight_path(
                    self.ax,
                    seg_df["LON"],
                    seg_df["LAT"],
                    color="#111827",
                    zorder=8.5,
                    linewidth=2.4,
                    halo_width=HIGHLIGHT_HALO_WIDTH + 1.0,
                    halo_color="#fef3c7",
                    alpha=0.35,
                )

        if self.active_start_idx is not None and self.current_cursor_idx is not None:
            preview = build_segment_metrics(self.current_track_df, self.active_start_idx, self.current_cursor_idx)
            preview_df = self.current_track_df.iloc[preview.start_idx : preview.end_idx + 1]
            self._draw_highlight_path(
                self.ax,
                preview_df["LON"],
                preview_df["LAT"],
                color="#dc2626",
                zorder=9,
                linewidth=PREVIEW_LINEWIDTH,
                halo_width=PREVIEW_HALO_WIDTH,
                halo_color="#fee2e2",
            )

        self.hover_marker = self.ax.scatter(
            [],
            [],
            s=70,
            c="#0ea5e9",
            edgecolors="white",
            linewidths=0.8,
            transform=ccrs.PlateCarree(),
            zorder=10,
            visible=False,
        )

        if self.selected_point_index is not None and 0 <= self.selected_point_index < len(self.current_track_df):
            row = self.current_track_df.iloc[self.selected_point_index]
            self.ax.scatter(
                [row["LON"]],
                [row["LAT"]],
                s=90,
                c="#1d4ed8",
                edgecolors="white",
                linewidths=1.0,
                transform=ccrs.PlateCarree(),
                zorder=11,
            )

        if self.active_start_idx is not None:
            row = self.current_track_df.iloc[self.active_start_idx]
            self.ax.scatter(
                [row["LON"]],
                [row["LAT"]],
                s=110,
                c="#16a34a",
                edgecolors="white",
                linewidths=1.0,
                transform=ccrs.PlateCarree(),
                zorder=12,
            )

        if self.current_cursor_idx is not None and 0 <= self.current_cursor_idx < len(self.current_track_df):
            row = self.current_track_df.iloc[self.current_cursor_idx]
            color = "#dc2626" if self.active_start_idx is not None else "#0ea5e9"
            self.ax.scatter(
                [row["LON"]],
                [row["LAT"]],
                s=86,
                c=color,
                edgecolors="white",
                linewidths=0.9,
                transform=ccrs.PlateCarree(),
                zorder=12,
            )

        sid = self.current_track_meta.get("SID", "")
        name = self.current_track_meta.get("NAME", "")
        season = safe_int_text(self.current_track_meta.get("SEASON"))
        title_mode = "底图视图" if self.view_mode_var.get() == "basemap" else "聚焦放大视图"
        self.ax.set_extent(target_extent, crs=ccrs.PlateCarree())
        self.ax.set_title(f"{sid} | {name} | {season} | {title_mode}", fontsize=11)
        self.last_rendered_view_mode = self.view_mode_var.get()
        self.last_rendered_sid = self.current_sid
        self.canvas.draw_idle()

    def on_timeline_changed(self, _value=None) -> None:
        if self._timeline_internal_update or self.current_track_df.empty:
            return
        idx = int(float(self.timeline_var.get()))
        idx = max(0, min(idx, len(self.current_track_df) - 1))
        self.current_cursor_idx = idx
        if self.active_start_idx is None:
            self.selected_point_index = idx
            seg = point_in_segment(idx, self.current_segments)
            self.selected_segment_order = seg.segment_order if seg else None
        self._update_timeline_info()
        self._refresh_info_panels()
        self.render_current_track()

    def on_pick_event(self, event) -> None:
        artist = event.artist
        mouse_event = event.mouseevent
        button = getattr(mouse_event, "button", None)

        if getattr(artist, "my_artist_type", None) == "track_points":
            if not event.ind:
                return
            idx = int(event.ind[0])
            self.context_point_index = idx
            self.selected_point_index = idx
            self.current_cursor_idx = idx
            self._timeline_internal_update = True
            self.timeline_var.set(idx)
            self._timeline_internal_update = False
            existing = point_in_segment(idx, self.current_segments)
            self.selected_segment_order = existing.segment_order if existing else self.selected_segment_order

            if button == 1 and self.active_start_idx is not None:
                self.current_cursor_idx = idx
            elif button == 1:
                self.selected_segment_order = existing.segment_order if existing else None

            self._update_timeline_info()
            self._refresh_info_panels()
            self.render_current_track()

            if button == 3:
                try:
                    self.context_menu.tk_popup(mouse_event.guiEvent.x_root, mouse_event.guiEvent.y_root)
                finally:
                    self.context_menu.grab_release()
            return

        if getattr(artist, "my_artist_type", None) == "saved_segment":
            order = int(getattr(artist, "segment_order"))
            self.selected_segment_order = order
            seg = self.get_segment_by_order(order)
            if seg:
                self.selected_point_index = seg.start_idx
                self.current_cursor_idx = seg.end_idx
                self._timeline_internal_update = True
                self.timeline_var.set(seg.end_idx)
                self._timeline_internal_update = False
            self._refresh_segment_listbox()
            self._refresh_info_panels()
            self.render_current_track()

    def on_motion_event(self, event) -> None:
        if self.track_scatter is None or self.current_track_df.empty or self.hover_marker is None:
            return
        visible = False
        annotation_text = None
        hover_idx = None

        if event.inaxes == self.ax:
            contains, info = self.track_scatter.contains(event)
            indices = info.get("ind", []) if isinstance(info, dict) else []
            if contains and len(indices) > 0:
                idx = int(indices[0])
                hover_idx = idx
                row = self.current_track_df.iloc[idx]
                self.hover_marker.set_offsets([[row["LON"], row["LAT"]]])
                self.hover_marker.set_visible(True)
                visible = True
                existing = point_in_segment(idx, self.current_segments)
                if existing:
                    annotation_text = f"已标记：第 {existing.segment_order} 段"
                elif self.active_start_idx is None:
                    annotation_text = "可右键设为起始点"
                else:
                    annotation_text = "可左键预览终点 / 右键完成"

        if self.hover_annotation is None:
            self.hover_annotation = self.ax.annotate(
                "",
                xy=(0, 0),
                xytext=(10, 10),
                textcoords="offset points",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#64748b", alpha=0.95),
                fontsize=9,
            )
            self.hover_annotation.set_visible(False)

        if visible and annotation_text and event.xdata is not None and event.ydata is not None:
            self.hover_annotation.xy = (event.xdata, event.ydata)
            self.hover_annotation.set_text(annotation_text)
            self.hover_annotation.set_visible(True)
        else:
            self.hover_marker.set_visible(False)
            if self.hover_annotation is not None:
                self.hover_annotation.set_visible(False)
        self.hover_point_index = hover_idx
        self.canvas.draw_idle()

    def set_current_point_as_start(self) -> None:
        if self.current_cursor_idx is None:
            return
        self.set_start_index(self.current_cursor_idx)

    def set_start_index(self, index: int) -> None:
        self.active_start_idx = int(index)
        self.current_cursor_idx = int(index)
        self.selected_point_index = int(index)
        self._timeline_internal_update = True
        self.timeline_var.set(int(index))
        self._timeline_internal_update = False
        self._update_timeline_info()
        self._refresh_info_panels()
        self.render_current_track()

    def set_end_index(self, index: int) -> None:
        if self.active_start_idx is None:
            self.set_start_index(index)
            return
        self.current_cursor_idx = int(index)
        self.selected_point_index = int(index)
        self._timeline_internal_update = True
        self.timeline_var.set(int(index))
        self._timeline_internal_update = False
        self._update_timeline_info()
        self._refresh_info_panels()
        self.render_current_track()

    def clear_active_segment(self) -> None:
        self.active_start_idx = None
        self._update_timeline_info()
        self._refresh_info_panels()
        self.render_current_track()

    def commit_active_segment(self) -> None:
        if self.current_track_df.empty:
            return
        if self.active_start_idx is None or self.current_cursor_idx is None:
            messagebox.showinfo("缺少起止点", "请先确定起始点，再通过时间轴或点击轨迹点确定结束点。")
            return

        preview = build_segment_metrics(self.current_track_df, self.active_start_idx, self.current_cursor_idx)
        if overlap_exists(self.current_segments, preview.start_idx, preview.end_idx):
            messagebox.showwarning("区间重叠", "新标记的异常段与已保存异常段重叠，请先删除旧段或重新选择。")
            return

        preview.segment_order = len(self.current_segments) + 1
        self.current_segments.append(preview)
        self.current_segments = self._renumber_segments(self.current_segments)
        remapped = point_in_segment(preview.start_idx, self.current_segments)
        self.selected_segment_order = remapped.segment_order if remapped else None
        self.current_completed = False
        self.save_current_state(action="segment_add", detail=f"新增异常段 {preview.start_idx}-{preview.end_idx}")
        self.clear_active_segment()
        self._refresh_segment_listbox()
        self._refresh_info_panels()
        self.render_current_track()

    def delete_selected_segment(self) -> None:
        if self.selected_segment_order is None:
            messagebox.showinfo("未选中异常段", "请先在右侧列表或图上选中一个已标记异常段。")
            return
        before = len(self.current_segments)
        self.current_segments = [seg for seg in self.current_segments if seg.segment_order != self.selected_segment_order]
        if len(self.current_segments) == before:
            return
        deleted_order = self.selected_segment_order
        self.current_segments = self._renumber_segments(self.current_segments)
        self.selected_segment_order = None
        self.current_completed = False
        self.save_current_state(action="segment_delete", detail=f"删除异常段 {deleted_order}")
        self._refresh_segment_listbox()
        self._refresh_info_panels()
        self.render_current_track()

    def context_set_start(self) -> None:
        if self.context_point_index is None:
            return
        self.set_start_index(self.context_point_index)

    def context_set_end_only(self) -> None:
        if self.context_point_index is None:
            return
        self.set_end_index(self.context_point_index)

    def context_set_end_and_commit(self) -> None:
        if self.context_point_index is None:
            return
        self.set_end_index(self.context_point_index)
        self.commit_active_segment()

    def context_select_existing_segment(self) -> None:
        if self.context_point_index is None:
            return
        seg = point_in_segment(self.context_point_index, self.current_segments)
        if not seg:
            return
        self.selected_segment_order = seg.segment_order
        self._refresh_segment_listbox()
        self._refresh_info_panels()
        self.render_current_track()

    def on_segment_list_selected(self, _event=None) -> None:
        selection = self.segment_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        if idx >= len(self.current_segments):
            return
        order = self.current_segments[idx].segment_order
        self.selected_segment_order = order
        self._refresh_info_panels()
        self.render_current_track()

    def get_segment_by_order(self, order: Optional[int]) -> Optional[SegmentRecord]:
        if order is None:
            return None
        for seg in self.current_segments:
            if seg.segment_order == order:
                return seg
        return None

    def _renumber_segments(self, segments: Sequence[SegmentRecord]) -> List[SegmentRecord]:
        result = []
        for order, seg in enumerate(sorted(segments, key=lambda item: (item.start_idx, item.end_idx)), start=1):
            result.append(
                SegmentRecord(
                    segment_order=order,
                    start_idx=seg.start_idx,
                    end_idx=seg.end_idx,
                    start_time=seg.start_time,
                    end_time=seg.end_time,
                    point_count=seg.point_count,
                    duration_hours=seg.duration_hours,
                    point_ratio=seg.point_ratio,
                    time_ratio=seg.time_ratio,
                )
            )
        return result

    def _refresh_info_panels(self) -> None:
        if self.current_track_df.empty:
            return

        sid = self.current_track_meta.get("SID", "-")
        name = self.current_track_meta.get("NAME", "-")
        season = safe_int_text(self.current_track_meta.get("SEASON"))
        total_points = int(self.current_track_meta.get("TOTAL_POINTS", len(self.current_track_df)))
        total_hours = float(self.current_track_meta.get("TOTAL_DURATION_HOURS", 0.0))

        self.basic_info_var.set(
            "\n".join(
                [
                    f"SID: {sid}",
                    f"NAME: {name}",
                    f"SEASON: {season}",
                    f"轨迹点总数: {total_points} 个",
                    f"完整路径持续时长: {safe_hours(total_hours)}",
                ]
            )
        )

        seg_count = len(self.current_segments)
        total_ratio = float(sum(seg.time_ratio for seg in self.current_segments))
        max_ratio = float(max([seg.time_ratio for seg in self.current_segments], default=0.0))
        self.track_stats_var.set(
            "\n".join(
                [
                    f"异常段数量: {seg_count}",
                    f"所有异常段累积时序占比: {safe_percent(total_ratio)}",
                    f"最大异常段时序占比: {safe_percent(max_ratio)}",
                ]
            )
        )

        state = self.repo.get_track_state(self.current_sid) if self.repo and self.current_sid else {}
        completed_text = "是" if int(state.get("completed", 0) or 0) else "否"
        last_edit = state.get("last_edit_at") or "-"
        self.db_state_var.set(
            "\n".join(
                [
                    f"是否标记完成: {completed_text}",
                    f"最后一次编辑: {last_edit}",
                    f"数据库记录异常段数: {int(state.get('segment_count', seg_count) or 0)}",
                ]
            )
        )

        if self.selected_point_index is not None and 0 <= self.selected_point_index < len(self.current_track_df):
            row = self.current_track_df.iloc[self.selected_point_index]
            self.point_info_var.set(
                "\n".join(
                    [
                        f"索引: {self.selected_point_index}",
                        f"时间: {format_timestamp(row['ISO_TIME'])}",
                        f"经纬度: ({row['LAT']:.2f}N, {row['LON']:.2f}E)",
                        f"风速: {safe_int_text(row['USA_WIND'])} kt",
                    ]
                )
            )
        else:
            self.point_info_var.set("未选中轨迹点")

        selected = self.get_segment_by_order(self.selected_segment_order)
        if selected is None and self.active_start_idx is not None and self.current_cursor_idx is not None:
            selected = build_segment_metrics(self.current_track_df, self.active_start_idx, self.current_cursor_idx)
            title = "当前预览区间"
        elif selected is not None:
            title = f"第 {selected.segment_order} 段"
        else:
            title = ""

        if selected is None:
            self.segment_info_var.set("未选中异常段")
        else:
            self.segment_info_var.set(
                "\n".join(
                    [
                        title,
                        f"索引区间: {selected.start_idx} - {selected.end_idx}",
                        f"时间区间: {selected.start_time} -> {selected.end_time}",
                        f"轨迹点数: {selected.point_count}",
                        f"累计时间: {safe_hours(selected.duration_hours)}",
                        f"点数占比: {safe_percent(selected.point_ratio)}",
                        f"时序占比: {safe_percent(selected.time_ratio)}",
                    ]
                )
            )

    def _refresh_segment_listbox(self) -> None:
        self.segment_listbox.delete(0, END)
        for seg in self.current_segments:
            label = f"第{seg.segment_order}段 | {seg.start_idx}-{seg.end_idx} | {safe_percent(seg.time_ratio)} | {safe_hours(seg.duration_hours)}"
            self.segment_listbox.insert(END, label)

        if self.selected_segment_order is not None:
            for idx, seg in enumerate(self.current_segments):
                if seg.segment_order == self.selected_segment_order:
                    self.segment_listbox.selection_clear(0, END)
                    self.segment_listbox.selection_set(idx)
                    self.segment_listbox.see(idx)
                    return

    def save_current_state(self, action: str, detail: str) -> None:
        if self.repo is None or self.current_track_df.empty:
            return
        self.repo.save_track_segments(self.current_track_meta, self.current_segments, self.current_completed, action, detail)
        self.export_current_track_outputs()
        self.export_all_tracks_summary()
        filtered_df = self.catalog_df[self.catalog_df["SID"].isin(self.filtered_sids)] if self.filtered_sids else self.catalog_df
        self._refresh_track_listbox(filtered_df)
        if self.current_sid in self.listbox_sids:
            idx = self.listbox_sids.index(self.current_sid)
            self.track_listbox.selection_clear(0, END)
            self.track_listbox.selection_set(idx)
            self.track_listbox.see(idx)
        self._refresh_info_panels()

    def mark_current_track_completed(self) -> None:
        if self.current_track_df.empty:
            return
        self.current_completed = True
        self.save_current_state(action="track_complete", detail="保存并标记完成")
        self.status_var.set(f"{self.current_sid} 已标记为完成。")

    def mark_current_track_incomplete(self) -> None:
        if self.current_track_df.empty:
            return
        self.current_completed = False
        self.save_current_state(action="track_incomplete", detail="标记为未完成")
        self.status_var.set(f"{self.current_sid} 已切换为未完成状态。")

    def regenerate_outputs(self) -> None:
        if self.current_track_df.empty:
            return
        self.save_current_state(action="regenerate", detail="重新生成当前路径所有输出")
        self.status_var.set(f"{self.current_sid} 的输出已重生成。")

    def _remove_track_output_files(self, track_meta: pd.Series) -> None:
        if track_meta.empty:
            return
        stub = build_output_stub(track_meta)
        for file_path in [
            TEXT_OUTPUT_DIR / f"{stub}_summary.txt",
            PLOT_OUTPUT_DIR / f"{stub}_marked.png",
            TRACK_XLSX_OUTPUT_DIR / f"{stub}_summary.xlsx",
        ]:
            file_path.unlink(missing_ok=True)
        segment_dir = SEGMENT_OUTPUT_DIR / stub
        if segment_dir.exists():
            shutil.rmtree(segment_dir, ignore_errors=True)

    def export_current_track_outputs(self) -> None:
        if self.current_track_df.empty:
            return

        stub = build_output_stub(self.current_track_meta)
        if not self.current_segments:
            self._remove_track_output_files(self.current_track_meta)
            self.cleanup_segment_output_dirs()
            return

        segment_dir = ensure_dir(SEGMENT_OUTPUT_DIR / stub)
        for old_csv in segment_dir.glob("*.csv"):
            old_csv.unlink(missing_ok=True)

        segment_rows = []
        for seg in self.current_segments:
            seg_df = self.current_track_df.iloc[seg.start_idx : seg.end_idx + 1].copy()
            csv_path = segment_dir / f"{stub}_segment_{seg.segment_order:02d}.csv"
            seg_df[BASE_COLUMNS].to_csv(csv_path, index=False, encoding=CSV_OUTPUT_ENCODING)
            segment_rows.append(
                {
                    "segment_order": seg.segment_order,
                    "start_idx": seg.start_idx,
                    "end_idx": seg.end_idx,
                    "start_time": seg.start_time,
                    "end_time": seg.end_time,
                    "point_count": seg.point_count,
                    "duration_hours": seg.duration_hours,
                    "point_ratio": seg.point_ratio,
                    "time_ratio": seg.time_ratio,
                    "csv_path": str(csv_path),
                }
            )

        text_path = TEXT_OUTPUT_DIR / f"{stub}_summary.txt"
        plot_path = PLOT_OUTPUT_DIR / f"{stub}_marked.png"
        xlsx_path = TRACK_XLSX_OUTPUT_DIR / f"{stub}_summary.xlsx"

        self._write_track_summary_text(text_path, segment_rows)
        self._write_track_plot(plot_path)
        self._write_track_summary_xlsx(xlsx_path, segment_rows)
        self.cleanup_segment_output_dirs()

    def regenerate_all_marked_pngs(self) -> None:
        if self.repo is None:
            return
        state_df = self.repo.get_all_track_states()
        if state_df.empty or "segment_count" not in state_df.columns:
            messagebox.showinfo("无已标记路径", "当前数据库里没有可批量导出的已标记路径。")
            return

        marked_sids = state_df[state_df["segment_count"].fillna(0).astype(int) > 0]["sid"].astype(str).tolist()
        if not marked_sids:
            messagebox.showinfo("无已标记路径", "当前数据库里没有可批量导出的已标记路径。")
            return

        original_sid = self.current_sid
        exported_count = 0
        for sid in marked_sids:
            group = self.track_df[self.track_df["SID"] == sid].copy()
            meta_df = self.catalog_df[self.catalog_df["SID"] == sid]
            if group.empty or meta_df.empty:
                continue
            group = group.sort_values(by=["ISO_TIME", "ROW_ORDER"], kind="stable").reset_index(drop=True)
            segments = self.repo.load_segments(sid)
            if not segments:
                continue
            current_snapshot = (
                self.current_sid,
                self.current_track_df,
                self.current_track_meta,
                self.current_segments,
                self.current_completed,
            )
            self.current_sid = sid
            self.current_track_df = group
            self.current_track_meta = meta_df.iloc[0]
            self.current_segments = segments
            self.current_completed = bool(int(self.repo.get_track_state(sid).get("completed", 0) or 0))
            plot_path = PLOT_OUTPUT_DIR / f"{build_output_stub(self.current_track_meta)}_marked.png"
            self._write_track_plot(plot_path)
            exported_count += 1
            self.current_sid, self.current_track_df, self.current_track_meta, self.current_segments, self.current_completed = current_snapshot

        if original_sid:
            self.load_track_by_sid(original_sid)
        self.cleanup_segment_output_dirs()
        self.status_var.set(f"已重生成 {exported_count} 条已标记路径的 PNG。")

    def cleanup_segment_output_dirs(self) -> None:
        ensure_dir(SEGMENT_OUTPUT_DIR)
        if self.repo is None:
            return
        state_df = self.repo.get_all_track_states()
        marked_sids = (
            state_df[state_df["segment_count"].fillna(0).astype(int) > 0]["sid"].astype(str).tolist()
            if not state_df.empty and "segment_count" in state_df.columns
            else []
        )
        valid_stubs = set()
        for sid in marked_sids:
            meta_df = self.catalog_df[self.catalog_df["SID"] == sid]
            if meta_df.empty:
                continue
            valid_stubs.add(build_output_stub(meta_df.iloc[0]))

        for child in SEGMENT_OUTPUT_DIR.iterdir():
            if child.is_dir() and child.name not in valid_stubs:
                shutil.rmtree(child, ignore_errors=True)

    def _write_track_summary_text(self, path: Path, segment_rows: Sequence[Dict[str, object]]) -> None:
        ensure_dir(path.parent)
        sid = self.current_track_meta.get("SID", "-")
        name = self.current_track_meta.get("NAME", "-")
        season = safe_int_text(self.current_track_meta.get("SEASON"))
        total_points = int(self.current_track_meta.get("TOTAL_POINTS", len(self.current_track_df)))
        total_hours = float(self.current_track_meta.get("TOTAL_DURATION_HOURS", 0.0))
        total_ratio = float(sum(seg["time_ratio"] for seg in segment_rows))
        max_ratio = float(max([seg["time_ratio"] for seg in segment_rows], default=0.0))

        lines = [
            APP_TITLE,
            f"生成时间: {now_text()}",
            "",
            f"SID: {sid}",
            f"NAME: {name}",
            f"SEASON: {season}",
            f"轨迹点总数: {total_points}",
            f"完整路径时长: {total_hours:.1f} h",
            f"异常段数量: {len(segment_rows)}",
            f"所有异常段累积时序占比: {safe_percent(total_ratio)}",
            f"最大异常段时序占比: {safe_percent(max_ratio)}",
            f"当前完成状态: {'是' if self.current_completed else '否'}",
            "",
            "异常段明细：",
        ]

        if not segment_rows:
            lines.append("无异常段。")
        else:
            for row in segment_rows:
                lines.extend(
                    [
                        f"- 第 {row['segment_order']} 段",
                        f"  索引区间: {row['start_idx']} - {row['end_idx']}",
                        f"  时间区间: {row['start_time']} -> {row['end_time']}",
                        f"  轨迹点数: {row['point_count']}",
                        f"  时间累计: {row['duration_hours']:.1f} h",
                        f"  点数占比: {safe_percent(row['point_ratio'])}",
                        f"  时序占比: {safe_percent(row['time_ratio'])}",
                        f"  CSV: {row['csv_path']}",
                    ]
                )
        with path.open("w", encoding=OUTPUT_ENCODING) as fp:
            fp.write("\n".join(lines))

    def _write_track_summary_xlsx(self, path: Path, segment_rows: Sequence[Dict[str, object]]) -> None:
        ensure_dir(path.parent)
        overall_df = pd.DataFrame(
            [
                {
                    "SID": self.current_track_meta.get("SID"),
                    "NAME": self.current_track_meta.get("NAME"),
                    "SEASON": self.current_track_meta.get("SEASON"),
                    "CATALOG_ORDER": self.current_track_meta.get("CATALOG_ORDER"),
                    "TOTAL_POINTS": self.current_track_meta.get("TOTAL_POINTS"),
                    "TOTAL_DURATION_HOURS": self.current_track_meta.get("TOTAL_DURATION_HOURS"),
                    "SEGMENT_COUNT": len(segment_rows),
                    "TOTAL_ANOMALY_RATIO": float(sum(seg["time_ratio"] for seg in segment_rows)),
                    "MAX_SEGMENT_RATIO": float(max([seg["time_ratio"] for seg in segment_rows], default=0.0)),
                    "COMPLETED": int(self.current_completed),
                    "LAST_EXPORT_AT": now_text(),
                }
            ]
        )
        segments_df = pd.DataFrame(segment_rows)
        with pd.ExcelWriter(path) as writer:
            overall_df.to_excel(writer, index=False, sheet_name="overall")
            segments_df.to_excel(writer, index=False, sheet_name="segments")

    def _write_track_plot(self, path: Path) -> None:
        ensure_dir(path.parent)
        fig = Figure(figsize=(10, 6.8))
        ax = fig.add_subplot(111, projection=ccrs.PlateCarree())
        ax.coastlines(resolution="110m")
        ax.add_feature(cfeature.BORDERS, linewidth=0.4)
        ax.add_feature(cfeature.LAND, facecolor="#f1f5f9", alpha=0.85)
        ax.add_feature(cfeature.OCEAN, facecolor="#dbeafe", alpha=0.85)
        ax.gridlines(draw_labels=False, linewidth=0.3, color="#94a3b8", alpha=0.6)
        ax.set_extent(self._get_focus_extent(), crs=ccrs.PlateCarree())

        lons = self.current_track_df["LON"].to_numpy()
        lats = self.current_track_df["LAT"].to_numpy()
        winds = self.current_track_df["USA_WIND"].to_numpy()
        point_colors = get_intensity_colors(winds, self.intensity_bins, self.intensity_colors)
        pts = np.array([lons, lats]).T.reshape(-1, 1, 2)
        if len(pts) >= 2:
            segs = np.concatenate([pts[:-1], pts[1:]], axis=1)
            base_colors = point_colors[:-1] if len(point_colors) >= 2 else point_colors
            lc = LineCollection(
                segs,
                colors=base_colors,
                linewidths=BASE_TRACK_LINEWIDTH,
                transform=ccrs.PlateCarree(),
                alpha=BASE_TRACK_ALPHA,
            )
            ax.add_collection(lc)

        for seg in self.current_segments:
            seg_df = self.current_track_df.iloc[seg.start_idx : seg.end_idx + 1]
            color = segment_color(seg.segment_order)
            self._draw_highlight_path(
                ax,
                seg_df["LON"],
                seg_df["LAT"],
                color,
                zorder=7,
                linewidth=HIGHLIGHT_LINEWIDTH,
                halo_width=HIGHLIGHT_HALO_WIDTH,
            )

        sid = self.current_track_meta.get("SID", "")
        name = self.current_track_meta.get("NAME", "")
        season = safe_int_text(self.current_track_meta.get("SEASON"))
        ax.set_title(f"{sid} | {name} | {season} | 异常段标记图", fontsize=12)
        fig.savefig(path, dpi=EXPORT_PNG_DPI, bbox_inches="tight")

    def export_all_tracks_summary(self) -> None:
        if self.repo is None:
            return
        ensure_dir(MARKED_TRACKS_XLSX_PATH.parent)
        track_states = self.repo.get_all_track_states()
        segment_states = self.repo.get_all_segments()
        if "segment_count" not in track_states.columns:
            track_states = pd.DataFrame(
                columns=[
                    "sid",
                    "catalog_order",
                    "season",
                    "name",
                    "basin",
                    "total_points",
                    "total_duration_hours",
                    "completed",
                    "last_edit_at",
                    "segment_count",
                    "total_anomaly_ratio",
                    "max_segment_ratio",
                ]
            )
        if "sid" not in segment_states.columns:
            segment_states = pd.DataFrame(
                columns=[
                    "sid",
                    "segment_order",
                    "start_idx",
                    "end_idx",
                    "start_time",
                    "end_time",
                    "point_count",
                    "duration_hours",
                    "point_ratio",
                    "time_ratio",
                ]
            )
        marked_states = track_states[track_states["segment_count"].fillna(0).astype(int) > 0].copy()
        completed_marked_states = marked_states[marked_states["completed"].fillna(0).astype(int) == 1].copy()

        marked_tracks_df = self._collect_summary_rows(marked_states)
        completed_tracks_df = self._collect_summary_rows(completed_marked_states)

        marked_sids = marked_tracks_df["SID"].astype(str).tolist() if not marked_tracks_df.empty else []
        completed_sids = completed_tracks_df["SID"].astype(str).tolist() if not completed_tracks_df.empty else []

        marked_segments_df = segment_states[segment_states["sid"].astype(str).isin(marked_sids)].copy()
        completed_segments_df = segment_states[segment_states["sid"].astype(str).isin(completed_sids)].copy()

        with pd.ExcelWriter(MARKED_TRACKS_XLSX_PATH) as writer:
            marked_tracks_df.to_excel(writer, index=False, sheet_name="tracks")
            marked_segments_df.to_excel(writer, index=False, sheet_name="segments")

        with pd.ExcelWriter(COMPLETED_TRACKS_XLSX_PATH) as writer:
            completed_tracks_df.to_excel(writer, index=False, sheet_name="tracks")
            completed_segments_df.to_excel(writer, index=False, sheet_name="segments")

    def _collect_summary_rows(self, state_df: pd.DataFrame) -> pd.DataFrame:
        if state_df.empty:
            return pd.DataFrame(
                columns=[
                    "SID",
                    "NAME",
                    "SEASON",
                    "CATALOG_ORDER",
                    "TOTAL_POINTS",
                    "TOTAL_DURATION_HOURS",
                    "SEGMENT_COUNT",
                    "TOTAL_ANOMALY_RATIO",
                    "MAX_SEGMENT_RATIO",
                    "COMPLETED",
                    "LAST_EXPORT_AT",
                ]
            )

        state_map = {str(row["sid"]): row for _, row in state_df.iterrows()}
        records: List[Dict[str, object]] = []
        ordered_catalog = self.catalog_df[self.catalog_df["SID"].astype(str).isin(state_map.keys())].copy()
        ordered_catalog = ordered_catalog.sort_values(by=["CATALOG_ORDER", "SID"], kind="stable")

        for _, meta in ordered_catalog.iterrows():
            sid = str(meta["SID"])
            state_row = state_map.get(sid)
            if state_row is None:
                continue
            row = self._read_single_track_summary_row(meta)
            if not row:
                row = {
                    "SID": sid,
                    "NAME": meta.get("NAME"),
                    "SEASON": meta.get("SEASON"),
                    "CATALOG_ORDER": meta.get("CATALOG_ORDER"),
                    "TOTAL_POINTS": meta.get("TOTAL_POINTS"),
                    "TOTAL_DURATION_HOURS": meta.get("TOTAL_DURATION_HOURS"),
                    "SEGMENT_COUNT": int(state_row.get("segment_count", 0) or 0),
                    "TOTAL_ANOMALY_RATIO": float(state_row.get("total_anomaly_ratio", 0.0) or 0.0),
                    "MAX_SEGMENT_RATIO": float(state_row.get("max_segment_ratio", 0.0) or 0.0),
                    "COMPLETED": int(state_row.get("completed", 0) or 0),
                    "LAST_EXPORT_AT": state_row.get("last_edit_at"),
                }
            else:
                row["CATALOG_ORDER"] = meta.get("CATALOG_ORDER")
                row["COMPLETED"] = int(state_row.get("completed", row.get("COMPLETED", 0)) or 0)
                row["SEGMENT_COUNT"] = int(state_row.get("segment_count", row.get("SEGMENT_COUNT", 0)) or 0)
                row["TOTAL_ANOMALY_RATIO"] = float(
                    state_row.get("total_anomaly_ratio", row.get("TOTAL_ANOMALY_RATIO", 0.0)) or 0.0
                )
                row["MAX_SEGMENT_RATIO"] = float(
                    state_row.get("max_segment_ratio", row.get("MAX_SEGMENT_RATIO", 0.0)) or 0.0
                )
                row["LAST_EXPORT_AT"] = state_row.get("last_edit_at", row.get("LAST_EXPORT_AT"))
            records.append(row)

        result = pd.DataFrame(records)
        if result.empty:
            return result
        return result.sort_values(by=["CATALOG_ORDER", "SID"], kind="stable").reset_index(drop=True)

    def _read_single_track_summary_row(self, track_meta: pd.Series) -> Dict[str, object]:
        stub = build_output_stub(track_meta)
        xlsx_path = TRACK_XLSX_OUTPUT_DIR / f"{stub}_summary.xlsx"
        if not xlsx_path.exists():
            return {}
        try:
            overall_df = pd.read_excel(xlsx_path, sheet_name="overall")
        except Exception:
            return {}
        if overall_df.empty:
            return {}
        record = dict(overall_df.iloc[0].to_dict())
        record["SID"] = str(record.get("SID", track_meta.get("SID")))
        record["NAME"] = record.get("NAME", track_meta.get("NAME"))
        record["SEASON"] = record.get("SEASON", track_meta.get("SEASON"))
        return record

    def _open_path(self, path: Path, create: bool = False) -> None:
        if create:
            ensure_dir(path)
        elif not path.exists():
            messagebox.showinfo("路径不存在", f"目标路径不存在：\n{path}")
            return
        os.startfile(str(path))

    def open_plot_folder(self) -> None:
        self._open_path(PLOT_OUTPUT_DIR, create=True)

    def open_segment_folder(self) -> None:
        if self.current_track_meta.empty:
            self._open_path(SEGMENT_OUTPUT_DIR, create=True)
            return
        stub = build_output_stub(self.current_track_meta)
        segment_dir = SEGMENT_OUTPUT_DIR / stub
        if segment_dir.exists():
            self._open_path(segment_dir)
        else:
            self._open_path(SEGMENT_OUTPUT_DIR, create=True)

    def open_text_folder(self) -> None:
        self._open_path(TEXT_OUTPUT_DIR, create=True)

    def open_xlsx_folder(self) -> None:
        self._open_path(TRACK_XLSX_OUTPUT_DIR, create=True)

    def open_summary_folder(self) -> None:
        self._open_path(SUMMARY_OUTPUT_DIR, create=True)

    def on_close(self) -> None:
        if self.repo is not None:
            self.repo.close()
        self.root.destroy()


if __name__ == "__main__":
    ensure_output_layout()
    root = Tk()
    app = TyphoonStatisticApp(root)
    root.mainloop()
