from __future__ import annotations

import json
import os
import pickle
import tempfile
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.io as pio
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QLockFile, Qt, QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QRadioButton,
    QListWidget,
    QListWidgetItem,
    QTabWidget,
    QTableView,
    QHeaderView,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtWebEngineWidgets import QWebEngineView

from native_app.browser_dialog import CookieSnapshot, SharePointLoginDialog
from native_app.revenue_service import RevenueDataset, load_local_dataset
from native_app.sharepoint_service import SharePointCookieRecord, download_sharepoint_dataset


DEBUG_LOG_PATH = Path(tempfile.gettempdir()) / "revenue_sharepoint_debug.log"
LOCK_FILE_PATH = Path(tempfile.gettempdir()) / "revenue_analysis_desktop.lock"
SESSION_CACHE_PATH = Path(os.getenv("LOCALAPPDATA", tempfile.gettempdir())) / "RevenueAnalysisDesktop" / "sharepoint_session.json"
DATASET_CACHE_DIR = Path(os.getenv("LOCALAPPDATA", tempfile.gettempdir())) / "RevenueAnalysisDesktop" / "dataset_cache"
DATASET_CACHE_FILE = DATASET_CACHE_DIR / "dataset.pkl"
DATASET_META_FILE = DATASET_CACHE_DIR / "meta.json"
APP_LOCK: QLockFile | None = None
WINDOW_CHOICES = [
    "Current month",
    "Last month",
    "Last 2 months",
    "Last 3 months",
    "Last 6 months",
    "Last 12 months",
    "All months",
]
METRIC_CHOICES = ["GM%", "Revenue", "Cost"]


def _reset_debug_log() -> None:
    try:
        DEBUG_LOG_PATH.write_text("", encoding="utf-8")
    except Exception:
        pass


def _append_debug_log(message: str) -> None:
    try:
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(message + "\n")
    except Exception:
        pass


def _session_ttl_minutes() -> int:
    raw = os.getenv("RAS_NATIVE_SESSION_TTL_MIN", "480").strip()
    try:
        value = int(raw)
        return max(value, 15)
    except ValueError:
        return 480


def _save_session_cache(folder_url: str, cookies: list[CookieSnapshot]) -> None:
    try:
        SESSION_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "cached_at_utc": datetime.now(UTC).isoformat(),
            "folder_url": folder_url,
            "cookies": [
                {
                    "name": cookie.name,
                    "value": cookie.value,
                    "domain": cookie.domain,
                    "path": cookie.path,
                    "secure": cookie.secure,
                }
                for cookie in cookies
            ],
        }
        SESSION_CACHE_PATH.write_text(json.dumps(payload), encoding="utf-8")
        _append_debug_log(f"sharepoint_session_cache_saved | path={SESSION_CACHE_PATH} | cookie_count={len(cookies)}")
    except Exception as exc:  # noqa: BLE001
        _append_debug_log(f"sharepoint_session_cache_save_failed | error={exc}")


def _load_session_cache(folder_url: str) -> list[CookieSnapshot]:
    if not SESSION_CACHE_PATH.exists():
        return []

    try:
        payload = json.loads(SESSION_CACHE_PATH.read_text(encoding="utf-8"))
        cached_url = str(payload.get("folder_url") or "").strip()
        if not cached_url or cached_url != folder_url:
            _append_debug_log("sharepoint_session_cache_ignored | reason=url_mismatch")
            return []

        cached_at_raw = str(payload.get("cached_at_utc") or "").strip()
        if cached_at_raw:
            cached_at = datetime.fromisoformat(cached_at_raw)
            if cached_at.tzinfo is None:
                cached_at = cached_at.replace(tzinfo=UTC)
            if datetime.now(UTC) - cached_at > timedelta(minutes=_session_ttl_minutes()):
                _append_debug_log("sharepoint_session_cache_ignored | reason=expired")
                return []

        cookies_raw = payload.get("cookies") or []
        cookies: list[CookieSnapshot] = []
        for item in cookies_raw:
            name = str(item.get("name") or "").strip()
            value = str(item.get("value") or "").strip()
            if not name or not value:
                continue
            cookies.append(
                CookieSnapshot(
                    name=name,
                    value=value,
                    domain=str(item.get("domain") or ""),
                    path=str(item.get("path") or "/"),
                    secure=bool(item.get("secure")),
                )
            )
        _append_debug_log(f"sharepoint_session_cache_loaded | cookie_count={len(cookies)}")
        return cookies
    except Exception as exc:  # noqa: BLE001
        _append_debug_log(f"sharepoint_session_cache_load_failed | error={exc}")
        return []


def _clear_session_cache() -> None:
    try:
        if SESSION_CACHE_PATH.exists():
            SESSION_CACHE_PATH.unlink()
        _append_debug_log("sharepoint_session_cache_cleared")
    except Exception:
        pass


def _safe_stat_signature(path: Path) -> str:
    if not path.exists():
        return "missing"
    stat = path.stat()
    return f"{path.name}:{int(stat.st_mtime)}:{stat.st_size}"


def _local_source_signature(data_folder: Path, template_file: Path) -> str:
    if not data_folder.exists() or not template_file.exists():
        return "invalid"
    file_patterns = ["*.xlsx", "*.xlsm", "*.xls"]
    entries: list[str] = [_safe_stat_signature(template_file)]
    seen: set[Path] = set()
    for pattern in file_patterns:
        for file_path in sorted(data_folder.glob(pattern)):
            if file_path in seen:
                continue
            seen.add(file_path)
            entries.append(_safe_stat_signature(file_path))
    return "|".join(entries)


def _dataset_fingerprint(dataset: RevenueDataset) -> dict[str, float | int | str]:
    merged = dataset.merged if dataset.merged is not None else pd.DataFrame()
    revenue = float(merged["rev_month"].sum()) if "rev_month" in merged.columns else 0.0
    cost = float(merged["cost_month"].sum()) if "cost_month" in merged.columns else 0.0
    months = merged.get("month", pd.Series(dtype="datetime64[ns]")).dropna()
    max_month = ""
    if not months.empty:
        max_month = str(pd.to_datetime(months).max().date())

    return {
        "rows": int(len(merged)),
        "snapshot_rows": int(len(dataset.snapshot)),
        "summary_rows": int(len(dataset.summary)),
        "revenue": round(revenue, 2),
        "cost": round(cost, 2),
        "max_month": max_month,
    }


def _save_dataset_cache(dataset: RevenueDataset, meta: dict[str, object]) -> None:
    try:
        DATASET_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        with DATASET_CACHE_FILE.open("wb") as handle:
            pickle.dump(dataset, handle)
        payload = {
            **meta,
            "cached_at_utc": datetime.now(UTC).isoformat(),
            "fingerprint": _dataset_fingerprint(dataset),
        }
        DATASET_META_FILE.write_text(json.dumps(payload), encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        _append_debug_log(f"dataset_cache_save_failed | error={exc}")


def _load_dataset_cache() -> tuple[RevenueDataset | None, dict[str, object]]:
    if not DATASET_CACHE_FILE.exists() or not DATASET_META_FILE.exists():
        return None, {}
    try:
        with DATASET_CACHE_FILE.open("rb") as handle:
            dataset = pickle.load(handle)
        meta = json.loads(DATASET_META_FILE.read_text(encoding="utf-8"))
        if not isinstance(dataset, RevenueDataset):
            return None, {}
        return dataset, meta
    except Exception as exc:  # noqa: BLE001
        _append_debug_log(f"dataset_cache_load_failed | error={exc}")
        return None, {}


def _acquire_single_instance_lock() -> bool:
    global APP_LOCK

    lock = QLockFile(str(LOCK_FILE_PATH))
    lock.setStaleLockTime(0)
    if not lock.tryLock(0):
        return False

    APP_LOCK = lock
    return True


class DataFrameModel(QAbstractTableModel):
    def __init__(self, frame: pd.DataFrame | None = None) -> None:
        super().__init__()
        self._frame = frame if frame is not None else pd.DataFrame()

    def set_frame(self, frame: pd.DataFrame) -> None:
        self.beginResetModel()
        self._frame = frame.copy()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._frame.index)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._frame.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role not in {Qt.DisplayRole, Qt.TextAlignmentRole}:
            return None
        if role == Qt.TextAlignmentRole:
            return Qt.AlignLeft | Qt.AlignVCenter
        value = self._frame.iat[index.row(), index.column()]
        if pd.isna(value):
            return ""
        if isinstance(value, float):
            return f"{value:,.2f}"
        return str(value)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._frame.columns[section])
        return str(section + 1)


class RevenueLoaderThread(QThread):
    finished_with_data = Signal(object)
    failed = Signal(str)

    def __init__(self, loader_callable, *loader_args) -> None:
        super().__init__()
        self._loader_callable = loader_callable
        self._loader_args = loader_args

    def run(self) -> None:
        try:
            result = self._loader_callable(*self._loader_args)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
            return
        self.finished_with_data.emit(result)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Revenue Analysis Desktop")
        self.resize(1500, 950)

        self._cookies: list[CookieSnapshot] = []
        self._session_from_cache = False
        self._data_thread: RevenueLoaderThread | None = None
        self._background_thread: RevenueLoaderThread | None = None
        self._current_dataset: RevenueDataset | None = None
        self._current_fingerprint: dict[str, float | int | str] = {}
        self._syncing_filter_controls = False
        self._dataset_cache_restored = False
        self._raw_merged = pd.DataFrame()
        self._raw_snapshot = pd.DataFrame()
        self._inputs_visible = True

        sharepoint_url = os.getenv("RAS_NATIVE_SHAREPOINT_URL", "").strip()
        if not sharepoint_url:
            fallback_url = os.getenv("RAS_SP_FOLDER_PATH", "").strip()
            if fallback_url.lower().startswith("http"):
                sharepoint_url = fallback_url

        self._folder_edit = QLineEdit(os.getenv("RAS_NATIVE_DATA_FOLDER", ""), self)
        self._template_edit = QLineEdit(os.getenv("RAS_NATIVE_TEMPLATE_FILE", ""), self)
        self._sharepoint_url_edit = QLineEdit(sharepoint_url, self)
        self._sharepoint_url_edit.setReadOnly(True)

        self._local_mode_radio = QRadioButton("Local files", self)
        self._sharepoint_mode_radio = QRadioButton("SharePoint session", self)
        startup_mode = os.getenv("RAS_NATIVE_MODE", "sharepoint").strip().lower()
        if startup_mode == "local":
            self._local_mode_radio.setChecked(True)
        else:
            self._sharepoint_mode_radio.setChecked(True)
        self._auto_load = QCheckBox("Auto-load on app start", self)
        self._auto_load.setChecked(True)

        self._mode_stack = QWidget(self)
        self._stack_layout = QVBoxLayout(self._mode_stack)
        self._stack_layout.setContentsMargins(0, 0, 0, 0)
        self._local_panel = self._build_local_panel()
        self._sharepoint_panel = self._build_sharepoint_panel()
        self._stack_layout.addWidget(self._local_panel)
        self._stack_layout.addWidget(self._sharepoint_panel)

        self._load_button = QPushButton("Load Dashboard", self)
        self._load_button.clicked.connect(self._load_dashboard)
        self._toggle_inputs_button = QPushButton("Hide Inputs", self)
        self._toggle_inputs_button.clicked.connect(self._toggle_inputs_panel)

        self._status = QLabel("Ready.", self)
        self._progress = QProgressBar(self)
        self._progress.setRange(0, 0)
        self._progress.hide()

        mode_row = QHBoxLayout()
        mode_row.addWidget(self._local_mode_radio)
        mode_row.addWidget(self._sharepoint_mode_radio)
        mode_row.addWidget(self._auto_load)
        mode_row.addStretch(1)
        mode_row.addWidget(self._toggle_inputs_button)

        controls = QVBoxLayout()
        controls.addLayout(mode_row)
        controls.addWidget(self._mode_stack)
        controls.addWidget(self._load_button)
        controls.addWidget(self._progress)
        controls.addWidget(self._status)

        self._tabs = QTabWidget(self)
        self._dashboard_tab = QWidget(self)
        self._summary_tab = QWidget(self)
        self._trends_tab = QWidget(self)
        self._tabs.addTab(self._dashboard_tab, "Dashboard")
        self._tabs.addTab(self._summary_tab, "Summary")
        self._tabs.addTab(self._trends_tab, "Trends")

        self._kpi_labels = {
            "revenue": QLabel("$0.00", self),
            "cost": QLabel("$0.00", self),
            "margin": QLabel("$0.00", self),
            "gm": QLabel("0.00%", self),
        }
        self._snapshot_model = DataFrameModel()
        self._summary_model = DataFrameModel()

        self._snapshot_table = QTableView(self)
        self._snapshot_table.setModel(self._snapshot_model)
        self._snapshot_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._snapshot_table.setHorizontalScrollMode(QTableView.ScrollPerPixel)
        self._summary_table = QTableView(self)
        self._summary_table.setModel(self._summary_model)
        self._summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._summary_table.setHorizontalScrollMode(QTableView.ScrollPerPixel)

        self._trend_view = QWebEngineView(self)
        self._trend_view.setHtml("<html><body><h3>No data loaded yet.</h3></body></html>")

        self._window_combo = QComboBox(self)
        self._window_combo.addItems(WINDOW_CHOICES)
        self._window_combo.setCurrentText("Last 3 months")

        self._summary_window_combo = QComboBox(self)
        self._summary_window_combo.addItems(WINDOW_CHOICES)
        self._summary_window_combo.setCurrentText("Last 3 months")

        self._account_list = QListWidget(self)
        self._account_list.setSelectionMode(QListWidget.MultiSelection)
        self._account_list.setMaximumHeight(90)
        self._account_search_edit = QLineEdit(self)
        self._account_search_edit.setPlaceholderText("Search accounts...")
        self._account_hint = QLabel("No account selected: using all accounts.", self)
        self._project_list = QListWidget(self)
        self._project_list.setSelectionMode(QListWidget.MultiSelection)
        self._project_list.setMaximumHeight(90)
        self._project_search_edit = QLineEdit(self)
        self._project_search_edit.setPlaceholderText("Search projects...")
        self._project_hint = QLabel("No project selected: using all projects.", self)
        self._all_accounts: list[str] = []
        self._all_projects: list[str] = []
        self._metric_list = QListWidget(self)
        self._metric_list.setSelectionMode(QListWidget.MultiSelection)
        self._metric_list.setMaximumHeight(90)
        for metric in METRIC_CHOICES:
            item = QListWidgetItem(metric)
            if metric == "GM%":
                item.setSelected(True)
            self._metric_list.addItem(item)

        self._summary_metric_list = QListWidget(self)
        self._summary_metric_list.setSelectionMode(QListWidget.MultiSelection)
        self._summary_metric_list.setMaximumHeight(90)
        for metric in METRIC_CHOICES:
            item = QListWidgetItem(metric)
            if metric == "GM%":
                item.setSelected(True)
            self._summary_metric_list.addItem(item)

        self._compact_view = QCheckBox("Compact Columns", self)
        self._compact_view.setChecked(True)

        self._trend_metric_combo = QComboBox(self)
        self._trend_metric_combo.addItems(["Revenue", "Cost", "Margin"])
        self._apply_filters_button = QPushButton("Apply Filters", self)
        self._clear_filters_button = QPushButton("Clear Filters", self)
        self._toggle_dashboard_filters_button = QPushButton("Hide Filters", self)
        self._toggle_summary_filters_button = QPushButton("Hide Summary Filters", self)
        self._filter_summary = QLabel("No filters applied.", self)

        self._apply_filters_button.clicked.connect(self._apply_filters)
        self._clear_filters_button.clicked.connect(self._clear_filters)
        self._account_list.itemSelectionChanged.connect(self._sync_project_options)
        self._account_list.itemSelectionChanged.connect(self._update_selection_hints)
        self._project_list.itemSelectionChanged.connect(self._update_selection_hints)
        self._account_search_edit.textChanged.connect(self._filter_account_options)
        self._project_search_edit.textChanged.connect(self._filter_project_options)
        self._window_combo.currentIndexChanged.connect(self._apply_filters)
        self._metric_list.itemSelectionChanged.connect(self._apply_filters)
        self._summary_window_combo.currentIndexChanged.connect(self._apply_filters_from_summary_controls)
        self._summary_metric_list.itemSelectionChanged.connect(self._apply_filters_from_summary_controls)
        self._compact_view.toggled.connect(self._apply_filters)
        self._trend_metric_combo.currentIndexChanged.connect(self._apply_filters)
        self._toggle_dashboard_filters_button.clicked.connect(self._toggle_dashboard_filters)
        self._toggle_summary_filters_button.clicked.connect(self._toggle_summary_filters)

        self._refresh_timer = QTimer(self)
        self._refresh_timer.setInterval(max(1, int(os.getenv("RAS_NATIVE_REFRESH_MINUTES", "5"))) * 60 * 1000)
        self._refresh_timer.timeout.connect(self._trigger_background_refresh)

        self._build_tabs()

        root = QWidget(self)
        layout = QVBoxLayout(root)
        layout.addLayout(controls)
        layout.addWidget(self._tabs, 1)
        self.setCentralWidget(root)

        self._local_mode_radio.toggled.connect(self._sync_mode)
        self._sharepoint_mode_radio.toggled.connect(self._sync_mode)
        self._sync_mode()
        self._restore_cached_sharepoint_session()
        self._dataset_cache_restored = self._restore_dataset_cache()
        QTimer.singleShot(0, self._auto_startup)

    def _build_local_panel(self) -> QWidget:
        panel = QGroupBox("Local File Inputs", self)
        form = QFormLayout(panel)

        folder_row = QHBoxLayout()
        folder_row.addWidget(self._folder_edit)
        folder_button = QPushButton("Browse", self)
        folder_button.clicked.connect(self._choose_folder)
        folder_row.addWidget(folder_button)

        template_row = QHBoxLayout()
        template_row.addWidget(self._template_edit)
        template_button = QPushButton("Browse", self)
        template_button.clicked.connect(self._choose_template)
        template_row.addWidget(template_button)

        form.addRow("Revenue folder", self._wrap_layout(folder_row))
        form.addRow("Template file", self._wrap_layout(template_row))
        return panel

    def _build_sharepoint_panel(self) -> QWidget:
        panel = QGroupBox("SharePoint Session Inputs", self)
        form = QFormLayout(panel)

        login_row = QHBoxLayout()
        self._cookies_label = QLabel("No session captured yet.", self)
        login_button = QPushButton("Open Login Browser", self)
        login_button.clicked.connect(self._open_sharepoint_login)
        login_row.addWidget(login_button)
        login_row.addWidget(self._cookies_label)
        login_row.addStretch(1)

        form.addRow("Configured Folder URL", self._sharepoint_url_edit)
        form.addRow("Session", self._wrap_layout(login_row))
        return panel

    def _wrap_layout(self, layout: QHBoxLayout) -> QWidget:
        wrapper = QWidget(self)
        wrapper.setLayout(layout)
        return wrapper

    def _build_tabs(self) -> None:
        dashboard_layout = QVBoxLayout(self._dashboard_tab)
        dashboard_toggle_row = QHBoxLayout()
        dashboard_toggle_row.addWidget(self._toggle_dashboard_filters_button)
        dashboard_toggle_row.addStretch(1)
        dashboard_layout.addLayout(dashboard_toggle_row)

        self._dashboard_filter_box = QGroupBox("Filters", self)
        filter_layout = QVBoxLayout(self._dashboard_filter_box)

        filter_row = QHBoxLayout()

        window_col = QVBoxLayout()
        window_col.addWidget(QLabel("Month Window", self))
        window_col.addWidget(self._window_combo)
        filter_row.addLayout(window_col, 1)

        account_col = QVBoxLayout()
        account_col.addWidget(QLabel("Accounts", self))
        account_col.addWidget(self._account_search_edit)
        account_col.addWidget(self._account_list)
        account_btn_row = QHBoxLayout()
        account_select_all_button = QPushButton("Select All", self)
        account_select_all_button.clicked.connect(lambda: self._select_all_visible(self._account_list))
        account_clear_button = QPushButton("Clear", self)
        account_clear_button.clicked.connect(self._account_list.clearSelection)
        account_btn_row.addWidget(account_select_all_button)
        account_btn_row.addWidget(account_clear_button)
        account_btn_row.addStretch(1)
        account_col.addLayout(account_btn_row)
        account_col.addWidget(self._account_hint)
        filter_row.addLayout(account_col, 2)

        project_col = QVBoxLayout()
        project_col.addWidget(QLabel("Projects", self))
        project_col.addWidget(self._project_search_edit)
        project_col.addWidget(self._project_list)
        project_btn_row = QHBoxLayout()
        project_select_all_button = QPushButton("Select All", self)
        project_select_all_button.clicked.connect(lambda: self._select_all_visible(self._project_list))
        project_clear_button = QPushButton("Clear", self)
        project_clear_button.clicked.connect(self._project_list.clearSelection)
        project_btn_row.addWidget(project_select_all_button)
        project_btn_row.addWidget(project_clear_button)
        project_btn_row.addStretch(1)
        project_col.addLayout(project_btn_row)
        project_col.addWidget(self._project_hint)
        filter_row.addLayout(project_col, 2)

        metric_col = QVBoxLayout()
        metric_col.addWidget(QLabel("Metrics", self))
        metric_col.addWidget(self._metric_list)
        metric_col.addWidget(self._compact_view)
        metric_col.addStretch(1)
        filter_row.addLayout(metric_col, 1)

        filter_layout.addLayout(filter_row)

        filter_buttons = QHBoxLayout()
        filter_buttons.addWidget(self._apply_filters_button)
        filter_buttons.addWidget(self._clear_filters_button)
        filter_buttons.addStretch(1)
        filter_layout.addLayout(filter_buttons)
        filter_layout.addWidget(self._filter_summary)

        dashboard_layout.addWidget(self._dashboard_filter_box)

        kpi_row = QGridLayout()
        kpi_row.addWidget(self._build_kpi_card("Revenue", self._kpi_labels["revenue"]), 0, 0)
        kpi_row.addWidget(self._build_kpi_card("Cost", self._kpi_labels["cost"]), 0, 1)
        kpi_row.addWidget(self._build_kpi_card("Margin", self._kpi_labels["margin"]), 0, 2)
        kpi_row.addWidget(self._build_kpi_card("GM%", self._kpi_labels["gm"]), 0, 3)
        dashboard_layout.addLayout(kpi_row)
        dashboard_layout.addWidget(QLabel("Project Snapshot", self))
        dashboard_layout.addWidget(self._snapshot_table, 1)

        summary_layout = QVBoxLayout(self._summary_tab)
        summary_toggle_top = QHBoxLayout()
        summary_toggle_top.addWidget(self._toggle_summary_filters_button)
        summary_toggle_top.addStretch(1)
        summary_layout.addLayout(summary_toggle_top)

        self._summary_filter_box = QGroupBox("Summary Filters", self)
        summary_filter_layout = QVBoxLayout(self._summary_filter_box)

        summary_row = QHBoxLayout()
        summary_window_col = QVBoxLayout()
        summary_window_col.addWidget(QLabel("Month Window", self))
        summary_window_col.addWidget(self._summary_window_combo)
        summary_row.addLayout(summary_window_col, 1)

        summary_metric_col = QVBoxLayout()
        summary_metric_col.addWidget(QLabel("Metrics", self))
        summary_metric_col.addWidget(self._summary_metric_list)
        summary_metric_col.addStretch(1)
        summary_row.addLayout(summary_metric_col, 1)
        summary_filter_layout.addLayout(summary_row)

        summary_layout.addWidget(self._summary_filter_box)
        summary_layout.addWidget(QLabel("Summarized Revenue View (month + metric pivot)", self))
        summary_layout.addWidget(self._summary_table, 1)

        trends_layout = QVBoxLayout(self._trends_tab)
        trend_controls = QHBoxLayout()
        trend_controls.addWidget(QLabel("Metric", self))
        trend_controls.addWidget(self._trend_metric_combo)
        trend_controls.addStretch(1)
        trends_layout.addLayout(trend_controls)
        trends_layout.addWidget(self._trend_view, 1)

    def _toggle_inputs_panel(self) -> None:
        self._inputs_visible = not self._inputs_visible
        self._mode_stack.setVisible(self._inputs_visible)
        self._load_button.setVisible(self._inputs_visible)
        self._toggle_inputs_button.setText("Hide Inputs" if self._inputs_visible else "Show Inputs")

    def _toggle_dashboard_filters(self) -> None:
        visible = self._dashboard_filter_box.isVisible()
        self._dashboard_filter_box.setVisible(not visible)
        self._toggle_dashboard_filters_button.setText("Hide Filters" if not visible else "Show Filters")

    def _toggle_summary_filters(self) -> None:
        visible = self._summary_filter_box.isVisible()
        self._summary_filter_box.setVisible(not visible)
        self._toggle_summary_filters_button.setText("Hide Summary Filters" if not visible else "Show Summary Filters")

    def _set_metric_selection(self, widget: QListWidget, values: list[str]) -> None:
        selected = set(values)
        widget.blockSignals(True)
        for idx in range(widget.count()):
            item = widget.item(idx)
            item.setSelected(item.text() in selected)
        widget.blockSignals(False)

    def _sync_summary_controls_from_main(self) -> None:
        if self._syncing_filter_controls:
            return
        self._syncing_filter_controls = True
        self._summary_window_combo.blockSignals(True)
        self._summary_window_combo.setCurrentText(self._window_combo.currentText())
        self._summary_window_combo.blockSignals(False)
        self._set_metric_selection(self._summary_metric_list, self._selected_metrics())
        self._syncing_filter_controls = False

    def _apply_filters_from_summary_controls(self) -> None:
        if self._syncing_filter_controls:
            return
        self._syncing_filter_controls = True
        self._window_combo.blockSignals(True)
        self._window_combo.setCurrentText(self._summary_window_combo.currentText())
        self._window_combo.blockSignals(False)
        summary_metrics = self._selected_list_values(self._summary_metric_list) or ["GM%"]
        self._set_metric_selection(self._metric_list, summary_metrics)
        self._syncing_filter_controls = False
        self._apply_filters()

    def _select_all_visible(self, widget: QListWidget) -> None:
        widget.blockSignals(True)
        for idx in range(widget.count()):
            item = widget.item(idx)
            if not item.isHidden():
                item.setSelected(True)
        widget.blockSignals(False)
        self._update_selection_hints()
        self._apply_filters()

    def _filter_list_widget(self, widget: QListWidget, query: str) -> None:
        needle = query.strip().lower()
        for idx in range(widget.count()):
            item = widget.item(idx)
            show = not needle or needle in item.text().lower()
            item.setHidden(not show)

    def _filter_account_options(self) -> None:
        self._filter_list_widget(self._account_list, self._account_search_edit.text())

    def _filter_project_options(self) -> None:
        self._filter_list_widget(self._project_list, self._project_search_edit.text())

    def _update_selection_hints(self) -> None:
        account_selected = len(self._selected_list_values(self._account_list))
        project_selected = len(self._selected_list_values(self._project_list))
        account_total = len(self._all_accounts)
        project_total = len(self._all_projects)

        if account_selected == 0:
            self._account_hint.setText(f"No account selected: using all {account_total} accounts.")
        else:
            self._account_hint.setText(f"Selected {account_selected} of {account_total} accounts.")

        if project_selected == 0:
            self._project_hint.setText(f"No project selected: using all {project_total} projects.")
        else:
            self._project_hint.setText(f"Selected {project_selected} of {project_total} projects.")

    def _build_kpi_card(self, title: str, value_label: QLabel) -> QWidget:
        card = QFrame(self)
        card.setFrameShape(QFrame.StyledPanel)
        layout = QVBoxLayout(card)
        layout.addWidget(QLabel(title, card))
        layout.addWidget(value_label)
        return card

    def _sync_mode(self) -> None:
        self._local_panel.setVisible(self._local_mode_radio.isChecked())
        self._sharepoint_panel.setVisible(self._sharepoint_mode_radio.isChecked())

    def _choose_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select revenue folder")
        if folder:
            self._folder_edit.setText(folder)

    def _choose_template(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Select template file", filter="Excel files (*.xlsx)")
        if file_name:
            self._template_edit.setText(file_name)

    def _auto_startup(self) -> None:
        if not self._auto_load.isChecked():
            return

        if self._dataset_cache_restored:
            self._status.setText("Loaded cached data. Checking for updates in background...")
            self._trigger_background_refresh()
            self._refresh_timer.start()
            return

        if self._local_mode_radio.isChecked() and self._folder_edit.text().strip() and self._template_edit.text().strip():
            self._load_dashboard()
            return

        if self._sharepoint_mode_radio.isChecked() and self._sharepoint_url_edit.text().strip():
            if self._cookies:
                self._load_dashboard()
                return
            self._open_sharepoint_login()
            if self._cookies:
                self._load_dashboard()

    def _cache_source_key(self) -> str:
        if self._local_mode_radio.isChecked():
            return f"local|{self._folder_edit.text().strip()}|{self._template_edit.text().strip()}"
        return f"sharepoint|{self._sharepoint_url_edit.text().strip()}"

    def _cache_source_signature(self) -> str:
        if self._local_mode_radio.isChecked():
            data_folder = Path(self._folder_edit.text().strip())
            template_file = Path(self._template_edit.text().strip())
            return _local_source_signature(data_folder, template_file)
        return "sharepoint_dynamic"

    def _restore_dataset_cache(self) -> bool:
        dataset, meta = _load_dataset_cache()
        if dataset is None:
            return False
        if str(meta.get("source_key") or "") != self._cache_source_key():
            return False
        self._apply_dataset(dataset, loaded_from_cache=True)
        return True

    def _trigger_background_refresh(self) -> None:
        if self._background_thread is not None and self._background_thread.isRunning():
            return

        if self._local_mode_radio.isChecked():
            data_folder = Path(self._folder_edit.text().strip())
            template_file = Path(self._template_edit.text().strip())
            if not data_folder.exists() or not template_file.exists():
                return
            loader = load_local_dataset
            args = (data_folder, template_file)
        else:
            folder_url = self._sharepoint_url_edit.text().strip()
            if not folder_url or not self._cookies:
                return
            loader = self._load_sharepoint_dataset
            args = (folder_url, tuple(self._cookies))

        self._background_thread = RevenueLoaderThread(loader, *args)
        self._background_thread.finished_with_data.connect(self._on_background_refresh_ready)
        self._background_thread.failed.connect(self._on_background_refresh_failed)
        self._background_thread.start()

    def _on_background_refresh_ready(self, dataset: RevenueDataset) -> None:
        refreshed_fp = _dataset_fingerprint(dataset)
        if refreshed_fp != self._current_fingerprint:
            self._apply_dataset(dataset, loaded_from_cache=False, from_background=True)
            self._status.setText("New data detected and loaded in background.")
        else:
            self._status.setText("Background check complete. No new data.")

    def _on_background_refresh_failed(self, message: str) -> None:
        _append_debug_log(f"background_refresh_failed | error={message}")

    def _restore_cached_sharepoint_session(self) -> None:
        folder_url = self._sharepoint_url_edit.text().strip()
        if not folder_url:
            return
        cached = _load_session_cache(folder_url)
        if not cached:
            return
        self._cookies = cached
        self._session_from_cache = True
        self._cookies_label.setText(f"Using cached session ({len(cached)} cookies).")
        self._status.setText("Cached SharePoint session restored.")

    def _open_sharepoint_login(self) -> None:
        folder_url = self._sharepoint_url_edit.text().strip()
        if not folder_url:
            QMessageBox.warning(
                self,
                "SharePoint URL Missing",
                "SharePoint folder URL is not configured. Set RAS_NATIVE_SHAREPOINT_URL before launch.",
            )
            return

        _append_debug_log(f"sharepoint_dialog_opening | pid={os.getpid()} | folder_url={folder_url}")
        dialog = SharePointLoginDialog(folder_url, self)
        dialog_result = dialog.exec()
        accepted_code = int(QDialog.DialogCode.Accepted)
        captured_cookies = dialog.captured_cookies()
        cookie_names = ", ".join(sorted({cookie.name for cookie in captured_cookies})) or "none"
        cookie_domains = ", ".join(sorted({cookie.domain for cookie in captured_cookies if cookie.domain})) or "none"

        _append_debug_log(
            " | ".join(
                [
                    "sharepoint_dialog_result",
                    f"dialog_result={dialog_result}",
                    f"accepted_value={accepted_code}",
                    f"cookie_count={len(captured_cookies)}",
                    f"cookie_names={cookie_names}",
                    f"cookie_domains={cookie_domains}",
                ]
            )
        )

        if not captured_cookies and dialog_result != accepted_code:
            self._cookies = []
            self._cookies_label.setText("No session captured yet.")
            self._status.setText(
                f"SharePoint dialog closed without acceptance. Cookies returned: {len(captured_cookies)} ({cookie_names})"
            )
            return

        self._cookies = captured_cookies
        self._session_from_cache = False
        if not self._cookies:
            QMessageBox.warning(
                self,
                "Session Capture Failed",
                "The sign-in dialog closed without any captured SharePoint cookies.",
            )
            self._cookies_label.setText("No session captured yet.")
            self._status.setText(
                f"SharePoint dialog accepted but returned 0 cookies ({cookie_names})."
            )
            return
        self._cookies_label.setText(f"Captured {len(self._cookies)} cookies in memory.")
        self._status.setText(
            f"SharePoint session captured with {len(self._cookies)} cookies ({cookie_names})."
        )
        _append_debug_log(
            f"sharepoint_session_stored | cookie_count={len(self._cookies)} | cookie_names={cookie_names}"
        )
        _save_session_cache(folder_url, self._cookies)

    def _set_busy(self, busy: bool, message: str = "") -> None:
        self._progress.setVisible(busy)
        self._load_button.setEnabled(not busy)
        if message:
            self._status.setText(message)

    def _load_dashboard(self) -> None:
        if self._local_mode_radio.isChecked():
            data_folder = Path(self._folder_edit.text().strip())
            template_file = Path(self._template_edit.text().strip())
            if not data_folder.exists():
                QMessageBox.warning(self, "Missing folder", "Choose a valid local revenue folder.")
                return
            if not template_file.exists():
                QMessageBox.warning(self, "Missing template", "Choose a valid template file.")
                return

            loader = load_local_dataset
            args = (data_folder, template_file)
        else:
            if not self._cookies:
                QMessageBox.warning(self, "No session", "Open the login browser and capture the SharePoint session first.")
                return

            folder_url = self._sharepoint_url_edit.text().strip()
            if not folder_url:
                QMessageBox.warning(
                    self,
                    "SharePoint URL Missing",
                    "SharePoint folder URL is not configured. Set RAS_NATIVE_SHAREPOINT_URL before launch.",
                )
                return

            loader = self._load_sharepoint_dataset
            args = (folder_url, tuple(self._cookies))

        self._set_busy(True, "Loading revenue data...")
        self._data_thread = RevenueLoaderThread(loader, *args)
        self._data_thread.finished_with_data.connect(self._render_dataset)
        self._data_thread.failed.connect(self._on_load_failed)
        self._data_thread.start()

    def _load_sharepoint_dataset(self, folder_url: str, cookies: tuple[CookieSnapshot, ...]) -> RevenueDataset:
        cookie_records = [
            SharePointCookieRecord(name=cookie.name, value=cookie.value, domain=cookie.domain, path=cookie.path, secure=cookie.secure)
            for cookie in cookies
        ]
        result = download_sharepoint_dataset(folder_url, cookie_records)
        return load_local_dataset(result.data_folder, result.template_file)

    def _on_load_failed(self, message: str) -> None:
        self._set_busy(False, "Load failed.")
        if self._sharepoint_mode_radio.isChecked() and self._session_from_cache:
            _clear_session_cache()
            self._cookies = []
            self._session_from_cache = False
            self._cookies_label.setText("No session captured yet.")
            self._status.setText("Cached SharePoint session expired. Please sign in again.")
        QMessageBox.critical(self, "Load failed", message)

    def _selected_list_values(self, widget: QListWidget) -> list[str]:
        return [item.text() for item in widget.selectedItems() if item.text().strip()]

    def _selected_metrics(self) -> list[str]:
        values = self._selected_list_values(self._metric_list)
        return values or ["GM%"]

    def _clear_filters(self) -> None:
        self._window_combo.setCurrentText("Last 3 months")
        self._summary_window_combo.setCurrentText("Last 3 months")
        self._account_list.clearSelection()
        self._project_list.clearSelection()
        self._account_search_edit.clear()
        self._project_search_edit.clear()
        self._metric_list.clearSelection()
        for idx in range(self._metric_list.count()):
            item = self._metric_list.item(idx)
            if item.text() == "GM%":
                item.setSelected(True)
        self._set_metric_selection(self._summary_metric_list, ["GM%"])
        self._compact_view.setChecked(True)
        self._update_selection_hints()
        self._apply_filters()

    def _build_month_metric_pivot(
        self,
        data: pd.DataFrame,
        index_columns: list[str],
        selected_months: list[pd.Timestamp],
        selected_metrics: list[str],
    ) -> pd.DataFrame:
        if data.empty or not selected_months or not selected_metrics:
            return pd.DataFrame(columns=index_columns)

        working = data.copy()
        working = working.loc[working["month"].isin(selected_months)].copy()
        if working.empty:
            return pd.DataFrame(columns=index_columns)

        grouped = (
            working.groupby(index_columns + ["month"], dropna=False)
            .agg(
                rev_month=("rev_month", "sum"),
                cost_month=("cost_month", "sum"),
            )
            .reset_index()
        )
        grouped["margin_month"] = grouped["rev_month"] - grouped["cost_month"]
        grouped["gm_month_pct"] = (
            grouped["margin_month"] / grouped["rev_month"].where(grouped["rev_month"] != 0) * 100
        ).fillna(0)
        grouped["month_label"] = pd.to_datetime(grouped["month"]).dt.strftime("%b-%Y")

        metric_map = {
            "GM%": ("gm_month_pct", "GM%"),
            "Revenue": ("rev_month", "Revenue"),
            "Cost": ("cost_month", "Cost"),
        }
        month_order = [pd.Timestamp(m).strftime("%b-%Y") for m in sorted(selected_months, reverse=True)]

        result = pd.DataFrame()
        for metric in selected_metrics:
            field, suffix = metric_map[metric]
            pivot = (
                grouped.pivot_table(
                    index=index_columns,
                    columns="month_label",
                    values=field,
                    aggfunc="sum",
                    fill_value=0,
                )
                .reindex(columns=month_order, fill_value=0.0)
                .reset_index()
            )
            rename_cols = {m: f"{m} {suffix}" for m in month_order if m in pivot.columns}
            pivot = pivot.rename(columns=rename_cols)
            if result.empty:
                result = pivot
            else:
                result = result.merge(pivot, on=index_columns, how="outer")

        return result.fillna(0)

    def _format_for_display(self, frame: pd.DataFrame) -> pd.DataFrame:
        if frame.empty:
            return frame
        out = frame.copy()
        if "month" in out.columns:
            out["month"] = pd.to_datetime(out["month"], errors="coerce").dt.strftime("%b-%Y")
        for col in ["start_date", "end_date"]:
            if col in out.columns:
                out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%Y-%m-%d")
        return out

    def _sync_project_options(self) -> None:
        if self._raw_merged.empty:
            return
        selected_accounts = set(self._selected_list_values(self._account_list))
        working = self._raw_merged
        if selected_accounts and "customer_name" in working.columns:
            working = working.loc[working["customer_name"].isin(selected_accounts)].copy()

        projects = sorted({str(value).strip() for value in working.get("project_name", pd.Series(dtype=str)).dropna().tolist() if str(value).strip()})
        self._all_projects = projects
        existing_selected = set(self._selected_list_values(self._project_list))

        self._project_list.blockSignals(True)
        self._project_list.clear()
        for project in projects:
            item = QListWidgetItem(project)
            if project in existing_selected:
                item.setSelected(True)
            self._project_list.addItem(item)
        self._project_list.blockSignals(False)
        self._filter_project_options()
        self._update_selection_hints()

    def _select_month_window(self, months: list[pd.Timestamp], choice: str) -> list[pd.Timestamp]:
        if not months:
            return []

        months_sorted = sorted(pd.Timestamp(value) for value in months)
        current = months_sorted[-1]

        if choice == "Current month":
            return [current]
        if choice == "Last month":
            return months_sorted[-2:-1] or [current]
        if choice == "Last 2 months":
            if len(months_sorted) <= 1:
                return [current]
            return months_sorted[max(0, len(months_sorted) - 3) : -1]
        if choice == "All months":
            return months_sorted
        size_by_choice = {
            "Last 3 months": 3,
            "Last 6 months": 6,
            "Last 12 months": 12,
        }
        size = size_by_choice.get(choice, 3)
        return months_sorted[-size:]

    def _apply_filters(self) -> None:
        if self._raw_merged.empty:
            return

        self._sync_summary_controls_from_main()

        working = self._raw_merged.copy()
        selected_accounts = self._selected_list_values(self._account_list)
        if selected_accounts and "customer_name" in working.columns:
            working = working.loc[working["customer_name"].isin(selected_accounts)].copy()

        selected_projects = self._selected_list_values(self._project_list)
        if selected_projects and "project_name" in working.columns:
            working = working.loc[working["project_name"].isin(selected_projects)].copy()

        month_values = sorted([pd.Timestamp(value) for value in working.get("month", pd.Series(dtype="datetime64[ns]")).dropna().unique()])
        selected_months = self._select_month_window(month_values, self._window_combo.currentText())
        if selected_months and "month" in working.columns:
            working = working.loc[working["month"].isin(selected_months)].copy()

        selected_metrics = self._selected_metrics()

        if working.empty:
            self._kpi_labels["revenue"].setText("$0.00")
            self._kpi_labels["cost"].setText("$0.00")
            self._kpi_labels["margin"].setText("$0.00")
            self._kpi_labels["gm"].setText("0.00%")
            self._snapshot_model.set_frame(pd.DataFrame())
            self._summary_model.set_frame(pd.DataFrame())
            self._trend_view.setHtml("<html><body><h3>No data for selected filters.</h3></body></html>")
            self._filter_summary.setText("No rows match selected filters.")
            return

        revenue = float(working["rev_month"].sum()) if "rev_month" in working.columns else 0.0
        cost = float(working["cost_month"].sum()) if "cost_month" in working.columns else 0.0
        margin = revenue - cost
        gm_pct = (margin / revenue * 100.0) if revenue else 0.0

        self._kpi_labels["revenue"].setText(f"${revenue:,.2f}")
        self._kpi_labels["cost"].setText(f"${cost:,.2f}")
        self._kpi_labels["margin"].setText(f"${margin:,.2f}")
        self._kpi_labels["gm"].setText(f"{gm_pct:,.2f}%")

        if not self._raw_snapshot.empty and {"customer_id", "project_id"}.issubset(working.columns):
            snapshot_filtered = self._raw_snapshot.merge(
                working[["customer_id", "project_id"]].drop_duplicates(),
                on=["customer_id", "project_id"],
                how="inner",
            )
        else:
            snapshot_filtered = pd.DataFrame()

        latest_snapshot = snapshot_filtered.copy()
        base_snapshot_columns = [
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "start_date",
            "end_date",
            "target_gm_pct",
            "projected_margin",
            "cost_increased",
            "revenue_changed",
        ]
        latest_snapshot = latest_snapshot[[col for col in base_snapshot_columns if col in latest_snapshot.columns]]

        project_month_pivot = self._build_month_metric_pivot(
            working,
            index_columns=["customer_id", "project_id"],
            selected_months=selected_months,
            selected_metrics=selected_metrics,
        )
        dashboard = latest_snapshot.merge(project_month_pivot, on=["customer_id", "project_id"], how="left")

        dynamic_columns = [
            col for col in dashboard.columns if any(col.endswith(suffix) for suffix in ["GM%", "Revenue", "Cost"])
        ]
        if dynamic_columns:
            dashboard[dynamic_columns] = dashboard[dynamic_columns].fillna(0)

        if "cost_increased" in dashboard.columns and "revenue_changed" in dashboard.columns:
            dashboard["risk_flag"] = dashboard["cost_increased"].fillna(False) | dashboard["revenue_changed"].fillna(False)
        else:
            dashboard["risk_flag"] = False
        dashboard["risk_status"] = dashboard["risk_flag"].map({True: "Risk", False: "Normal"})

        compact_columns = [
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "target_gm_pct",
            "projected_margin",
            "risk_status",
            "risk_flag",
        ]
        full_columns = [
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "start_date",
            "end_date",
            "target_gm_pct",
            "projected_margin",
            "risk_status",
            "cost_increased",
            "revenue_changed",
            "risk_flag",
        ]
        selected_base = compact_columns if self._compact_view.isChecked() else full_columns
        selected_cols = [col for col in selected_base if col in dashboard.columns] + dynamic_columns
        dashboard = dashboard[selected_cols].sort_values(["customer_name", "project_name"], na_position="last")

        summary_months = self._select_month_window(month_values, self._window_combo.currentText())
        summary_source = self._raw_merged.loc[self._raw_merged["month"].isin(summary_months)].copy()
        summary_wide = self._build_month_metric_pivot(
            summary_source,
            index_columns=["customer_name"],
            selected_months=summary_months,
            selected_metrics=selected_metrics,
        )
        summary_wide = summary_wide.sort_values(["customer_name"], na_position="last").fillna(0)

        self._snapshot_model.set_frame(self._format_for_display(dashboard.reset_index(drop=True)))
        self._summary_model.set_frame(self._format_for_display(summary_wide.reset_index(drop=True)))

        monthly = (
            working.groupby(["month", "customer_name"], dropna=False)[["rev_month", "cost_month"]]
            .sum()
            .reset_index()
            .sort_values("month")
        )
        if monthly.empty:
            self._trend_view.setHtml("<html><body><h3>No trend data available.</h3></body></html>")
        else:
            monthly["margin_month"] = monthly["rev_month"] - monthly["cost_month"]
            field_map = {"Revenue": "rev_month", "Cost": "cost_month", "Margin": "margin_month"}
            metric_field = field_map[self._trend_metric_combo.currentText()]
            fig = px.line(
                monthly,
                x="month",
                y=metric_field,
                color="customer_name",
                markers=True,
                title=f"{self._trend_metric_combo.currentText()} Trend by Customer",
            )
            fig.update_layout(template="plotly_white", legend_title_text="Customer")
            html = pio.to_html(fig, include_plotlyjs="cdn", full_html=False)
            self._trend_view.setHtml(html)

        self._filter_summary.setText(
            f"Rows: {len(working)} | Accounts: {len(selected_accounts) or 'All'} | Projects: {len(selected_projects) or 'All'} | Months: {len(selected_months) or len(month_values)}"
        )

    def _render_dataset(self, dataset: RevenueDataset) -> None:
        self._apply_dataset(dataset, loaded_from_cache=False)

    def _apply_dataset(self, dataset: RevenueDataset, loaded_from_cache: bool = False, from_background: bool = False) -> None:
        self._current_dataset = dataset
        self._current_fingerprint = _dataset_fingerprint(dataset)
        self._set_busy(False, "Data loaded successfully." if not loaded_from_cache else "Loaded cached data.")

        self._raw_merged = dataset.merged.copy()
        self._raw_snapshot = dataset.snapshot.copy()

        accounts = sorted(
            {
                str(value).strip()
                for value in self._raw_merged.get("customer_name", pd.Series(dtype=str)).dropna().tolist()
                if str(value).strip()
            }
        )
        self._account_list.blockSignals(True)
        self._account_list.clear()
        self._all_accounts = accounts
        for account in accounts:
            self._account_list.addItem(account)
        self._account_list.blockSignals(False)
        self._filter_account_options()

        self._sync_project_options()
        self._update_selection_hints()
        self._apply_filters()

        if not from_background:
            self._toggle_inputs_button.setText("Show Inputs")
            self._inputs_visible = False
            self._mode_stack.setVisible(False)
            self._load_button.setVisible(False)

        _save_dataset_cache(
            dataset,
            {
                "source_key": self._cache_source_key(),
                "source_signature": self._cache_source_signature(),
                "mode": "local" if self._local_mode_radio.isChecked() else "sharepoint",
            },
        )
        if not self._refresh_timer.isActive():
            self._refresh_timer.start()


def run_app() -> int:
    app = QApplication([])
    if not _acquire_single_instance_lock():
        QMessageBox.warning(
            None,
            "Revenue Analysis Desktop",
            "Another Revenue Analysis Desktop window is already running. Close it first, then relaunch.",
        )
        return 1

    _reset_debug_log()
    _append_debug_log(f"app_started | pid={os.getpid()} | temp={tempfile.gettempdir()}")
    window = MainWindow()
    window.show()
    window.raise_()
    window.activateWindow()
    return app.exec()
