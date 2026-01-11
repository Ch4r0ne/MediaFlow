from __future__ import annotations

import os
import sys
import threading
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import List, Optional, Set, Tuple

from PyQt6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    QSettings,
    Qt,
    QThread,
    pyqtSignal,
)
from PyQt6.QtGui import QColor, QFont, QIcon, QPalette, QStandardItem, QStandardItemModel
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QSpinBox,
    QDoubleSpinBox,
    QTableView,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

# --- Optional deps (graceful) ---
try:
    import cv2  # type: ignore
except Exception:
    cv2 = None

try:
    from send2trash import send2trash  # type: ignore
except Exception:
    send2trash = None

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
except Exception:
    pythoncom = None
    win32com = None


# =========================
# App Identity
# =========================
APP_NAME = "MediaFlow"
APP_ORG = "TimTools"
APP_SETTINGS = "MediaFlow_BusinessClean"


def resource_path(relative: str) -> str:
    """
    PyInstaller-friendly asset loader:
    - dev: relative path from cwd
    - onefile: from sys._MEIPASS
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return str(Path(base) / relative)
    return str(Path(relative))


# =========================
# Theme (Business Dark)
# =========================
def apply_business_dark(app: QApplication) -> None:
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor(24, 24, 24))
    pal.setColor(QPalette.ColorRole.WindowText, QColor(235, 235, 235))
    pal.setColor(QPalette.ColorRole.Base, QColor(18, 18, 18))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor(28, 28, 28))
    pal.setColor(QPalette.ColorRole.Text, QColor(235, 235, 235))
    pal.setColor(QPalette.ColorRole.Button, QColor(42, 42, 42))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor(235, 235, 235))
    pal.setColor(QPalette.ColorRole.Highlight, QColor(90, 90, 90))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(pal)

    app.setStyleSheet(
        """
        QWidget { font-size: 10pt; }
        QGroupBox { font-weight: 600; }
        QLineEdit, QTableView, QProgressBar, QTreeWidget, QComboBox, QDoubleSpinBox, QSpinBox {
            border: 1px solid #3a3a3a;
            border-radius: 8px;
        }
        QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox { padding: 6px; }
        QHeaderView::section {
            background: #2a2a2a;
            border: 1px solid #3a3a3a;
            padding: 6px;
            font-weight: 700;
        }
        QPushButton, QToolButton {
            padding: 7px 12px;
            border: 1px solid #3a3a3a;
            border-radius: 8px;
            background: #2b2b2b;
            font-weight: 650;
        }
        QPushButton:hover, QToolButton:hover { border: 1px solid #5a5a5a; background: #303030; }
        QPushButton:pressed, QToolButton:pressed { background: #262626; }
        QPushButton:disabled, QToolButton:disabled { color: #9a9a9a; background: #242424; border: 1px solid #2f2f2f; }
        QProgressBar { text-align: center; }
        QProgressBar::chunk { background-color: #6a6a6a; }
        QToolTip {
            color: #f0f0f0;
            background-color: #202020;
            border: 1px solid #4a4a4a;
            padding: 6px;
        }
        QListWidget {
            border: 1px solid #3a3a3a;
            border-radius: 10px;
            padding: 6px;
        }
        """
    )


# =========================
# Shared helpers
# =========================
def sanitize_folder_name(name: str) -> str:
    name = (name or "").strip()
    name = name.replace("/", "_").replace("\\", "_").replace("..", "_")
    return name


def compute_output_root(source_dir: Path, output_name: str) -> Path:
    out = sanitize_folder_name(output_name)
    return source_dir if out == "" else (source_dir / out)


def is_under(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def enumerate_files(source_dir: Path, recursive: bool, exclude_dirs: List[Path]) -> List[Path]:
    def allowed(p: Path) -> bool:
        for ex in exclude_dirs:
            if is_under(p, ex):
                return False
        return True

    if not recursive:
        return [p for p in source_dir.iterdir() if p.is_file() and allowed(p)]

    out: List[Path] = []
    for p in source_dir.rglob("*"):
        if not p.is_file():
            continue
        if allowed(p):
            out.append(p)
    return out


def unique_dest(dest: Path) -> Path:
    if not dest.exists():
        return dest
    stem, ext = dest.stem, dest.suffix
    parent = dest.parent
    i = 1
    while True:
        cand = parent / f"{stem} ({i}){ext}"
        if not cand.exists():
            return cand
        i += 1


def move_file(src: Path, dest: Path, overwrite: bool) -> None:
    if overwrite and dest.exists() and dest.is_file():
        dest.unlink()
    try:
        src.rename(dest)
    except OSError:
        import shutil

        if overwrite and dest.exists() and dest.is_file():
            dest.unlink()
        shutil.move(str(src), str(dest))


# =========================
# Sort Media (merged: AspectRatioSorter + ImageVideoSorter)
# =========================
class SortMode(str, Enum):
    ORIENTATION = "Orientation (Portrait / Landscape)"
    TYPE = "Type (Images / Videos)"


SUPPORTED_IMAGE_EXT_TYPE = {
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff",
    ".heic", ".heif",
    ".raw", ".cr2", ".nef", ".arw", ".dng",
}
SUPPORTED_VIDEO_EXT_TYPE = {
    ".mp4", ".mov", ".m4v", ".mkv", ".avi", ".wmv", ".webm",
    ".mpg", ".mpeg", ".3gp", ".flv",
}

# For ORIENTATION we keep a practical subset where cv2 is most likely to work.
SUPPORTED_IMAGE_EXT_ORIENT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
SUPPORTED_VIDEO_EXT_ORIENT = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".wmv", ".webm", ".mpg", ".mpeg"}


def classify_type(p: Path) -> Optional[str]:
    ext = p.suffix.lower()
    if ext in SUPPORTED_IMAGE_EXT_TYPE:
        return "image"
    if ext in SUPPORTED_VIDEO_EXT_TYPE:
        return "video"
    return None


def classify_dimensions(p: Path) -> Tuple[str, int, int]:
    """
    Returns: (kind, width, height)
    kind: image|video
    """
    if cv2 is None:
        raise RuntimeError("OpenCV (cv2) not installed. Orientation mode requires opencv-python.")

    ext = p.suffix.lower()

    if ext in SUPPORTED_VIDEO_EXT_ORIENT:
        cap = cv2.VideoCapture(str(p))
        try:
            if not cap.isOpened():
                raise RuntimeError("Could not open video.")
            w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            if w <= 0 or h <= 0:
                raise RuntimeError(f"Invalid video dimensions: {w}x{h}")
            return "video", w, h
        finally:
            cap.release()

    if ext in SUPPORTED_IMAGE_EXT_ORIENT:
        im = cv2.imread(str(p))
        if im is None:
            raise RuntimeError("Could not read image.")
        h, w = im.shape[:2]
        if w <= 0 or h <= 0:
            raise RuntimeError(f"Invalid image dimensions: {w}x{h}")
        return "image", w, h

    raise RuntimeError("Unsupported format for Orientation mode.")


def orientation_bucket(w: int, h: int) -> str:
    return "portrait" if (w / h) < 1 else "landscape"


@dataclass(frozen=True)
class SortConfig:
    source_dir: Path
    output_name: str
    recursive: bool
    lowercase: bool
    dry_run: bool
    dup_mode: str  # auto_rename | skip | overwrite
    remember_settings: bool
    sort_mode: SortMode


@dataclass(frozen=True)
class SortPreviewItem:
    src: Path
    kind: str                 # image | video | ?
    width: int                # 0 if n/a
    height: int               # 0 if n/a
    bucket: str               # portrait|landscape|Images|Videos|?
    dest: Path
    status: str               # OK | SKIP.. | ERROR..


@dataclass
class SortStats:
    found: int = 0
    supported: int = 0
    images: int = 0
    videos: int = 0
    portrait: int = 0
    landscape: int = 0
    skipped_unsupported: int = 0
    skipped_duplicates: int = 0
    errors: int = 0
    moved: int = 0


class SortPreviewModel(QAbstractTableModel):
    HEADERS = ["File", "Type", "WxH", "Bucket", "Destination", "Status"]

    def __init__(self) -> None:
        super().__init__()
        self._items: List[SortPreviewItem] = []

    def set_items(self, items: List[SortPreviewItem]) -> None:
        self.beginResetModel()
        self._items = items
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._items)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self.HEADERS[section]
        return str(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        it = self._items[index.row()]
        col = index.column()

        if role == Qt.ItemDataRole.DisplayRole:
            if col == 0:
                return it.src.name
            if col == 1:
                return it.kind
            if col == 2:
                return "-" if it.width <= 0 or it.height <= 0 else f"{it.width}x{it.height}"
            if col == 3:
                return it.bucket
            if col == 4:
                return "-" if str(it.dest) == "-" else f"{it.dest.parent.name}\\{it.dest.name}"
            if col == 5:
                return it.status

        if role == Qt.ItemDataRole.ForegroundRole and col == 5:
            if it.status.startswith("ERROR"):
                return QColor(220, 130, 130)
            if it.status.startswith("SKIP"):
                return QColor(170, 170, 170)
            return QColor(200, 200, 200)

        return None


class SortAnalyzerWorker(QObject):
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(list, object)
    failed = pyqtSignal(str)

    def __init__(self, cfg: SortConfig, cancel_event: threading.Event):
        super().__init__()
        self.cfg = cfg
        self.cancel_event = cancel_event

    def _ensure_output_dirs(self, output_root: Path) -> Tuple[Path, Path]:
        if self.cfg.sort_mode == SortMode.ORIENTATION:
            a = output_root / "portrait"
            b = output_root / "landscape"
        else:
            a = output_root / "Images"
            b = output_root / "Videos"
        a.mkdir(parents=True, exist_ok=True)
        b.mkdir(parents=True, exist_ok=True)
        return a, b

    def run(self) -> None:
        src = self.cfg.source_dir
        if not src.exists() or not src.is_dir():
            self.failed.emit("Source does not exist or is not a folder.")
            return

        out_root = compute_output_root(src, self.cfg.output_name)
        dir_a, dir_b = self._ensure_output_dirs(out_root)

        exclude_dirs: List[Path] = []
        if out_root != src:
            exclude_dirs.append(out_root)
        exclude_dirs.append(dir_a)
        exclude_dirs.append(dir_b)

        files = enumerate_files(src, self.cfg.recursive, exclude_dirs)
        stats = SortStats(found=len(files))
        preview: List[SortPreviewItem] = []

        self.progress.emit(0, stats.found)
        done = 0

        for p in files:
            if self.cancel_event.is_set():
                break

            try:
                if self.cfg.sort_mode == SortMode.TYPE:
                    kind = classify_type(p)
                    if kind is None:
                        stats.skipped_unsupported += 1
                        done += 1
                        self.progress.emit(done, stats.found)
                        continue

                    stats.supported += 1
                    if kind == "image":
                        stats.images += 1
                        bucket = "Images"
                        dest_dir = dir_a
                    else:
                        stats.videos += 1
                        bucket = "Videos"
                        dest_dir = dir_b

                    out_name = p.name.lower() if self.cfg.lowercase else p.name
                    dest = dest_dir / out_name

                    if dest.exists():
                        if self.cfg.dup_mode == "skip":
                            stats.skipped_duplicates += 1
                            preview.append(SortPreviewItem(p, kind, 0, 0, bucket, dest, "SKIP (duplicate)"))
                            done += 1
                            self.progress.emit(done, stats.found)
                            continue
                        if self.cfg.dup_mode == "auto_rename":
                            dest = unique_dest(dest)

                    status = "OK"
                    if self.cfg.dup_mode == "overwrite" and (dest_dir / out_name).exists():
                        status = "OK (overwrite)"

                    preview.append(SortPreviewItem(p, kind, 0, 0, bucket, dest, status))

                else:
                    # ORIENTATION
                    ext = p.suffix.lower()
                    if ext not in SUPPORTED_IMAGE_EXT_ORIENT and ext not in SUPPORTED_VIDEO_EXT_ORIENT:
                        stats.skipped_unsupported += 1
                        done += 1
                        self.progress.emit(done, stats.found)
                        continue

                    kind, w, h = classify_dimensions(p)
                    stats.supported += 1

                    if kind == "image":
                        stats.images += 1
                    else:
                        stats.videos += 1

                    bucket = orientation_bucket(w, h)
                    if bucket == "portrait":
                        stats.portrait += 1
                        dest_dir = dir_a
                    else:
                        stats.landscape += 1
                        dest_dir = dir_b

                    out_name = p.name.lower() if self.cfg.lowercase else p.name
                    dest = dest_dir / out_name

                    if dest.exists():
                        if self.cfg.dup_mode == "skip":
                            stats.skipped_duplicates += 1
                            preview.append(SortPreviewItem(p, kind, w, h, bucket, dest, "SKIP (duplicate)"))
                            done += 1
                            self.progress.emit(done, stats.found)
                            continue
                        if self.cfg.dup_mode == "auto_rename":
                            dest = unique_dest(dest)

                    status = "OK"
                    if self.cfg.dup_mode == "overwrite" and (dest_dir / out_name).exists():
                        status = "OK (overwrite)"

                    preview.append(SortPreviewItem(p, kind, w, h, bucket, dest, status))

            except Exception as e:
                stats.errors += 1
                preview.append(SortPreviewItem(p, "?", 0, 0, "?", Path("-"), f"ERROR: {e}"))

            done += 1
            self.progress.emit(done, stats.found)

        self.finished.emit(preview, stats)


class SortExecuteWorker(QObject):
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, cfg: SortConfig, items: List[SortPreviewItem], cancel_event: threading.Event):
        super().__init__()
        self.cfg = cfg
        self.items = items
        self.cancel_event = cancel_event

    def _ensure_output_dirs(self, output_root: Path) -> Tuple[Path, Path]:
        if self.cfg.sort_mode == SortMode.ORIENTATION:
            a = output_root / "portrait"
            b = output_root / "landscape"
        else:
            a = output_root / "Images"
            b = output_root / "Videos"
        a.mkdir(parents=True, exist_ok=True)
        b.mkdir(parents=True, exist_ok=True)
        return a, b

    def run(self) -> None:
        ok_items = [x for x in self.items if x.status.startswith("OK")]
        if not ok_items:
            self.failed.emit("No executable items. Run Analyze first.")
            return

        src = self.cfg.source_dir
        out_root = compute_output_root(src, self.cfg.output_name)
        dir_a, dir_b = self._ensure_output_dirs(out_root)

        total = len(ok_items)
        stats = SortStats()
        self.progress.emit(0, total)

        moved = 0
        for it in ok_items:
            if self.cancel_event.is_set():
                break

            if self.cfg.sort_mode == SortMode.ORIENTATION:
                dest_dir = dir_a if it.bucket == "portrait" else dir_b
            else:
                dest_dir = dir_a if it.bucket == "Images" else dir_b

            out_name = it.src.name.lower() if self.cfg.lowercase else it.src.name
            dest = dest_dir / out_name

            if dest.exists():
                if self.cfg.dup_mode == "skip":
                    stats.skipped_duplicates += 1
                    continue
                if self.cfg.dup_mode == "auto_rename":
                    dest = unique_dest(dest)

            if not self.cfg.dry_run:
                try:
                    move_file(it.src, dest, overwrite=(self.cfg.dup_mode == "overwrite"))
                except Exception:
                    stats.errors += 1
                    continue

            moved += 1
            stats.moved = moved
            self.progress.emit(moved, total)

        self.finished.emit(stats)


class SortMediaPage(QWidget):
    def __init__(self, settings: QSettings):
        super().__init__()
        self.settings = settings
        self.cancel_event = threading.Event()

        self.thread: Optional[QThread] = None
        self.worker: Optional[QObject] = None

        self.preview_items: List[SortPreviewItem] = []
        self.preview_cfg: Optional[SortConfig] = None

        self.model = SortPreviewModel()

        self._build_ui()
        self._load()
        self._update_tree()
        self._set_stats(SortStats())
        self._refresh()

    def _build_ui(self) -> None:
        main = QVBoxLayout(self)

        header = QFrame()
        hl = QHBoxLayout(header)
        title = QLabel("Sort Media")
        tf = QFont()
        tf.setPointSize(12)
        tf.setBold(True)
        title.setFont(tf)

        micro = QLabel("Analyze → Preview → Execute")
        micro.setStyleSheet("color: #cfcfcf;")
        micro.setToolTip("Analyze builds the preview table. Execute moves files using the preview configuration.")
        hl.addWidget(title)
        hl.addStretch(1)
        hl.addWidget(micro)
        main.addWidget(header)

        setup = QGroupBox("Setup")
        gl = QGridLayout(setup)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem(SortMode.ORIENTATION.value, SortMode.ORIENTATION)
        self.mode_combo.addItem(SortMode.TYPE.value, SortMode.TYPE)
        self.mode_combo.currentIndexChanged.connect(self._on_any_change)
        self.mode_combo.setToolTip("Choose how you want to sort media files.")

        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("Source folder")
        self.source_edit.setClearButtonEnabled(True)
        self.source_edit.setToolTip("Folder that contains your media.")
        self.source_edit.textChanged.connect(self._on_any_change)

        self.btn_browse = QPushButton("Browse…")
        self.btn_browse.clicked.connect(self._browse)
        self.btn_browse.setToolTip("Select the source folder.")

        self.btn_open_src = QPushButton("Open")
        self.btn_open_src.clicked.connect(self._open_source)
        self.btn_open_src.setToolTip("Open the source folder in Explorer.")

        self.output_edit = QLineEdit("")
        self.output_edit.setPlaceholderText("Output folder (optional)")
        self.output_edit.setToolTip("Leave empty to write into the source folder (creates destination folders there).")
        self.output_edit.textChanged.connect(self._on_any_change)

        for b in (self.btn_browse, self.btn_open_src):
            b.setMinimumHeight(34)
            b.setFixedWidth(110)

        gl.addWidget(QLabel("Sort mode"), 0, 0)
        gl.addWidget(self.mode_combo, 0, 1, 1, 3)

        gl.addWidget(QLabel("Source"), 1, 0)
        gl.addWidget(self.source_edit, 1, 1)
        gl.addWidget(self.btn_browse, 1, 2)
        gl.addWidget(self.btn_open_src, 1, 3)

        gl.addWidget(QLabel("Output"), 2, 0)
        gl.addWidget(self.output_edit, 2, 1, 1, 3)

        main.addWidget(setup)

        body = QHBoxLayout()
        main.addLayout(body, 1)

        left = QVBoxLayout()
        body.addLayout(left, 1)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setToolTip("Folder structure that will be created.")
        left.addWidget(self.tree, 1)

        options = QGroupBox("Options")
        ol = QGridLayout(options)

        self.cb_dry = QCheckBox("Dry-run")
        self.cb_dry.setToolTip("No files are moved. Use this to validate the preview.")
        self.cb_dry.stateChanged.connect(self._on_any_change)

        self.cb_lower = QCheckBox("Lower-case filenames")
        self.cb_lower.setToolTip("Renames moved files to lower-case.")
        self.cb_lower.setChecked(True)
        self.cb_lower.stateChanged.connect(self._on_any_change)

        self.cb_recursive = QCheckBox("Recursive")
        self.cb_recursive.setToolTip("Include subfolders. Destination folders are excluded automatically.")
        self.cb_recursive.stateChanged.connect(self._on_any_change)

        self.cb_remember = QCheckBox("Remember settings")
        self.cb_remember.setToolTip("Stores your selections for the next start.")
        self.cb_remember.setChecked(True)
        self.cb_remember.stateChanged.connect(self._save)

        self.dup_combo = QComboBox()
        self.dup_combo.addItem("Duplicates: Auto-rename", "auto_rename")
        self.dup_combo.addItem("Duplicates: Skip", "skip")
        self.dup_combo.addItem("Duplicates: Overwrite", "overwrite")
        self.dup_combo.setToolTip(
            "Auto-rename: keeps both files (adds a suffix).\n"
            "Skip: leaves existing destination file unchanged.\n"
            "Overwrite: replaces existing destination file."
        )
        self.dup_combo.currentIndexChanged.connect(self._on_any_change)

        ol.addWidget(self.cb_dry, 0, 0)
        ol.addWidget(self.cb_lower, 0, 1)
        ol.addWidget(self.cb_recursive, 1, 0)
        ol.addWidget(self.cb_remember, 1, 1)
        ol.addWidget(self.dup_combo, 2, 0, 1, 2)
        left.addWidget(options)

        actions = QGroupBox("Actions")
        al = QHBoxLayout(actions)

        self.btn_analyze = QPushButton("Analyze")
        self.btn_execute = QPushButton("Execute")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setEnabled(False)

        for b in (self.btn_analyze, self.btn_execute, self.btn_cancel):
            b.setMinimumHeight(34)
            b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.btn_analyze.clicked.connect(self._analyze)
        self.btn_execute.clicked.connect(self._execute)
        self.btn_cancel.clicked.connect(self._cancel)

        al.addWidget(self.btn_analyze)
        al.addWidget(self.btn_execute)
        al.addWidget(self.btn_cancel)
        left.addWidget(actions)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        left.addWidget(self.progress)

        self.stats = QLabel("Ready.")
        self.stats.setStyleSheet("color: #cfcfcf;")
        left.addWidget(self.stats)

        right = QVBoxLayout()
        body.addLayout(right, 2)

        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        right.addWidget(self.table, 1)

        # Drag & drop
        self.setAcceptDrops(True)

    def _mode(self) -> SortMode:
        m = self.mode_combo.currentData()
        return m if isinstance(m, SortMode) else SortMode.TYPE

    def _cfg(self) -> SortConfig:
        return SortConfig(
            source_dir=Path(self.source_edit.text().strip()),
            output_name=self.output_edit.text(),
            recursive=self.cb_recursive.isChecked(),
            lowercase=self.cb_lower.isChecked(),
            dry_run=self.cb_dry.isChecked(),
            dup_mode=str(self.dup_combo.currentData()),
            remember_settings=self.cb_remember.isChecked(),
            sort_mode=self._mode(),
        )

    def _invalidate_preview(self) -> None:
        self.preview_items = []
        self.preview_cfg = None
        self.model.set_items([])
        self.progress.setValue(0)
        self._set_stats(SortStats())

    def _on_any_change(self) -> None:
        if self.thread is None:
            self._invalidate_preview()
        self._save()
        self._update_tree()
        self._refresh()

    def _load(self) -> None:
        self.source_edit.setText(self.settings.value("sort/source", "", type=str))
        self.output_edit.setText(self.settings.value("sort/output", "", type=str))
        self.cb_dry.setChecked(self.settings.value("sort/dry", False, type=bool))
        self.cb_lower.setChecked(self.settings.value("sort/lower", True, type=bool))
        self.cb_recursive.setChecked(self.settings.value("sort/recursive", False, type=bool))
        self.cb_remember.setChecked(self.settings.value("sort/remember", True, type=bool))

        dup = self.settings.value("sort/dup", "auto_rename", type=str)
        idx = self.dup_combo.findData(dup)
        if idx >= 0:
            self.dup_combo.setCurrentIndex(idx)

        mode = self.settings.value("sort/mode", SortMode.ORIENTATION.value, type=str)
        if mode == SortMode.TYPE.value:
            self.mode_combo.setCurrentIndex(1)
        else:
            self.mode_combo.setCurrentIndex(0)

    def _save(self) -> None:
        if not self.cb_remember.isChecked():
            return
        self.settings.setValue("sort/source", self.source_edit.text().strip())
        self.settings.setValue("sort/output", self.output_edit.text())
        self.settings.setValue("sort/dry", self.cb_dry.isChecked())
        self.settings.setValue("sort/lower", self.cb_lower.isChecked())
        self.settings.setValue("sort/recursive", self.cb_recursive.isChecked())
        self.settings.setValue("sort/remember", self.cb_remember.isChecked())
        self.settings.setValue("sort/dup", self.dup_combo.currentData())
        self.settings.setValue("sort/mode", self._mode().value)

    def _update_tree(self) -> None:
        self.tree.clear()
        src_txt = self.source_edit.text().strip()
        src_label = src_txt if src_txt else "Source"

        out_name = sanitize_folder_name(self.output_edit.text())
        use_source_as_output = (out_name == "")

        mode = self._mode()
        if mode == SortMode.ORIENTATION:
            a, b = "portrait", "landscape"
        else:
            a, b = "Images", "Videos"

        src_item = QTreeWidgetItem([src_label])
        self.tree.addTopLevelItem(src_item)

        if use_source_as_output:
            src_item.addChild(QTreeWidgetItem([a]))
            src_item.addChild(QTreeWidgetItem([b]))
        else:
            out_item = QTreeWidgetItem([out_name])
            src_item.addChild(out_item)
            out_item.addChild(QTreeWidgetItem([a]))
            out_item.addChild(QTreeWidgetItem([b]))

        self.tree.expandAll()

    def _set_stats(self, s: SortStats) -> None:
        skipped = s.skipped_unsupported + s.skipped_duplicates
        mode = self._mode()
        if mode == SortMode.ORIENTATION:
            self.stats.setText(
                f"Found: {s.found}  |  Supported: {s.supported}  |  Images: {s.images}  |  Videos: {s.videos}  |  "
                f"Portrait: {s.portrait}  |  Landscape: {s.landscape}  |  Skipped: {skipped}  |  Errors: {s.errors}  |  Moved: {s.moved}"
            )
        else:
            self.stats.setText(
                f"Found: {s.found}  |  Supported: {s.supported}  |  Images: {s.images}  |  Videos: {s.videos}  |  "
                f"Skipped: {skipped}  |  Errors: {s.errors}  |  Moved: {s.moved}"
            )

    def _refresh(self) -> None:
        running = self.thread is not None
        has_source_text = bool(self.source_edit.text().strip())
        has_preview = len(self.preview_items) > 0 and self.preview_cfg is not None

        self.btn_open_src.setEnabled(has_source_text and not running)
        self.btn_analyze.setEnabled(has_source_text and not running)
        self.btn_execute.setEnabled(has_preview and not running)
        self.btn_cancel.setEnabled(running)

        for w in (
            self.mode_combo, self.source_edit, self.output_edit, self.cb_dry, self.cb_lower,
            self.cb_recursive, self.cb_remember, self.dup_combo, self.btn_browse
        ):
            w.setEnabled(not running)

    def _browse(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Select Source Folder", self.source_edit.text().strip() or str(Path.home()))
        if d:
            self.source_edit.setText(d)

    def _open_source(self) -> None:
        p = Path(self.source_edit.text().strip())
        if p.exists() and p.is_dir():
            try:
                os.startfile(str(p))  # Windows
            except Exception:
                from PyQt6.QtCore import QUrl
                from PyQt6.QtGui import QDesktopServices
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))

    def _analyze(self) -> None:
        if self.thread is not None:
            return

        cfg = self._cfg()
        if not cfg.source_dir.exists() or not cfg.source_dir.is_dir():
            QMessageBox.warning(self, "Invalid Source", "Select a valid source folder.")
            return

        if cfg.sort_mode == SortMode.ORIENTATION and cv2 is None:
            QMessageBox.critical(self, "Missing dependency", "Orientation mode requires OpenCV. Install: py -m pip install opencv-python")
            return

        self._save()
        self.progress.setValue(0)

        self.preview_items = []
        self.preview_cfg = None
        self.model.set_items([])
        self._set_stats(SortStats())

        self.cancel_event.clear()
        self.thread = QThread()
        self.worker = SortAnalyzerWorker(cfg, self.cancel_event)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.failed.connect(self._on_failed)
        self.worker.finished.connect(self._on_analyze_finished)

        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._cleanup)

        self.thread.start()
        self._refresh()

    def _execute(self) -> None:
        if self.thread is not None:
            return
        if not self.preview_items or self.preview_cfg is None:
            return

        cfg = self.preview_cfg
        out_root = compute_output_root(cfg.source_dir, cfg.output_name)

        mode = cfg.sort_mode
        if mode == SortMode.ORIENTATION:
            msg = f"Move files into:\n{out_root}\\portrait and {out_root}\\landscape\n\nContinue?"
        else:
            msg = f"Move files into:\n{out_root}\\Images and {out_root}\\Videos\n\nContinue?"

        if not cfg.dry_run:
            res = QMessageBox.question(self, "Confirm", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if res != QMessageBox.StandardButton.Yes:
                return

        self.progress.setValue(0)

        self.cancel_event.clear()
        self.thread = QThread()
        self.worker = SortExecuteWorker(cfg, self.preview_items, self.cancel_event)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.failed.connect(self._on_failed)
        self.worker.finished.connect(self._on_execute_finished)

        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self._cleanup)

        self.thread.start()
        self._refresh()

    def _cancel(self) -> None:
        if self.thread is None:
            return
        self.cancel_event.set()

    def _on_progress(self, done: int, total: int) -> None:
        if total <= 0:
            self.progress.setValue(0)
            return
        self.progress.setValue(max(0, min(100, int((done / total) * 100))))

    def _on_failed(self, msg: str) -> None:
        QMessageBox.critical(self, "Error", msg)

    def _on_analyze_finished(self, preview: list, stats: object) -> None:
        self.preview_items = list(preview)
        self.preview_cfg = self._cfg()
        self.model.set_items(self.preview_items)
        self.table.resizeColumnsToContents()
        if isinstance(stats, SortStats):
            self._set_stats(stats)
        self._refresh()

    def _on_execute_finished(self, stats: object) -> None:
        if isinstance(stats, SortStats):
            self._set_stats(stats)
        if self.preview_cfg and self.preview_cfg.dry_run:
            QMessageBox.information(self, "Dry-run", "Dry-run finished. No files were moved.")
        self._refresh()

    def _cleanup(self) -> None:
        self.thread = None
        self.worker = None
        self._refresh()

    # Drag & drop folder
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        p = Path(urls[0].toLocalFile())
        if p.is_dir():
            self.source_edit.setText(str(p))


# =========================
# Short Video Cleaner (merged: DeleteShortVideos)
# =========================
DEFAULT_EXTS = "mp4,mov,mkv,avi,wmv,webm,m4v,mts,m2ts"


class ActionMode(str, Enum):
    ANALYZE = "Dry-Run (no deletion)"
    RECYCLE = "Move to Recycle Bin"
    HARD_DELETE = "Delete permanently"


@dataclass(frozen=True)
class CleanerSettings:
    directory: Path
    recursive: bool
    threshold_seconds: float
    extensions: Set[str]
    action: ActionMode


@dataclass
class CleanerStats:
    found: int = 0
    scanned: int = 0
    short: int = 0
    deleted: int = 0
    kept: int = 0
    unknown: int = 0
    errors: int = 0


def parse_extensions(text: str) -> Set[str]:
    raw = (text or "").replace(";", ",").replace(" ", "")
    parts = [p for p in raw.split(",") if p]
    out: Set[str] = set()
    for p in parts:
        p = p.strip().lower()
        if p.startswith("."):
            p = p[1:]
        if p:
            out.add(p)
    return out


def iter_files_fast(root: Path, recursive: bool):
    stack = [root]
    while stack:
        cur = stack.pop()
        try:
            with os.scandir(cur) as it:
                for e in it:
                    try:
                        if e.is_dir(follow_symlinks=False):
                            if recursive:
                                stack.append(Path(e.path))
                        elif e.is_file(follow_symlinks=False):
                            yield Path(e.path)
                    except OSError:
                        continue
        except OSError:
            continue


class WindowsDurationReader:
    def __init__(self) -> None:
        if not sys.platform.startswith("win"):
            raise RuntimeError("Short Video Cleaner is Windows-only.")
        if pythoncom is None or win32com is None:
            raise RuntimeError("Missing dependency: pywin32. Install: py -m pip install pywin32")

        self._shell = win32com.client.Dispatch("Shell.Application")
        self._folder_cache = {}

    def duration_seconds(self, file_path: Path) -> Optional[float]:
        folder_path = str(file_path.parent)
        folder = self._folder_cache.get(folder_path)
        if folder is None:
            folder = self._shell.NameSpace(folder_path)
            self._folder_cache[folder_path] = folder
        if folder is None:
            return None

        item = folder.ParseName(file_path.name)
        if item is None:
            return None

        try:
            v = item.ExtendedProperty("System.Media.Duration")
            if v is None:
                return None
            ticks_100ns = int(v)
            if ticks_100ns <= 0:
                return None
            return ticks_100ns / 10_000_000.0
        except Exception:
            return None


class CleanerWorker(QObject):
    progress = pyqtSignal(int, int)
    status = pyqtSignal(str)
    stats = pyqtSignal(object)
    row = pyqtSignal(object)  # (status, duration, filename, fullpath)
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, settings: CleanerSettings):
        super().__init__()
        self.s = settings
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    def _delete(self, p: Path) -> None:
        if self.s.action == ActionMode.ANALYZE:
            return
        if self.s.action == ActionMode.RECYCLE:
            if send2trash is None:
                raise RuntimeError("send2trash not installed")
            send2trash(str(p))  # type: ignore[misc]
            return
        p.unlink(missing_ok=True)

    def run(self) -> None:
        try:
            if not sys.platform.startswith("win"):
                raise RuntimeError("Short Video Cleaner is Windows-only.")
            if pythoncom is None or win32com is None:
                raise RuntimeError("pywin32 is missing. Install: py -m pip install pywin32")

            pythoncom.CoInitialize()
            try:
                reader = WindowsDurationReader()
                exts = self.s.extensions

                files: List[Path] = []
                for f in iter_files_fast(self.s.directory, self.s.recursive):
                    ext = f.suffix.lower().lstrip(".")
                    if ext in exts:
                        files.append(f)

                st = CleanerStats(found=len(files))
                self.stats.emit(st)

                total = len(files)
                self.status.emit(f"Found {total} files")

                for i, p in enumerate(files, start=1):
                    if self._cancel:
                        self.status.emit("Canceled")
                        break

                    st.scanned += 1
                    dur = reader.duration_seconds(p)

                    if dur is None:
                        st.unknown += 1
                        self.row.emit(("UNKNOWN", None, p.name, str(p)))
                    else:
                        if dur < self.s.threshold_seconds:
                            st.short += 1
                            if self.s.action == ActionMode.ANALYZE:
                                self.row.emit(("SHORT", dur, p.name, str(p)))
                            else:
                                try:
                                    self._delete(p)
                                    st.deleted += 1
                                    self.row.emit(("DELETED", dur, p.name, str(p)))
                                except Exception:
                                    st.errors += 1
                                    self.row.emit(("ERROR", dur, p.name, str(p)))
                        else:
                            st.kept += 1
                            self.row.emit(("KEEP", dur, p.name, str(p)))

                    self.stats.emit(st)
                    self.progress.emit(i, total)

                self.finished.emit(st)

            finally:
                pythoncom.CoUninitialize()

        except Exception as e:
            self.failed.emit(str(e))


class ShortVideoCleanerPage(QWidget):
    def __init__(self, settings: QSettings):
        super().__init__()
        self.settings = settings

        self._thread: Optional[QThread] = None
        self._worker: Optional[CleanerWorker] = None

        self.model = QStandardItemModel(0, 3)
        self.model.setHorizontalHeaderLabels(["Status", "Duration (s)", "File"])

        self._build_ui()
        self._load()
        self.refresh_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)

        header = QFrame()
        hl = QHBoxLayout(header)
        title = QLabel("Short Video Cleaner")
        tf = QFont()
        tf.setPointSize(12)
        tf.setBold(True)
        title.setFont(tf)
        micro = QLabel("Scan → Review → Apply")
        micro.setStyleSheet("color: #cfcfcf;")
        hl.addWidget(title)
        hl.addStretch(1)
        hl.addWidget(micro)
        outer.addWidget(header)

        main_box = QGroupBox("Setup")
        main_l = QVBoxLayout(main_box)

        row_folder = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Select a folder...")
        self.path_edit.setToolTip("Folder to scan for videos.")
        self.btn_browse = QPushButton("Select")
        self.btn_open = QPushButton("Open")
        row_folder.addWidget(QLabel("Folder:"))
        row_folder.addWidget(self.path_edit, stretch=1)
        row_folder.addWidget(self.btn_browse)
        row_folder.addWidget(self.btn_open)
        main_l.addLayout(row_folder)

        row_rules = QHBoxLayout()
        self.threshold = QDoubleSpinBox()
        self.threshold.setRange(0.0, 86400.0)
        self.threshold.setDecimals(3)
        self.threshold.setValue(3.0)
        self.threshold.setSuffix(" s")
        self.threshold.setToolTip("Videos shorter than this duration are treated as SHORT.")

        self.action = QComboBox()
        self.action.addItem(ActionMode.ANALYZE.value, ActionMode.ANALYZE)
        self.action.addItem(ActionMode.RECYCLE.value, ActionMode.RECYCLE)
        self.action.addItem(ActionMode.HARD_DELETE.value, ActionMode.HARD_DELETE)
        self.action.setCurrentIndex(0)
        self.action.setToolTip("Choose what happens to SHORT videos. UNKNOWN duration is never deleted.")

        row_rules.addWidget(QLabel("Shorter than:"))
        row_rules.addWidget(self.threshold)
        row_rules.addSpacing(12)
        row_rules.addWidget(QLabel("Action:"))
        row_rules.addWidget(self.action, stretch=1)

        main_l.addLayout(row_rules)

        outer.addWidget(main_box)

        self.adv_toggle = QToolButton()
        self.adv_toggle.setText("Advanced")
        self.adv_toggle.setCheckable(True)
        self.adv_toggle.setChecked(False)
        self.adv_toggle.setToolTip("Show/hide advanced options (extensions and recursion).")
        outer.addWidget(self.adv_toggle, alignment=Qt.AlignmentFlag.AlignLeft)

        self.adv_box = QGroupBox("Advanced options")
        adv_l = QHBoxLayout(self.adv_box)

        self.ext_edit = QLineEdit(DEFAULT_EXTS)
        self.ext_edit.setPlaceholderText("Extensions, e.g. mp4,mov,mkv")
        self.ext_edit.setToolTip("Comma-separated extensions to include in the scan.")
        self.recursive_chk = QCheckBox("Recursive")
        self.recursive_chk.setChecked(True)
        self.recursive_chk.setToolTip("Include subfolders in the scan.")

        adv_l.addWidget(QLabel("Extensions:"))
        adv_l.addWidget(self.ext_edit, stretch=1)
        adv_l.addSpacing(12)
        adv_l.addWidget(self.recursive_chk)

        self.adv_box.setVisible(False)
        outer.addWidget(self.adv_box)

        controls = QHBoxLayout()
        self.btn_start = QPushButton("Start")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_clear = QPushButton("Clear")

        for b in (self.btn_start, self.btn_cancel, self.btn_clear):
            b.setMinimumHeight(40)
            b.setMinimumWidth(140)

        self.btn_cancel.setEnabled(False)

        controls.addWidget(self.btn_start)
        controls.addWidget(self.btn_cancel)
        controls.addWidget(self.btn_clear)
        controls.addStretch(1)
        outer.addLayout(controls)

        info = QHBoxLayout()
        self.stats_lbl = QLabel("Ready.")
        self.stats_lbl.setStyleSheet("font-weight: 800;")
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #cfcfcf;")
        info.addWidget(self.stats_lbl)
        info.addStretch(1)
        info.addWidget(self.status_lbl)
        outer.addLayout(info)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        outer.addWidget(self.progress)

        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setToolTip("Results table. Hover a row to see the full path.")
        outer.addWidget(self.table, stretch=1)

        # wiring
        self.btn_browse.clicked.connect(self.on_browse)
        self.btn_open.clicked.connect(self.on_open_folder)
        self.btn_start.clicked.connect(self.on_start)
        self.btn_cancel.clicked.connect(self.on_cancel)
        self.btn_clear.clicked.connect(self.on_clear)

        self.path_edit.textChanged.connect(self.refresh_ui)
        self.threshold.valueChanged.connect(self.refresh_ui)
        self.action.currentIndexChanged.connect(self.refresh_ui)
        self.ext_edit.textChanged.connect(self.refresh_ui)
        self.recursive_chk.stateChanged.connect(self.refresh_ui)

        self.adv_toggle.toggled.connect(self.on_toggle_advanced)

        # Drag & drop
        self.setAcceptDrops(True)

    def _load(self) -> None:
        self.path_edit.setText(self.settings.value("cleaner/dir", "", type=str))
        self.threshold.setValue(float(self.settings.value("cleaner/threshold", 3.0, type=float)))
        self.ext_edit.setText(self.settings.value("cleaner/exts", DEFAULT_EXTS, type=str))
        self.recursive_chk.setChecked(self.settings.value("cleaner/recursive", True, type=bool))
        action = self.settings.value("cleaner/action", ActionMode.ANALYZE.value, type=str)
        idx = self.action.findText(action)
        if idx >= 0:
            self.action.setCurrentIndex(idx)

    def _save(self) -> None:
        self.settings.setValue("cleaner/dir", self.path_edit.text().strip())
        self.settings.setValue("cleaner/threshold", float(self.threshold.value()))
        self.settings.setValue("cleaner/exts", self.ext_edit.text())
        self.settings.setValue("cleaner/recursive", self.recursive_chk.isChecked())
        self.settings.setValue("cleaner/action", str(self.action.currentText()))

    def on_toggle_advanced(self, checked: bool) -> None:
        self.adv_box.setVisible(checked)

    def is_running(self) -> bool:
        return self._thread is not None and self._worker is not None

    def selected_dir(self) -> Optional[Path]:
        txt = self.path_edit.text().strip()
        if not txt:
            return None
        p = Path(txt)
        if p.exists() and p.is_dir():
            return p
        return None

    def refresh_ui(self) -> None:
        valid_dir = self.selected_dir() is not None
        running = self.is_running()

        self.btn_open.setEnabled(valid_dir and not running)
        self.btn_browse.setEnabled(not running)
        self.btn_start.setEnabled(valid_dir and not running)
        self.btn_cancel.setEnabled(running)
        self.btn_clear.setEnabled(not running)

        # recycle availability
        recycle_idx = self.action.findData(ActionMode.RECYCLE)
        if recycle_idx >= 0:
            item = self.action.model().item(recycle_idx)  # type: ignore[union-attr]
            if item is not None:
                item.setEnabled(send2trash is not None)

        if pythoncom is None or win32com is None:
            self.status_lbl.setText("pywin32 missing")
        else:
            self.status_lbl.setText("")

        self._save()

    def on_browse(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "Select folder", self.path_edit.text() or str(Path.home()))
        if d:
            self.path_edit.setText(d)

    def on_open_folder(self) -> None:
        d = self.selected_dir()
        if d is None:
            QMessageBox.information(self, "Info", "Select a valid folder first.")
            return
        try:
            os.startfile(str(d))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open folder:\n{e!r}")

    def on_clear(self) -> None:
        self.model.removeRows(0, self.model.rowCount())
        self.progress.setValue(0)
        self.stats_lbl.setText("Ready.")
        self.status_lbl.setText("")

    def build_settings_or_error(self) -> Optional[CleanerSettings]:
        directory = self.selected_dir()
        if directory is None:
            QMessageBox.critical(self, "Error", "Select a valid folder.")
            return None

        exts = parse_extensions(self.ext_edit.text())
        if not exts:
            QMessageBox.critical(self, "Error", "Extensions are empty (e.g. mp4,mov,mkv).")
            return None

        if not sys.platform.startswith("win"):
            QMessageBox.critical(self, "Error", "Short Video Cleaner is Windows-only.")
            return None

        if pythoncom is None or win32com is None:
            QMessageBox.critical(self, "Error", "pywin32 is missing. Install: py -m pip install pywin32")
            return None

        action = self.action.currentData()
        if not isinstance(action, ActionMode):
            action = ActionMode.ANALYZE

        if action == ActionMode.RECYCLE and send2trash is None:
            QMessageBox.critical(self, "Error", "Recycle Bin mode is unavailable (send2trash missing).")
            return None

        if action == ActionMode.HARD_DELETE:
            r = QMessageBox.warning(
                self,
                "Permanent Delete",
                "Permanent deletion cannot be undone.\n\nContinue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return None

        return CleanerSettings(
            directory=directory,
            recursive=self.recursive_chk.isChecked(),
            threshold_seconds=float(self.threshold.value()),
            extensions=exts,
            action=action,
        )

    def lock_ui(self, locked: bool) -> None:
        self.path_edit.setEnabled(not locked)
        self.btn_browse.setEnabled(not locked)
        self.btn_open.setEnabled((self.selected_dir() is not None) and not locked)

        self.threshold.setEnabled(not locked)
        self.action.setEnabled(not locked)

        self.adv_toggle.setEnabled(not locked)
        self.ext_edit.setEnabled(not locked)
        self.recursive_chk.setEnabled(not locked)

        self.btn_start.setEnabled((self.selected_dir() is not None) and not locked)
        self.btn_cancel.setEnabled(locked)
        self.btn_clear.setEnabled(not locked)

    def on_start(self) -> None:
        s = self.build_settings_or_error()
        if s is None:
            return

        self.on_clear()
        self.lock_ui(True)
        self.status_lbl.setText("Scanning...")

        self._thread = QThread()
        self._worker = CleanerWorker(s)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.on_worker_progress)
        self._worker.status.connect(self.on_worker_status)
        self._worker.stats.connect(self.on_worker_stats)
        self._worker.row.connect(self.on_worker_row)
        self._worker.finished.connect(self.on_worker_finished)
        self._worker.failed.connect(self.on_worker_failed)

        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()

    def on_cancel(self) -> None:
        if self._worker:
            self._worker.cancel()

    def on_worker_progress(self, done: int, total: int) -> None:
        if total <= 0:
            self.progress.setValue(0)
            return
        self.progress.setValue(int((done / total) * 100))

    def on_worker_status(self, msg: str) -> None:
        self.status_lbl.setText(msg)

    def on_worker_stats(self, st_obj: object) -> None:
        if not isinstance(st_obj, CleanerStats):
            return
        self.stats_lbl.setText(
            f"Found={st_obj.found}  Scanned={st_obj.scanned}  Short={st_obj.short}  "
            f"Deleted={st_obj.deleted}  Unknown={st_obj.unknown}  Errors={st_obj.errors}"
        )

    def on_worker_row(self, row_obj: object) -> None:
        if not isinstance(row_obj, tuple) or len(row_obj) != 4:
            return

        status, dur, filename, fullpath = row_obj

        font_emphasis = QFont()
        font_emphasis.setBold(True)

        it_status = QStandardItem(str(status))
        it_dur = QStandardItem("" if dur is None else f"{float(dur):.3f}")
        it_file = QStandardItem(str(filename))

        tooltip = str(fullpath)
        for it in (it_status, it_dur, it_file):
            it.setToolTip(tooltip)

        s = str(status).upper()
        if s in {"DELETED", "ERROR"}:
            it_status.setFont(font_emphasis)
            it_file.setFont(font_emphasis)

        self.model.appendRow([it_status, it_dur, it_file])

    def on_worker_finished(self, st_obj: object) -> None:
        self.status_lbl.setText("Done.")
        self._worker = None
        self._thread = None
        self.lock_ui(False)
        self.refresh_ui()

    def on_worker_failed(self, msg: str) -> None:
        QMessageBox.critical(self, "Error", msg)
        self.status_lbl.setText("Error.")
        self._worker = None
        self._thread = None
        self.lock_ui(False)
        self.refresh_ui()

    # Drag & drop folder
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            p = Path(urls[0].toLocalFile())
            if p.is_dir():
                self.path_edit.setText(str(p))


# =========================
# Main Window (Navigation)
# =========================
class MediaFlowMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1380, 780)

        # Optional icon: assets/MediaFlow.png
        try:
            self.setWindowIcon(QIcon(resource_path("assets/MediaFlow.png")))
        except Exception:
            pass

        self.settings = QSettings(APP_ORG, APP_SETTINGS)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)

        self.nav = QListWidget()
        self.nav.setFixedWidth(220)
        self.nav.setToolTip("Choose a workflow.")
        layout.addWidget(self.nav)

        self.pages = QVBoxLayout()
        layout.addLayout(self.pages, 1)

        # header in content area
        content = QWidget()
        self.pages.addWidget(content, 1)
        content_layout = QVBoxLayout(content)

        top = QFrame()
        tl = QHBoxLayout(top)
        title = QLabel(APP_NAME)
        f = QFont()
        f.setPointSize(16)
        f.setBold(True)
        title.setFont(f)

        subtitle = QLabel("Media sorting & cleanup suite")
        subtitle.setStyleSheet("color: #cfcfcf;")

        tl.addWidget(title)
        tl.addSpacing(10)
        tl.addWidget(subtitle)
        tl.addStretch(1)
        content_layout.addWidget(top)

        # stacked-like with manual swap
        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.addWidget(self.container, 1)

        self.page_sort = SortMediaPage(self.settings)
        self.page_clean = ShortVideoCleanerPage(self.settings)

        self._current: Optional[QWidget] = None

        self._add_nav_item("Sort Media", "Analyze → Preview → Execute")
        self._add_nav_item("Short Video Cleaner", "Scan → Review → Apply")

        self.nav.currentRowChanged.connect(self._switch_page)
        self.nav.setCurrentRow(0)

        self.statusBar().showMessage("Ready.")

    def _add_nav_item(self, title: str, hint: str) -> None:
        it = QListWidgetItem(title)
        it.setToolTip(hint)
        self.nav.addItem(it)

    def _switch_page(self, idx: int) -> None:
        if self._current is not None:
            self._current.setParent(None)
            self._current = None

        if idx == 0:
            self._current = self.page_sort
        else:
            self._current = self.page_clean

        self.container_layout.addWidget(self._current)

    def closeEvent(self, event):
        # graceful cancel if needed
        try:
            if self.page_sort.thread is not None:
                self.page_sort.cancel_event.set()
        except Exception:
            pass
        try:
            if self.page_clean.is_running():
                self.page_clean.on_cancel()
        except Exception:
            pass
        super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    apply_business_dark(app)

    w = MediaFlowMainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
