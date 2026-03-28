import os
import re
import sys
import json
import html
from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple, Optional, Dict

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QTimer, QMimeData
from PyQt5.QtGui import QClipboard, QGuiApplication, QKeySequence
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QFileDialog,
    QMessageBox,
    QLabel,
    QPushButton,
    QTextEdit,
    QLineEdit,
    QHBoxLayout,
    QVBoxLayout,
    QFrame,
    QProgressBar,
    QCheckBox,
    QGridLayout,
    QToolButton,
    QShortcut,
    QDialog,
    QScrollArea,
)

# ------------------------------------------------------------
# Constants
# ------------------------------------------------------------
MARKER_FOLDERS = {
    "exr",
    "support_files",
    "paint",
    "comp",
    "roto",
    "plates",
    "renders",
    "delivery",
    "nuke",
    "_exr",
    "_support_files",
    "_paint",
    "_comp",
    "_roto",
    "_plates",
    "_renders",
    "_delivery",
    "_nuke",
}

SEQUENCE_EXTENSIONS_HINT = {
    ".exr",
    ".dpx",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
}

CURRENT_THEME = "dark"

THEMES = {
    "dark": {
        "root_bg": "#0d0f14",
        "panel_bg": "#151821",
        "toolbar_bg": "#11141a",
        "status_bg": "#11141a",
        "text_bg": "#0a0c10",
        "text_fg": "#f5f7fa",
        "file_fg": "#d7dce2",
        "folder_fg": "#ffffff",
        "missing_fg": "#ffb86b",
        "muted_fg": "#aeb6c2",
        "entry_bg": "#0b0d12",
        "entry_fg": "#f5f7fa",
        "border": "#2d3340",
        "button_bg": "#4f7cff",
        "button_fg": "#ffffff",
        "button_alt_bg": "#232833",
        "button_alt_fg": "#f5f7fa",
        "copy_rich_bg": "#2f8f5b",
        "copy_rich_fg": "#ffffff",
        "copy_rich_border": "#49b879",
        "copy_html_bg": "#b23b3b",
        "copy_html_fg": "#ffffff",
        "copy_html_border": "#df6c6c",
        "preview_count_bg": "#2f5fa7",
        "preview_count_fg": "#ffffff",
        "progress_bg": "#4f7cff",
        "section_title_fg": "#dfe7f7",
        "preview_header_bg": "#121722",
        "preview_header_border": "#2b3342",
        "primary_button_bg": "#4f7cff",
        "primary_button_hover": "#6790ff",
    },
    "light": {
        "root_bg": "#eef2f7",
        "panel_bg": "#ffffff",
        "toolbar_bg": "#f4f7fb",
        "status_bg": "#f4f7fb",
        "text_bg": "#ffffff",
        "text_fg": "#111827",
        "file_fg": "#374151",
        "folder_fg": "#111827",
        "missing_fg": "#b45309",
        "muted_fg": "#667085",
        "entry_bg": "#ffffff",
        "entry_fg": "#111827",
        "border": "#d2d9e3",
        "button_bg": "#2563eb",
        "button_fg": "#ffffff",
        "button_alt_bg": "#eef2f7",
        "button_alt_fg": "#111827",
        "copy_rich_bg": "#2f8f5b",
        "copy_rich_fg": "#ffffff",
        "copy_rich_border": "#4caf7c",
        "copy_html_bg": "#c94848",
        "copy_html_fg": "#ffffff",
        "copy_html_border": "#e17878",
        "preview_count_bg": "#dbeafe",
        "preview_count_fg": "#1d4ed8",
        "progress_bg": "#2563eb",
        "section_title_fg": "#334155",
        "preview_header_bg": "#f8fafc",
        "preview_header_border": "#d8dee8",
        "primary_button_bg": "#2563eb",
        "primary_button_hover": "#3b82f6",
    },
}


# ------------------------------------------------------------
# Data models
# ------------------------------------------------------------
@dataclass
class LineItem:
    item_type: str  # folder | file | missing | blank
    level: int
    text: str


@dataclass
class BuildResult:
    lines: List[LineItem]
    plain_text: str
    html_text: str
    clipboard_html: str
    folder_count: int
    file_count: int
    signature: Optional[Tuple]


# ------------------------------------------------------------
# Core logic
# ------------------------------------------------------------
class MailPrepLogic:
    @staticmethod
    def group_sequences(files: List[str]) -> List[str]:
        pattern = re.compile(r"^(.*?)(\d+)(\.[^.]+)$")
        grouped: Dict[Tuple[str, str], List[Tuple[int, int]]] = {}
        non_sequence = []

        for f in files:
            match = pattern.match(f)
            if match:
                base, num, ext = match.groups()
                key = (base, ext)
                grouped.setdefault(key, []).append((int(num), len(num)))
            else:
                non_sequence.append(f)

        result = []

        for (base, ext), items in grouped.items():
            num_to_width = {}
            for num, width in items:
                num_to_width[num] = max(width, num_to_width.get(num, 0))

            sorted_nums = sorted(num_to_width.keys())
            if not sorted_nums:
                continue

            max_width = max(num_to_width.values())
            start = sorted_nums[0]
            prev = sorted_nums[0]

            for current in sorted_nums[1:]:
                if current == prev + 1:
                    prev = current
                else:
                    if start == prev:
                        result.append(f"{base}{str(start).zfill(max_width)}{ext}")
                    else:
                        result.append(
                            f"{base}{str(start).zfill(max_width)}-{str(prev).zfill(max_width)}{ext}"
                        )
                    start = current
                    prev = current

            if start == prev:
                result.append(f"{base}{str(start).zfill(max_width)}{ext}")
            else:
                result.append(
                    f"{base}{str(start).zfill(max_width)}-{str(prev).zfill(max_width)}{ext}"
                )

        result.extend(sorted(non_sequence))
        return sorted(result)

    @staticmethod
    def _compress_number_ranges(numbers: List[int], width: int) -> str:
        if not numbers:
            return ""

        numbers = sorted(set(numbers))
        parts = []
        start = numbers[0]
        prev = numbers[0]

        for n in numbers[1:]:
            if n == prev + 1:
                prev = n
            else:
                if start == prev:
                    parts.append(str(start).zfill(width))
                else:
                    parts.append(f"{str(start).zfill(width)}-{str(prev).zfill(width)}")
                start = prev = n

        if start == prev:
            parts.append(str(start).zfill(width))
        else:
            parts.append(f"{str(start).zfill(width)}-{str(prev).zfill(width)}")

        return ", ".join(parts)

    @staticmethod
    def detect_missing_ranges(files: List[str]) -> List[str]:
        pattern = re.compile(r"^(.*?)(\d+)(\.[^.]+)$")
        grouped: Dict[Tuple[str, str], List[Tuple[int, int]]] = {}

        for f in files:
            match = pattern.match(f)
            if not match:
                continue

            base, num, ext = match.groups()
            if ext.lower() not in SEQUENCE_EXTENSIONS_HINT and len(num) < 3:
                continue

            key = (base, ext)
            grouped.setdefault(key, []).append((int(num), len(num)))

        missing_lines = []

        for (base, ext), items in grouped.items():
            unique = {}
            for num, width in items:
                unique[num] = max(width, unique.get(num, 0))

            nums = sorted(unique.keys())
            if len(nums) < 2:
                continue

            width = max(unique.values())
            missing = []

            for i in range(len(nums) - 1):
                current_num = nums[i]
                next_num = nums[i + 1]
                if next_num - current_num > 1:
                    missing.extend(range(current_num + 1, next_num))

            if not missing:
                continue

            compressed = MailPrepLogic._compress_number_ranges(missing, width)
            label = f"Missing frames: {compressed}"

            if len(grouped) > 1:
                label = f"Missing frames ({base}*{ext}): {compressed}"

            missing_lines.append(label)

        return sorted(missing_lines)

    @staticmethod
    def get_dir_entries(path: str) -> Tuple[List[str], List[str]]:
        try:
            entries = sorted(os.listdir(path))
        except Exception:
            return [], []

        dirs = []
        files = []

        for name in entries:
            full = os.path.join(path, name)
            if os.path.isdir(full):
                dirs.append(name)
            else:
                files.append(name)

        return dirs, files

    @staticmethod
    def is_single_shot_root(path: str) -> bool:
        try:
            entries = os.listdir(path)
        except Exception:
            return False

        subdirs = {
            name.lower() for name in entries if os.path.isdir(os.path.join(path, name))
        }
        return any(name in MARKER_FOLDERS for name in subdirs)

    @staticmethod
    def should_skip_root_for_single_marker(path: str) -> bool:
        dirs, files = MailPrepLogic.get_dir_entries(path)
        if files:
            return False
        if len(dirs) != 1:
            return False
        return dirs[0].lower() in MARKER_FOLDERS

    @staticmethod
    def walk_tree_files_first(
        current_path: str, current_level: int, lines: List[LineItem]
    ) -> None:
        dirs, files = MailPrepLogic.get_dir_entries(current_path)

        grouped_files = MailPrepLogic.group_sequences(files)
        for f in grouped_files:
            lines.append(LineItem("file", current_level, f))

        missing_lines = MailPrepLogic.detect_missing_ranges(files)
        for missing_text in missing_lines:
            lines.append(LineItem("missing", current_level, missing_text))

        for d in dirs:
            lines.append(LineItem("folder", current_level, d))
            MailPrepLogic.walk_tree_files_first(
                os.path.join(current_path, d),
                current_level + 1,
                lines,
            )

    @staticmethod
    def build_single_shot_lines(path: str, level: int = 0) -> List[LineItem]:
        lines: List[LineItem] = []
        root_name = os.path.basename(path.rstrip(os.sep))
        lines.append(LineItem("folder", level, root_name))
        MailPrepLogic.walk_tree_files_first(path, level + 1, lines)
        return lines

    @staticmethod
    def build_single_shot_contents_only(path: str, level: int = 0) -> List[LineItem]:
        lines: List[LineItem] = []
        MailPrepLogic.walk_tree_files_first(path, level, lines)
        return lines

    @staticmethod
    def build_shot_block(path: str, level: int = 0) -> List[LineItem]:
        lines: List[LineItem] = []
        shot_name = os.path.basename(path.rstrip(os.sep))
        lines.append(LineItem("folder", level, shot_name))
        MailPrepLogic.walk_tree_files_first(path, level + 1, lines)
        return lines

    @staticmethod
    def build_package_lines(path: str) -> List[LineItem]:
        if MailPrepLogic.should_skip_root_for_single_marker(path):
            return MailPrepLogic.build_single_shot_contents_only(path, level=0)

        if MailPrepLogic.is_single_shot_root(path):
            return MailPrepLogic.build_single_shot_lines(path)

        root_dirs, root_files = MailPrepLogic.get_dir_entries(path)

        if len(root_dirs) == 1 and not root_files:
            only_child = os.path.join(path, root_dirs[0])

            if MailPrepLogic.should_skip_root_for_single_marker(only_child):
                lines = [LineItem("folder", 0, root_dirs[0])]
                child_lines = MailPrepLogic.build_single_shot_contents_only(
                    only_child, level=1
                )
                lines.extend(child_lines)
                return lines

            if MailPrepLogic.is_single_shot_root(only_child):
                return MailPrepLogic.build_single_shot_lines(only_child)

            lines: List[LineItem] = []
            package_name = root_dirs[0]
            lines.append(LineItem("folder", 0, package_name))

            package_dirs, package_files = MailPrepLogic.get_dir_entries(only_child)

            grouped_files = MailPrepLogic.group_sequences(package_files)
            for f in grouped_files:
                lines.append(LineItem("file", 1, f))

            missing_lines = MailPrepLogic.detect_missing_ranges(package_files)
            for missing_text in missing_lines:
                lines.append(LineItem("missing", 1, missing_text))

            for idx, d in enumerate(package_dirs):
                shot_path = os.path.join(only_child, d)
                shot_block = MailPrepLogic.build_shot_block(shot_path, level=1)
                lines.extend(shot_block)
                if idx < len(package_dirs) - 1:
                    lines.append(LineItem("blank", 0, ""))

            return lines

        lines: List[LineItem] = []

        grouped_files = MailPrepLogic.group_sequences(root_files)
        for f in grouped_files:
            lines.append(LineItem("file", 0, f))

        missing_lines = MailPrepLogic.detect_missing_ranges(root_files)
        for missing_text in missing_lines:
            lines.append(LineItem("missing", 0, missing_text))

        for idx, d in enumerate(root_dirs):
            shot_path = os.path.join(path, d)
            shot_block = MailPrepLogic.build_shot_block(shot_path, level=0)
            lines.extend(shot_block)
            if idx < len(root_dirs) - 1:
                lines.append(LineItem("blank", 0, ""))

        return lines

    @staticmethod
    def make_plain_text_output(lines: List[LineItem]) -> str:
        output = []
        for item in lines:
            if item.item_type == "blank":
                output.append("")
            else:
                indent = "    " * item.level
                output.append(f"{indent}{item.text}")
        return "\n".join(output)

    @staticmethod
    def make_html_output(lines: List[LineItem], theme_name: str = "dark") -> str:
        theme = THEMES[theme_name]

        html_parts = [
            "<html><body>",
            (
                f'<div style="font-family:Calibri, Arial, sans-serif; '
                f'font-size:12pt; line-height:1.2; color:{theme["text_fg"]};">'
            ),
        ]

        for item in lines:
            if item.item_type == "blank":
                html_parts.append('<div style="height:0.45em;"></div>')
                continue

            safe_text = html.escape(item.text)
            indent = "&nbsp;" * (4 * item.level)

            if item.item_type == "folder":
                html_parts.append(
                    f'<div style="white-space:pre; color:{theme["folder_fg"]};">{indent}<b>{safe_text}</b></div>'
                )
            elif item.item_type == "missing":
                html_parts.append(
                    f'<div style="white-space:pre; color:{theme["missing_fg"]};">{indent}{safe_text}</div>'
                )
            else:
                html_parts.append(
                    f'<div style="white-space:pre; color:{theme["file_fg"]};">{indent}{safe_text}</div>'
                )

        html_parts.append("</div></body></html>")
        return "\n".join(html_parts)

    @staticmethod
    def _clipboard_line_html(text: str, level: int = 0, bold: bool = False) -> str:
        safe_text = html.escape(text)
        indent = "&nbsp;" * (4 * level)
        content = f"{indent}{safe_text}"

        if bold:
            content = f"<b>{content}</b>"

        return content

    @staticmethod
    def make_clipboard_rich_html(lines: List[LineItem]) -> str:
        parts = [
            "<html><body>",
            '<div style="font-family:Calibri, Arial, sans-serif; '
            'font-size:12pt; line-height:1.2; color:#000000; background-color:transparent;">',
        ]

        for item in lines:
            if item.item_type == "blank":
                parts.append("<br>")
                continue

            safe_text = html.escape(item.text)
            indent = "&nbsp;" * (4 * item.level)
            content = f"{indent}{safe_text}"

            if item.item_type == "folder":
                content = f"<b>{content}</b>"

            parts.append(f"{content}<br>")

        parts.append("</div></body></html>")
        return "".join(parts)

    @staticmethod
    def count_preview_items(lines: List[LineItem]) -> Tuple[int, int]:
        folder_count = sum(1 for item in lines if item.item_type == "folder")
        file_count = sum(1 for item in lines if item.item_type == "file")
        return folder_count, file_count

    @staticmethod
    def path_signature(path: str) -> Optional[Tuple]:
        try:
            dirs, files = MailPrepLogic.get_dir_entries(path)
            sig_parts = [("ROOT", int(os.path.getmtime(path)), 0)]

            for name in dirs + files:
                full = os.path.join(path, name)
                try:
                    stat = os.stat(full)
                    sig_parts.append((name, int(stat.st_mtime), stat.st_size))
                except Exception:
                    sig_parts.append((name, 0, 0))

            return tuple(sig_parts)
        except Exception:
            return None

    @staticmethod
    def build_result_from_lines(
        lines: List[LineItem],
        theme_name: str = "dark",
        signature: Optional[Tuple] = None,
    ) -> BuildResult:
        plain_text = MailPrepLogic.make_plain_text_output(lines)
        html_text = MailPrepLogic.make_html_output(lines, theme_name=theme_name)
        clipboard_html = MailPrepLogic.make_clipboard_rich_html(lines)
        folder_count, file_count = MailPrepLogic.count_preview_items(lines)

        return BuildResult(
            lines=lines,
            plain_text=plain_text,
            html_text=html_text,
            clipboard_html=clipboard_html,
            folder_count=folder_count,
            file_count=file_count,
            signature=signature,
        )

    @staticmethod
    def build_result_from_path(path: str, theme_name: str = "dark") -> BuildResult:
        lines = MailPrepLogic.build_package_lines(path)
        signature = MailPrepLogic.path_signature(path)
        return MailPrepLogic.build_result_from_lines(lines, theme_name, signature)

    @staticmethod
    def manifest_to_lines(manifest_data: dict) -> List[LineItem]:
        lines = []
        items = manifest_data.get("items", [])
        for item in items:
            try:
                lines.append(
                    LineItem(
                        item_type=str(item.get("type", "file")),
                        level=int(item.get("level", 0)),
                        text=str(item.get("text", "")),
                    )
                )
            except Exception:
                continue
        return lines

    @staticmethod
    def build_result_from_manifest(
        manifest_data: dict, theme_name: str = "dark"
    ) -> BuildResult:
        lines = MailPrepLogic.manifest_to_lines(manifest_data)
        signature = ("MANIFEST", len(lines), manifest_data.get("schema_version", "1.0"))
        return MailPrepLogic.build_result_from_lines(lines, theme_name, signature)

    @staticmethod
    def lines_to_manifest(lines: List[LineItem]) -> dict:
        return {
            "schema_version": "1.0",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "items": [
                {"type": item.item_type, "level": item.level, "text": item.text}
                for item in lines
            ],
        }

    @staticmethod
    def build_result(path: str, theme_name: str = "dark") -> BuildResult:
        return MailPrepLogic.build_result_from_path(path, theme_name)


# ------------------------------------------------------------
# Worker
# ------------------------------------------------------------
class PreviewWorker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, input_mode: str, source_data, theme_name: str):
        super().__init__()
        self.input_mode = input_mode
        self.source_data = source_data
        self.theme_name = theme_name

    def run(self):
        try:
            if self.input_mode == "path":
                result = MailPrepLogic.build_result_from_path(
                    self.source_data, self.theme_name
                )
            elif self.input_mode == "manifest":
                result = MailPrepLogic.build_result_from_manifest(
                    self.source_data, self.theme_name
                )
            else:
                raise ValueError(f"Unsupported input mode: {self.input_mode}")

            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


# ------------------------------------------------------------
# Preview Dialog
# ------------------------------------------------------------
class PreviewDialog(QDialog):
    def __init__(self, parent=None, html_content: str = "", theme_name: str = "dark"):
        super().__init__(parent)
        self.setWindowTitle("Preview")
        self.resize(900, 700)
        self.setMinimumSize(700, 500)

        theme = THEMES[theme_name]

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header = QLabel("Mail Preview")
        header.setObjectName("dialogHeader")

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setAcceptRichText(True)
        self.preview_box.setLineWrapMode(QTextEdit.NoWrap)
        self.preview_box.setHtml(html_content)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)

        footer_row = QHBoxLayout()
        footer_row.addStretch(1)
        footer_row.addWidget(close_btn)

        layout.addWidget(header)
        layout.addWidget(self.preview_box, 1)
        layout.addLayout(footer_row)

        self.setStyleSheet(
            f"""
            QDialog {{
                background: {theme["root_bg"]};
            }}
            QLabel#dialogHeader {{
                color: {theme["text_fg"]};
                font-size: 18px;
                font-weight: 700;
                padding: 4px 2px;
            }}
            QTextEdit {{
                background: {theme["text_bg"]};
                color: {theme["text_fg"]};
                border: 1px solid {theme["border"]};
                border-radius: 10px;
                padding: 10px;
            }}
            QPushButton {{
                background: {theme["button_alt_bg"]};
                color: {theme["button_alt_fg"]};
                border: 1px solid {theme["border"]};
                border-radius: 10px;
                padding: 8px 14px;
                font-weight: 600;
            }}
            """
        )


# ------------------------------------------------------------
# Main Window
# ------------------------------------------------------------
class MailPrepWindow(QMainWindow):
    COPY_FEEDBACK_MS = 1400

    def __init__(self):
        super().__init__()

        self.generated_html = ""
        self.generated_plain_text = ""
        self.generated_clipboard_html = ""
        self.generated_preview_html = ""
        self.generated_copy_html = ""
        self.generated_copy_plain_text = ""

        self.last_preview_signature = None
        self.loading_active = False
        self.current_job_id = 0
        self.current_theme = CURRENT_THEME
        self.last_refresh_display = "--"

        self.worker_thread: Optional[QThread] = None
        self.worker: Optional[PreviewWorker] = None

        self.shot_note_edits: Dict[str, QTextEdit] = {}
        self.current_shot_names: List[str] = []
        self.current_result_lines: List[LineItem] = []

        self.current_input_mode = "path"  # path | manifest
        self.current_manifest_data = None

        self.path_debounce_timer = QTimer(self)
        self.path_debounce_timer.setSingleShot(True)
        self.path_debounce_timer.timeout.connect(self._debounced_path_preview)

        self.auto_refresh_timer = QTimer(self)
        self.auto_refresh_timer.timeout.connect(self.auto_refresh_tick)
        self.auto_refresh_timer.start(3000)

        self.copy_feedback_timer = QTimer(self)
        self.copy_feedback_timer.setSingleShot(True)
        self.copy_feedback_timer.timeout.connect(self._restore_copy_button_labels)

        self._build_ui()
        self._setup_shortcuts()
        self.apply_theme()
        self._rebuild_shot_notes_panel([])

    def _build_ui(self):
        self.setWindowTitle("MailPrep Pro")
        self.resize(1380, 920)
        self.setMinimumSize(1120, 740)

        central = QWidget()
        self.setCentralWidget(central)

        self.outer_layout = QVBoxLayout(central)
        self.outer_layout.setContentsMargins(12, 12, 12, 12)
        self.outer_layout.setSpacing(10)

        self.header_frame = QFrame()
        self.header_frame.setObjectName("sectionCard")
        self.header_layout = QHBoxLayout(self.header_frame)
        self.header_layout.setContentsMargins(14, 12, 14, 12)
        self.header_layout.setSpacing(10)

        self.title_label = QLabel("📦 MailPrep Pro Tool")
        self.title_label.setObjectName("titleLabel")

        self.subtitle_label = QLabel(
            "Fast folder preview, Gmail-safe rich text copy, raw HTML copy"
        )
        self.subtitle_label.setObjectName("subtitleLabel")

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title_col.addWidget(self.title_label)
        title_col.addWidget(self.subtitle_label)

        self.theme_btn = QPushButton("☀ Light Mode")
        self.theme_btn.clicked.connect(self.toggle_theme)

        self.header_layout.addLayout(title_col, 1)
        self.header_layout.addWidget(self.theme_btn, 0, Qt.AlignRight | Qt.AlignVCenter)

        self.outer_layout.addWidget(self.header_frame)

        self.control_card = QFrame()
        self.control_card.setObjectName("sectionCard")
        self.control_layout = QVBoxLayout(self.control_card)
        self.control_layout.setContentsMargins(14, 14, 14, 14)
        self.control_layout.setSpacing(12)

        self.path_row = QHBoxLayout()
        self.path_row.setSpacing(8)

        self.path_label = QLabel("Folder")
        self.path_label.setObjectName("sectionLabel")

        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Select or paste folder path...")
        self.path_edit.textChanged.connect(self.on_path_change)

        self.browse_btn = QPushButton("📁 Browse")
        self.browse_btn.clicked.connect(self.select_folder)

        self.load_manifest_btn = QPushButton("📄 Manifest")
        self.load_manifest_btn.clicked.connect(self.load_manifest_file)

        self.export_manifest_btn = QPushButton("💾 Export")
        self.export_manifest_btn.clicked.connect(self.export_manifest_file)

        self.generate_btn = QPushButton("⚙ Generate Preview")
        self.generate_btn.setObjectName("primaryButton")
        self.generate_btn.clicked.connect(lambda: self.generate_preview(auto=False))
        self.generate_btn.setFixedHeight(34)

        self.auto_refresh_checkbox = QCheckBox("Auto Refresh")
        self.auto_refresh_checkbox.setChecked(True)
        self.auto_refresh_checkbox.toggled.connect(self.on_auto_refresh_toggle)

        self.path_row.addWidget(self.path_label)
        self.path_row.addWidget(self.path_edit, 1)
        self.path_row.addWidget(self.browse_btn)
        self.path_row.addWidget(self.load_manifest_btn)
        self.path_row.addWidget(self.export_manifest_btn)
        self.path_row.addWidget(self.generate_btn)
        self.path_row.addWidget(self.auto_refresh_checkbox)

        self.control_layout.addLayout(self.path_row)

        self.meta_grid = QGridLayout()
        self.meta_grid.setHorizontalSpacing(10)
        self.meta_grid.setVerticalSpacing(8)

        self.subject_label = QLabel("Subject")
        self.subject_label.setObjectName("sectionLabel")

        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("Optional mail subject...")
        self.subject_edit.textChanged.connect(self.on_metadata_change)

        self.shot_count_label = QLabel("Shot Count")
        self.shot_count_label.setObjectName("sectionLabel")

        self.shot_count_edit = QLineEdit()
        self.shot_count_edit.setPlaceholderText("Optional")
        self.shot_count_edit.setFixedWidth(120)
        self.shot_count_edit.textChanged.connect(self.on_metadata_change)

        self.preview_dialog_btn = QPushButton("🪟 Open Preview")
        self.preview_dialog_btn.clicked.connect(self.open_preview_dialog)

        self.clear_btn = QPushButton("🧹 Clear / Reset")
        self.clear_btn.clicked.connect(self.clear_form)

        self.meta_grid.addWidget(self.subject_label, 0, 0)
        self.meta_grid.addWidget(self.subject_edit, 0, 1)
        self.meta_grid.addWidget(self.shot_count_label, 0, 2)
        self.meta_grid.addWidget(self.shot_count_edit, 0, 3)
        self.meta_grid.addWidget(self.preview_dialog_btn, 0, 4)
        self.meta_grid.addWidget(self.clear_btn, 0, 5)

        self.control_layout.addLayout(self.meta_grid)

        self.action_row = QHBoxLayout()
        self.action_row.setSpacing(8)

        self.copy_rich_btn = QPushButton("📋 Copy Rich Text")
        self.copy_rich_btn.setObjectName("copyRichButton")
        self.copy_rich_btn.clicked.connect(self.copy_rich_text)

        self.copy_html_btn = QPushButton("</> Copy HTML")
        self.copy_html_btn.setObjectName("copyHtmlButton")
        self.copy_html_btn.clicked.connect(self.copy_html_source)

        self.advanced_toggle_btn = QToolButton()
        self.advanced_toggle_btn.setText("▶ Advanced Options")
        self.advanced_toggle_btn.setCheckable(True)
        self.advanced_toggle_btn.setChecked(False)
        self.advanced_toggle_btn.clicked.connect(self.toggle_advanced_options)

        self.action_row.addWidget(self.copy_rich_btn)
        self.action_row.addWidget(self.copy_html_btn)
        self.action_row.addStretch(1)
        self.action_row.addWidget(self.advanced_toggle_btn)

        self.control_layout.addLayout(self.action_row)

        self.advanced_frame = QFrame()
        self.advanced_frame.setObjectName("advancedFrame")
        self.advanced_layout = QHBoxLayout(self.advanced_frame)
        self.advanced_layout.setContentsMargins(10, 10, 10, 10)
        self.advanced_layout.setSpacing(18)

        self.include_subject_check = QCheckBox("Include Subject in Preview")
        self.include_subject_check.setChecked(True)
        self.include_subject_check.toggled.connect(self.on_metadata_change)

        self.include_shot_count_check = QCheckBox("Include Shot Count in Preview")
        self.include_shot_count_check.setChecked(True)
        self.include_shot_count_check.toggled.connect(self.on_metadata_change)

        self.include_submission_notes_check = QCheckBox("Include Submission Notes")
        self.include_submission_notes_check.setChecked(True)
        self.include_submission_notes_check.toggled.connect(self.on_metadata_change)

        self.include_meta_in_copy_check = QCheckBox("Include Metadata in Copy Output")
        self.include_meta_in_copy_check.setChecked(True)
        self.include_meta_in_copy_check.toggled.connect(self.on_metadata_change)

        self.complex_mode_check = QCheckBox("Complex Shot Mode")
        self.complex_mode_check.setChecked(False)
        self.complex_mode_check.toggled.connect(self.on_metadata_change)

        self.advanced_layout.addWidget(self.include_subject_check)
        self.advanced_layout.addWidget(self.include_shot_count_check)
        self.advanced_layout.addWidget(self.include_submission_notes_check)
        self.advanced_layout.addWidget(self.include_meta_in_copy_check)
        self.advanced_layout.addWidget(self.complex_mode_check)
        self.advanced_layout.addStretch(1)

        self.advanced_frame.hide()
        self.control_layout.addWidget(self.advanced_frame)

        self.outer_layout.addWidget(self.control_card)

        self.info_frame = QFrame()
        self.info_frame.setObjectName("sectionCard")
        self.info_layout = QHBoxLayout(self.info_frame)
        self.info_layout.setContentsMargins(14, 10, 14, 10)
        self.info_layout.setSpacing(8)

        self.preview_count_label = QLabel("Folders: 0   |   Files: 0")
        self.preview_count_label.setObjectName("previewCountLabel")

        self.last_refresh_label = QLabel("Last Refresh: --")
        self.last_refresh_label.setObjectName("hintLabel")

        self.source_label = QLabel("Source: Folder Path")
        self.source_label.setObjectName("hintLabel")

        self.loading_label = QLabel("Scanning folder...")
        self.loading_label.setObjectName("loadingLabel")
        self.loading_label.hide()

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFixedWidth(140)
        self.progress_bar.hide()

        self.hint_label = QLabel("Auto-refresh is ON")
        self.hint_label.setObjectName("hintLabel")

        self.info_layout.addWidget(self.preview_count_label, 0, Qt.AlignLeft)
        self.info_layout.addWidget(self.last_refresh_label, 0, Qt.AlignLeft)
        self.info_layout.addWidget(self.source_label, 0, Qt.AlignLeft)
        self.info_layout.addStretch(1)
        self.info_layout.addWidget(self.loading_label, 0, Qt.AlignRight)
        self.info_layout.addWidget(self.progress_bar, 0, Qt.AlignRight)
        self.info_layout.addWidget(self.hint_label, 0, Qt.AlignRight)

        self.outer_layout.addWidget(self.info_frame)

        self.main_content_row = QHBoxLayout()
        self.main_content_row.setSpacing(10)

        self.preview_card = QFrame()
        self.preview_card.setObjectName("previewCard")
        self.preview_layout = QVBoxLayout(self.preview_card)
        self.preview_layout.setContentsMargins(12, 12, 12, 12)
        self.preview_layout.setSpacing(8)

        self.preview_header = QFrame()
        self.preview_header.setObjectName("previewHeader")
        self.preview_header_layout = QHBoxLayout(self.preview_header)
        self.preview_header_layout.setContentsMargins(10, 8, 10, 8)

        self.preview_title = QLabel("Preview")
        self.preview_title.setObjectName("previewTitle")

        self.preview_subtitle = QLabel("One-shot rendered output")
        self.preview_subtitle.setObjectName("hintLabel")

        preview_head_col = QVBoxLayout()
        preview_head_col.setSpacing(0)
        preview_head_col.addWidget(self.preview_title)
        preview_head_col.addWidget(self.preview_subtitle)

        self.preview_header_layout.addLayout(preview_head_col, 1)

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setAcceptRichText(True)
        self.preview_box.setLineWrapMode(QTextEdit.NoWrap)

        self.preview_layout.addWidget(self.preview_header)
        self.preview_layout.addWidget(self.preview_box, 1)

        self.notes_card = QFrame()
        self.notes_card.setObjectName("sectionCard")
        self.notes_layout = QVBoxLayout(self.notes_card)
        self.notes_layout.setContentsMargins(12, 12, 12, 12)
        self.notes_layout.setSpacing(8)

        self.notes_title = QLabel("Shot Submission Notes")
        self.notes_title.setObjectName("previewTitle")

        self.notes_hint = QLabel("Editable notes for each derived version name")
        self.notes_hint.setObjectName("hintLabel")

        notes_head = QVBoxLayout()
        notes_head.setSpacing(0)
        notes_head.addWidget(self.notes_title)
        notes_head.addWidget(self.notes_hint)

        self.apply_all_row = QHBoxLayout()
        self.apply_all_row.setSpacing(8)

        self.common_note_edit = QLineEdit()
        self.common_note_edit.setPlaceholderText(
            "Common text to apply for all shots..."
        )
        self.common_note_edit.returnPressed.connect(self.apply_common_text_to_all_shots)

        self.apply_all_btn = QPushButton("⇢ Apply to All Shots")
        self.apply_all_btn.clicked.connect(self.apply_common_text_to_all_shots)

        self.apply_all_row.addWidget(self.common_note_edit, 1)
        self.apply_all_row.addWidget(self.apply_all_btn)

        self.notes_scroll = QScrollArea()
        self.notes_scroll.setWidgetResizable(True)
        self.notes_scroll.setFrameShape(QFrame.NoFrame)

        self.notes_scroll_widget = QWidget()
        self.notes_scroll_layout = QVBoxLayout(self.notes_scroll_widget)
        self.notes_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.notes_scroll_layout.setSpacing(10)

        self.notes_scroll.setWidget(self.notes_scroll_widget)

        self.notes_layout.addLayout(notes_head)
        self.notes_layout.addLayout(self.apply_all_row)
        self.notes_layout.addWidget(self.notes_scroll, 1)

        self.main_content_row.addWidget(self.preview_card, 3)
        self.main_content_row.addWidget(self.notes_card, 2)

        self.outer_layout.addLayout(self.main_content_row, 1)

        self.status_frame = QFrame()
        self.status_frame.setObjectName("sectionCard")
        self.status_layout = QHBoxLayout(self.status_frame)
        self.status_layout.setContentsMargins(14, 10, 14, 10)

        self.status_label = QLabel("Ready.")
        self.status_label.setObjectName("statusLabel")

        self.status_layout.addWidget(self.status_label)
        self.outer_layout.addWidget(self.status_frame)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+B"), self, activated=self.select_folder)
        QShortcut(
            QKeySequence("Ctrl+R"),
            self,
            activated=lambda: self.generate_preview(auto=False),
        )
        QShortcut(QKeySequence("Ctrl+Shift+R"), self, activated=self.copy_rich_text)
        QShortcut(QKeySequence("Ctrl+Shift+H"), self, activated=self.copy_html_source)
        QShortcut(QKeySequence("Ctrl+K"), self, activated=self.clear_form)
        QShortcut(QKeySequence("Ctrl+T"), self, activated=self.toggle_theme)
        QShortcut(QKeySequence("Ctrl+P"), self, activated=self.open_preview_dialog)
        QShortcut(QKeySequence("Ctrl+M"), self, activated=self.load_manifest_file)

    def apply_theme(self):
        theme = THEMES[self.current_theme]

        self.setStyleSheet(
            f"""
            QMainWindow {{
                background: {theme["root_bg"]};
            }}
            QFrame {{
                background: {theme["panel_bg"]};
                border: none;
            }}
            QFrame#sectionCard {{
                background: {theme["panel_bg"]};
                border: 1px solid {theme["border"]};
                border-radius: 14px;
            }}
            QFrame#advancedFrame {{
                background: {theme["toolbar_bg"]};
                border: 1px solid {theme["border"]};
                border-radius: 10px;
            }}
            #titleLabel {{
                color: {theme["text_fg"]};
                font-size: 20px;
                font-weight: 800;
                background: transparent;
            }}
            #subtitleLabel {{
                color: {theme["muted_fg"]};
                font-size: 12px;
                background: transparent;
            }}
            #sectionLabel {{
                color: {theme["section_title_fg"]};
                font-size: 12px;
                font-weight: 700;
                background: transparent;
                padding-bottom: 2px;
            }}
            QLineEdit, QTextEdit {{
                background: {theme["entry_bg"]};
                color: {theme["entry_fg"]};
                border: 1px solid {theme["border"]};
                border-radius: 12px;
                padding: 8px 10px;
                selection-background-color: {theme["button_bg"]};
            }}
            QPushButton {{
                background: {theme["button_alt_bg"]};
                color: {theme["button_alt_fg"]};
                border: 1px solid {theme["border"]};
                border-radius: 12px;
                padding: 9px 13px;
                font-weight: 700;
            }}
            QPushButton:hover {{
                border: 1px solid {theme["button_bg"]};
            }}
            QPushButton#primaryButton {{
                background: {theme["primary_button_bg"]};
                color: {theme["button_fg"]};
                border: none;
                border-radius: 16px;
                padding: 6px 16px;
                font-size: 12px;
                font-weight: 800;
                min-width: 130px;
            }}
            QPushButton#primaryButton:hover {{
                background: {theme["primary_button_hover"]};
            }}
            QPushButton#copyRichButton {{
                background: {theme["copy_rich_bg"]};
                color: {theme["copy_rich_fg"]};
                border: 1px solid {theme["copy_rich_border"]};
                border-radius: 12px;
                font-weight: 800;
            }}
            QPushButton#copyHtmlButton {{
                background: {theme["copy_html_bg"]};
                color: {theme["copy_html_fg"]};
                border: 1px solid {theme["copy_html_border"]};
                border-radius: 12px;
                font-weight: 800;
            }}
            QToolButton {{
                background: transparent;
                color: {theme["text_fg"]};
                border: none;
                font-weight: 700;
                padding: 6px 8px;
            }}
            QLabel {{
                color: {theme["text_fg"]};
                background: transparent;
            }}
            QCheckBox {{
                color: {theme["text_fg"]};
                spacing: 6px;
                background: transparent;
            }}
            #hintLabel, #loadingLabel, #statusLabel {{
                color: {theme["muted_fg"]};
            }}
            #previewCountLabel {{
                background: {theme["preview_count_bg"]};
                color: {theme["preview_count_fg"]};
                border-radius: 10px;
                padding: 6px 10px;
                font-weight: 800;
            }}
            #previewCard {{
                background: {theme["panel_bg"]};
                border-radius: 14px;
            }}
            #previewHeader {{
                background: {theme["preview_header_bg"]};
                border: 1px solid {theme["preview_header_border"]};
                border-radius: 12px;
            }}
            #previewTitle {{
                color: {theme["text_fg"]};
                font-size: 15px;
                font-weight: 800;
            }}
            QProgressBar {{
                border: 1px solid {theme["border"]};
                border-radius: 8px;
                background: {theme["panel_bg"]};
                text-align: center;
            }}
            QProgressBar::chunk {{
                background: {theme["progress_bg"]};
                border-radius: 8px;
            }}
            """
        )

        self.theme_btn.setText(
            "☀ Light Mode" if self.current_theme == "dark" else "🌙 Dark Mode"
        )

        if self.current_input_mode == "path":
            current_path = self.path_edit.text().strip()
            if (
                current_path
                and os.path.isdir(current_path)
                and self.generated_plain_text
            ):
                try:
                    result = MailPrepLogic.build_result_from_path(
                        current_path, self.current_theme
                    )
                    self.render_preview(result)
                except Exception:
                    pass
        elif self.current_input_mode == "manifest" and self.current_manifest_data:
            try:
                result = MailPrepLogic.build_result_from_manifest(
                    self.current_manifest_data, self.current_theme
                )
                self.render_preview(result)
            except Exception:
                pass

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme()
        self.set_status(f"Theme changed to {self.current_theme} mode.")

    def set_status(self, message: str):
        self.status_label.setText(message)

    def set_loading(self, is_loading: bool, message: str = "Loading..."):
        self.loading_active = is_loading
        self.loading_label.setText(message)
        self.loading_label.setVisible(is_loading)
        self.progress_bar.setVisible(is_loading)
        self.generate_btn.setDisabled(is_loading)
        self.browse_btn.setDisabled(is_loading)
        self.load_manifest_btn.setDisabled(is_loading)
        self.export_manifest_btn.setDisabled(is_loading)

    def _update_last_refresh(self):
        self.last_refresh_display = datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")
        self.last_refresh_label.setText(f"Last Refresh: {self.last_refresh_display}")

    def _restore_copy_button_labels(self):
        self.copy_rich_btn.setText("📋 Copy Rich Text")
        self.copy_html_btn.setText("</> Copy HTML")

    def _escape_with_breaks(self, value: str) -> str:
        return html.escape(value).replace("\n", "<br>")

    def _copy_safe_div(
        self, text: str, bold: bool = False, underline: bool = False
    ) -> str:
        safe_text = html.escape(text).replace("\n", "<br>")
        if underline:
            safe_text = f"<u>{safe_text}</u>"
        if bold:
            safe_text = f"<b>{safe_text}</b>"
        return safe_text

    def _looks_like_file_name(self, name: str) -> bool:
        return bool(re.search(r"\.[A-Za-z0-9]{2,8}$", name.strip()))

    def _is_ignored_submission_parent(self, name: str) -> bool:
        n = name.strip().lower()

        ignored_exact = {
            "h264",
            "h.264",
            "h265",
            "hevc",
            "mp4",
            "dnxhd",
            "dnxhr",
            "dnx",
            "prores",
            "pro_res",
            "proreshq",
            "avid",
            "mov",
            "mxf",
            "exr",
            "_exr",
            "jpeg",
            "jpg",
            "png",
            "tif",
            "tiff",
            "tif16",
            "tiff16",
            "delivery",
            "deliveries",
            "support_files",
            "_support_files",
            "support",
            "lut",
            "luts",
            "cdl",
            "ccc",
            "plates",
            "plate",
            "renders",
            "render",
            "output",
            "outputs",
            "preview",
            "previews",
            "qt",
            "quicktime",
            "review",
            "client",
            "publish",
            "published",
            "final",
            "temp",
            "export",
            "exports",
            "web",
            "maya",
            "nuke",
            "script",
            "scripts",
            "splines",
            "rotomation",
            "dailies",
            "fbx",
            "undistorted_plate",
            "holdout",
            "perspective",
            "perspective2",
            "shaded",
            "wireframe",
            "curves",
            "perspectivestab",
            "pointblast",
            "pointblastdigorychest",
        }

        if n in ignored_exact:
            return True

        fuzzy_tokens = (
            "h264",
            "h265",
            "hevc",
            "dnx",
            "prores",
            "quicktime",
            "qt",
            "review",
            "delivery",
            "deliver",
            "output",
            "export",
            "publish",
            "mov",
            "mp4",
            "mxf",
            "avid",
            "web",
            "support",
            "lut",
            "ccc",
            "cdl",
        )

        return any(token in n for token in fuzzy_tokens)

    def _cleanup_version_candidate(self, name: str) -> str:
        if not name:
            return ""

        candidate = name.strip()
        candidate = re.sub(r"\.[A-Za-z0-9]{2,8}$", "", candidate)
        candidate = re.sub(r"\.\d+-\d+$", "", candidate)
        candidate = re.sub(r"\.\d+$", "", candidate)

        removable_suffixes = [
            "h264",
            "h265",
            "hevc",
            "dnxhd",
            "dnxhr",
            "dnx",
            "prores",
            "pro_res",
            "proreshq",
            "qt",
            "quicktime",
            "mov",
            "mxf",
            "mp4",
            "avid",
            "review",
            "web",
            "final",
            "output",
            "export",
            "vfx",
            "rotoslap",
            "slap",
            "script",
            "sfx",
            "nk",
            "mb",
            "ma",
            "fbx",
            "jpg",
            "jpeg",
            "png",
            "tif",
            "tiff",
            "exr",
            "ccc",
            "cdl",
        ]

        changed = True
        while changed:
            changed = False
            for suffix in removable_suffixes:
                pattern = re.compile(rf"([_.-]){re.escape(suffix)}$", re.IGNORECASE)
                if pattern.search(candidate):
                    candidate = pattern.sub("", candidate)
                    changed = True

        version_matches = list(
            re.finditer(r"(?:^|[_.-])(v\d{1,4})(?=$|[_.-])", candidate, re.IGNORECASE)
        )
        if version_matches:
            last_match = version_matches[-1]
            candidate = candidate[: last_match.end(1)]

        candidate = candidate.strip("._- ")
        return candidate

    def _derive_submission_shot_name_from_filename(
        self, filename: str
    ) -> Optional[str]:
        if not filename:
            return None

        candidate = self._cleanup_version_candidate(filename)
        if not candidate:
            return None

        if self._is_ignored_submission_parent(candidate):
            return None

        return candidate

    def _looks_like_version_name(self, name: str) -> bool:
        if not name:
            return False

        cleaned = self._cleanup_version_candidate(name)
        if not cleaned:
            return False

        if self._is_ignored_submission_parent(cleaned):
            return False

        return bool(
            re.search(r"(?:^|[_.-])v\d{1,4}(?=$|[_.-])", cleaned, re.IGNORECASE)
        )

    def _extract_shot_names_default(self, lines: List[LineItem]) -> List[str]:
        results = []
        seen = set()

        for item in lines:
            if item.item_type != "folder":
                continue

            raw = item.text.strip()
            if not raw:
                continue

            if self._is_ignored_submission_parent(raw):
                continue

            if not self._looks_like_version_name(raw):
                continue

            version_name = self._cleanup_version_candidate(raw)
            if not version_name:
                continue

            if version_name not in seen:
                seen.add(version_name)
                results.append(version_name)

        if results:
            return results

        for item in lines:
            if item.item_type != "file":
                continue

            raw = item.text.strip()
            if not raw:
                continue

            if not self._looks_like_file_name(raw):
                continue

            version_name = self._derive_submission_shot_name_from_filename(raw)
            if not version_name:
                continue

            if not self._looks_like_version_name(version_name):
                continue

            if version_name not in seen:
                seen.add(version_name)
                results.append(version_name)

        return results

    def _candidate_priority_from_filename(self, filename: str) -> int:
        low = filename.lower()

        if low.endswith(".sfx"):
            return 0
        if low.endswith(".mb"):
            return 1
        if low.endswith(".ma"):
            return 2
        if low.endswith(".fbx"):
            return 3
        if low.endswith(".nk"):
            return 4

        return 99

    def _is_preview_like_candidate(self, candidate: str) -> bool:
        low = candidate.lower()
        preview_tokens = (
            "flattened",
            "grey",
            "gray",
            "wireframe",
            "rotolines",
            "overlay",
            "holdout",
            "perspective",
            "perspectivestab",
            "pointblast",
            "pointblastdigorychest",
            "shaded",
            "curves",
            "check",
            "rotocheck",
            "trkholdout",
            "trkperspective",
            "trkperspective2",
            "trkshaded",
            "trkwire",
            "trkcurves",
            "trkpointblast",
        )
        return any(token in low for token in preview_tokens)

    def _candidate_token_count(self, candidate: str) -> int:
        parts = [p for p in re.split(r"[._-]+", candidate) if p]
        return len(parts)

    def _extract_shot_names_complex(self, lines: List[LineItem]) -> List[str]:
        results = []
        seen = set()

        for item in lines:
            if item.item_type != "folder":
                continue
            if item.level != 0:
                continue

            raw = item.text.strip()
            if not raw:
                continue

            if self._is_ignored_submission_parent(raw):
                continue

            if not self._looks_like_version_name(raw):
                continue

            version_name = self._cleanup_version_candidate(raw)
            if not version_name:
                continue

            if version_name not in seen:
                seen.add(version_name)
                results.append(version_name)

        if results:
            return results

        candidates_by_root: Dict[str, Dict[str, Tuple[int, int, int, int]]] = {}
        current_root = None
        file_index = 0

        for item in lines:
            if item.item_type == "folder" and item.level == 0:
                current_root = item.text.strip()
                if current_root and current_root not in candidates_by_root:
                    candidates_by_root[current_root] = {}
                continue

            if item.item_type != "file":
                continue

            raw = item.text.strip()
            if not raw:
                continue

            if not self._looks_like_file_name(raw):
                continue

            candidate = self._derive_submission_shot_name_from_filename(raw)
            if not candidate:
                continue

            if not self._looks_like_version_name(candidate):
                continue

            if self._is_preview_like_candidate(candidate):
                continue

            root_key = current_root or "__ungrouped__"
            if root_key not in candidates_by_root:
                candidates_by_root[root_key] = {}

            priority = self._candidate_priority_from_filename(raw)
            token_count = self._candidate_token_count(candidate)
            length = len(candidate)

            root_map = candidates_by_root[root_key]

            if candidate not in root_map:
                root_map[candidate] = (priority, token_count, length, file_index)
            else:
                old_priority, old_token_count, old_length, old_index = root_map[
                    candidate
                ]
                root_map[candidate] = (
                    min(priority, old_priority),
                    min(token_count, old_token_count),
                    min(length, old_length),
                    old_index,
                )

            file_index += 1

        ordered_roots = list(candidates_by_root.keys())

        for root_key in ordered_roots:
            root_map = candidates_by_root.get(root_key, {})
            if not root_map:
                continue

            best_candidate = sorted(
                root_map.items(),
                key=lambda kv: (
                    kv[1][0],
                    kv[1][1],
                    kv[1][2],
                    kv[1][3],
                    kv[0].lower(),
                ),
            )[0][0]

            if best_candidate not in seen:
                seen.add(best_candidate)
                results.append(best_candidate)

        if results:
            return results

        return self._extract_shot_names_default(lines)

    def _extract_shot_names_from_lines(self, lines: List[LineItem]) -> List[str]:
        if self.complex_mode_check.isChecked():
            return self._extract_shot_names_complex(lines)
        return self._extract_shot_names_default(lines)

    def get_submission_notes_map(self) -> Dict[str, str]:
        data = {}
        for shot_name in self.current_shot_names:
            editor = self.shot_note_edits.get(shot_name)
            if not editor:
                continue
            data[shot_name] = editor.toPlainText()
        return data

    def _rebuild_shot_notes_panel(self, shot_names: List[str]):
        old_values = self.get_submission_notes_map()

        while self.notes_scroll_layout.count():
            item = self.notes_scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.shot_note_edits = {}
        self.current_shot_names = shot_names[:]

        if not shot_names:
            empty_label = QLabel("Generate preview to load shot-wise notes fields.")
            empty_label.setObjectName("hintLabel")
            self.notes_scroll_layout.addWidget(empty_label)
            self.notes_scroll_layout.addStretch(1)
            return

        for shot_name in shot_names:
            block = QFrame()
            block.setObjectName("sectionCard")
            block_layout = QVBoxLayout(block)
            block_layout.setContentsMargins(10, 10, 10, 10)
            block_layout.setSpacing(6)

            label = QLabel(shot_name)
            label.setStyleSheet("font-weight: 800;")

            editor = QTextEdit()
            editor.setAcceptRichText(False)
            editor.setPlaceholderText("Paste shot-specific submission note here...")
            editor.setFixedHeight(78)
            editor.textChanged.connect(self.on_metadata_change)

            if shot_name in old_values:
                editor.setPlainText(old_values[shot_name])

            block_layout.addWidget(label)
            block_layout.addWidget(editor)

            self.shot_note_edits[shot_name] = editor
            self.notes_scroll_layout.addWidget(block)

        self.notes_scroll_layout.addStretch(1)

    def apply_common_text_to_all_shots(self):
        common_text = self.common_note_edit.text().strip()
        if not common_text:
            self.set_status("Enter common text first.")
            return

        updated = 0
        for shot_name in self.current_shot_names:
            editor = self.shot_note_edits.get(shot_name)
            if editor is None:
                continue
            editor.setPlainText(common_text)
            updated += 1

        if updated:
            self.on_metadata_change()
            self.set_status(f"Applied common text to {updated} shot(s).")

    def _meta_rows(self) -> List[Tuple[str, str]]:
        rows = []

        subject = self.subject_edit.text().strip()
        shot_count = self.shot_count_edit.text().strip()

        if self.include_subject_check.isChecked() and subject:
            rows.append(("Subject", subject))

        if self.include_shot_count_check.isChecked() and shot_count:
            rows.append(("Shot Count", shot_count))

        return rows

    def _copy_meta_rows(self) -> List[Tuple[str, str]]:
        if not self.include_meta_in_copy_check.isChecked():
            return []
        return self._meta_rows()

    def _submission_notes_items(self) -> List[Tuple[str, str]]:
        if not self.include_submission_notes_check.isChecked():
            return []

        items = []
        for shot_name in self.current_shot_names:
            editor = self.shot_note_edits.get(shot_name)
            if not editor:
                continue
            note = editor.toPlainText().strip()
            if note:
                items.append((shot_name, note))
        return items

    def _compose_preview_html(self, tree_html: str) -> str:
        rows = self._meta_rows()
        notes_items = self._submission_notes_items()

        parts = [
            "<html><body>",
            (
                f'<div style="font-family:Calibri, Arial, sans-serif; '
                f'font-size:12pt; line-height:1.2; color:{THEMES[self.current_theme]["text_fg"]};">'
            ),
        ]

        if rows:
            theme = THEMES[self.current_theme]
            parts.append(
                f'<div style="margin-bottom:10px; padding:10px 12px; '
                f'border:1px solid {theme["preview_header_border"]}; '
                f'background:{theme["preview_header_bg"]}; border-radius:8px;">'
            )
            for label, value in rows:
                parts.append(
                    f'<div style="margin-bottom:4px;"><b>{html.escape(label)}:</b> {self._escape_with_breaks(value)}</div>'
                )
            parts.append("</div>")

        body_only = re.sub(r"^\s*<html><body>", "", tree_html, flags=re.IGNORECASE)
        body_only = re.sub(r"</body></html>\s*$", "", body_only, flags=re.IGNORECASE)
        parts.append(body_only)

        if notes_items:
            parts.append('<div style="height:1em;"></div>')
            parts.append("<div><b><u>Submission Notes :</u></b></div>")

            for idx, (shot_name, note) in enumerate(notes_items):
                parts.append(f"<div><b>{html.escape(shot_name)}</b></div>")
                for line in note.splitlines():
                    if line.strip():
                        parts.append(f"<div>{html.escape(line)}</div>")
                    else:
                        parts.append("<div><br></div>")

                if idx < len(notes_items) - 1:
                    parts.append("<div><br></div>")

        parts.append("</div></body></html>")
        return "".join(parts)

    def _compose_copy_rich_html(self, tree_clipboard_html: str) -> str:
        rows = self._copy_meta_rows()
        notes_items = self._submission_notes_items()

        parts = [
            "<html><body>",
            '<div style="font-family:Calibri, Arial, sans-serif; '
            'font-size:12pt; line-height:1.2; color:#000000; background-color:transparent;">',
        ]

        if rows:
            for label, value in rows:
                safe_value = html.escape(value).replace("\n", "<br>")
                parts.append(f"<b>{html.escape(label)}:</b> {safe_value}<br>")
            parts.append("<br>")

        body_only = re.sub(
            r"^\s*<html><body>", "", tree_clipboard_html, flags=re.IGNORECASE
        )
        body_only = re.sub(r"</body></html>\s*$", "", body_only, flags=re.IGNORECASE)

        body_only = re.sub(r"^\s*<div[^>]*>", "", body_only, flags=re.IGNORECASE)
        body_only = re.sub(r"</div>\s*$", "", body_only, flags=re.IGNORECASE)

        parts.append(body_only)

        if notes_items:
            parts.append("<br>")
            parts.append(
                self._copy_safe_div("Submission Notes :", bold=True, underline=True)
            )
            parts.append("<br>")

            for idx, (shot_name, note) in enumerate(notes_items):
                parts.append(self._copy_safe_div(shot_name, bold=True))
                parts.append("<br>")

                note_lines = note.splitlines()
                if note_lines:
                    for line in note_lines:
                        if line.strip():
                            parts.append(self._copy_safe_div(line))
                        parts.append("<br>")
                else:
                    parts.append("<br>")

                if idx < len(notes_items) - 1:
                    parts.append("<br>")

        parts.append("</div></body></html>")
        return "".join(parts)

    def _compose_copy_plain_text(self, tree_plain_text: str) -> str:
        rows = self._copy_meta_rows()
        notes_items = self._submission_notes_items()

        blocks = []

        if rows:
            for label, value in rows:
                blocks.append(f"{label}: {value}")

        blocks.append(tree_plain_text)

        if notes_items:
            notes_block = ["Submission Notes :", ""]
            for idx, (shot_name, note) in enumerate(notes_items):
                notes_block.append(shot_name)
                notes_block.append(note)
                if idx < len(notes_items) - 1:
                    notes_block.append("")
            blocks.append("\n".join(notes_block))

        return "\n\n".join([b for b in blocks if b])

    def clear_preview(self):
        self.generated_html = ""
        self.generated_plain_text = ""
        self.generated_clipboard_html = ""
        self.generated_preview_html = ""
        self.generated_copy_html = ""
        self.generated_copy_plain_text = ""
        self.last_preview_signature = None
        self.current_result_lines = []
        self.preview_box.clear()
        self.preview_count_label.setText("Folders: 0   |   Files: 0")
        self.last_refresh_label.setText("Last Refresh: --")
        self._rebuild_shot_notes_panel([])
        self.source_label.setText(
            "Source: Folder Path"
            if self.current_input_mode == "path"
            else "Source: Manifest"
        )

    def clear_form(self):
        self.path_edit.clear()
        self.subject_edit.clear()
        self.shot_count_edit.clear()
        self.common_note_edit.clear()
        self.current_input_mode = "path"
        self.current_manifest_data = None
        self.auto_refresh_checkbox.setEnabled(True)
        self.source_label.setText("Source: Folder Path")
        self.clear_preview()
        self.set_status("Form cleared.")

    def render_preview(self, result: BuildResult):
        existing_notes = self.get_submission_notes_map()

        self.generated_plain_text = result.plain_text
        self.generated_html = result.html_text
        self.generated_clipboard_html = result.clipboard_html
        self.last_preview_signature = result.signature
        self.current_result_lines = result.lines[:]

        shot_names = self._extract_shot_names_from_lines(result.lines)
        self._rebuild_shot_notes_panel(shot_names)

        for shot_name, text in existing_notes.items():
            editor = self.shot_note_edits.get(shot_name)
            if editor is not None and text:
                editor.setPlainText(text)

        self.generated_preview_html = self._compose_preview_html(result.html_text)
        self.generated_copy_html = self._compose_copy_rich_html(result.clipboard_html)
        self.generated_copy_plain_text = self._compose_copy_plain_text(
            result.plain_text
        )

        self.preview_box.setHtml(self.generated_preview_html)
        self.preview_count_label.setText(
            f"Folders: {result.folder_count}   |   Files: {result.file_count}"
        )
        self._update_last_refresh()

        if self.current_input_mode == "path":
            self.source_label.setText("Source: Folder Path")
        else:
            self.source_label.setText("Source: Manifest")

    def on_path_change(self):
        if self.current_input_mode != "path":
            self.current_input_mode = "path"
            self.current_manifest_data = None
            self.auto_refresh_checkbox.setEnabled(True)
            self.source_label.setText("Source: Folder Path")
        self.path_debounce_timer.start(700)

    def on_metadata_change(self):
        if self.generated_plain_text:
            self.generated_preview_html = self._compose_preview_html(
                self.generated_html
            )
            self.generated_copy_html = self._compose_copy_rich_html(
                self.generated_clipboard_html
            )
            self.generated_copy_plain_text = self._compose_copy_plain_text(
                self.generated_plain_text
            )
            self.preview_box.setHtml(self.generated_preview_html)

            existing_notes = self.get_submission_notes_map()
            shot_names = self._extract_shot_names_from_lines(self.current_result_lines)
            if shot_names != self.current_shot_names:
                self._rebuild_shot_notes_panel(shot_names)
                for shot_name, text in existing_notes.items():
                    editor = self.shot_note_edits.get(shot_name)
                    if editor is not None and text:
                        editor.setPlainText(text)
                self.generated_preview_html = self._compose_preview_html(
                    self.generated_html
                )
                self.generated_copy_html = self._compose_copy_rich_html(
                    self.generated_clipboard_html
                )
                self.generated_copy_plain_text = self._compose_copy_plain_text(
                    self.generated_plain_text
                )
                self.preview_box.setHtml(self.generated_preview_html)

    def on_auto_refresh_toggle(self, checked: bool):
        self.hint_label.setText(
            "Auto-refresh is ON" if checked else "Auto-refresh is OFF"
        )

    def toggle_advanced_options(self):
        is_open = self.advanced_toggle_btn.isChecked()
        self.advanced_frame.setVisible(is_open)
        self.advanced_toggle_btn.setText(
            "▼ Advanced Options" if is_open else "▶ Advanced Options"
        )

    def _debounced_path_preview(self):
        if self.current_input_mode != "path":
            return

        path = self.path_edit.text().strip()

        if not path:
            self.clear_preview()
            self.set_status("Ready.")
            return

        if os.path.isdir(path):
            self.start_preview_job(path, auto=True)
        else:
            self.clear_preview()
            self.set_status("Waiting for valid folder path...")

    def auto_refresh_tick(self):
        if self.current_input_mode != "path":
            return

        if not self.auto_refresh_checkbox.isChecked():
            return

        path = self.path_edit.text().strip()

        if (
            path
            and os.path.isdir(path)
            and self.generated_plain_text
            and not self.loading_active
        ):
            current_sig = MailPrepLogic.path_signature(path)
            if current_sig != self.last_preview_signature:
                self.start_preview_job(path, auto=True)

    def start_preview_job(self, source_data, auto: bool = False):
        if self.loading_active:
            return

        if self.current_input_mode == "path":
            path = str(source_data).strip() if source_data is not None else ""

            if not path:
                if not auto:
                    QMessageBox.warning(
                        self, "Warning", "Please select a folder first."
                    )
                self.clear_preview()
                self.set_status("Select a folder first.")
                return

            if not os.path.isdir(path):
                if not auto:
                    QMessageBox.critical(
                        self, "Error", "Selected path is not a valid folder."
                    )
                self.clear_preview()
                self.set_status("Invalid folder path.")
                return

        elif self.current_input_mode == "manifest":
            if not source_data:
                QMessageBox.warning(self, "Warning", "Load a manifest first.")
                self.clear_preview()
                self.set_status("Load a manifest first.")
                return

        self.current_job_id += 1
        job_id = self.current_job_id

        self.set_loading(True, "Scanning folder...")
        self.set_status(
            "Generating preview..." if not auto else "Auto refreshing preview..."
        )

        self.worker_thread = QThread()
        self.worker = PreviewWorker(
            self.current_input_mode, source_data, self.current_theme
        )
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(
            lambda result: self.finish_preview_job(result, source_data, job_id, auto)
        )
        self.worker.failed.connect(
            lambda error: self.fail_preview_job(error, source_data, job_id, auto)
        )

        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)

        self.worker_thread.start()

    def finish_preview_job(
        self, result: BuildResult, source_data, job_id: int, auto: bool
    ):
        if job_id != self.current_job_id:
            return

        self.set_loading(False)

        if self.current_input_mode == "path":
            current_path = self.path_edit.text().strip()
            source_path = str(source_data).strip() if source_data is not None else ""
            if current_path != source_path:
                return

        self.render_preview(result)
        self.set_status(
            "Preview auto-refreshed." if auto else "Preview generated successfully."
        )

    def fail_preview_job(self, error: str, source_data, job_id: int, auto: bool):
        if job_id != self.current_job_id:
            return

        self.set_loading(False)

        if not auto:
            QMessageBox.critical(self, "Error", f"Failed to generate preview:\n{error}")

        self.set_status(f"Error while generating preview: {error}")

    def generate_preview(self, auto: bool = False):
        if self.current_input_mode == "path":
            source = self.path_edit.text().strip()
        elif self.current_input_mode == "manifest":
            source = self.current_manifest_data
        else:
            source = None

        self.start_preview_job(source, auto=auto)

    def copy_rich_text(self):
        if not self.generated_copy_html:
            QMessageBox.warning(self, "Warning", "Generate preview first.")
            return

        try:
            mime = QMimeData()
            mime.setHtml(self.generated_copy_html)
            mime.setText(self.generated_copy_plain_text)
            mime.setData("text/html", self.generated_copy_html.encode("utf-8"))

            clipboard = QGuiApplication.clipboard()
            clipboard.setMimeData(mime, QClipboard.Clipboard)

            try:
                mime2 = QMimeData()
                mime2.setHtml(self.generated_copy_html)
                mime2.setText(self.generated_copy_plain_text)
                mime2.setData("text/html", self.generated_copy_html.encode("utf-8"))
                clipboard.setMimeData(mime2, QClipboard.Selection)
            except Exception:
                pass

            self.copy_rich_btn.setText("✅ Rich Text Copied")
            self.copy_feedback_timer.start(self.COPY_FEEDBACK_MS)
            self.set_status("Rich text copied.")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Copy failed:\n{exc}")
            self.set_status("Rich text copy failed.")

    def copy_html_source(self):
        if not self.generated_copy_html:
            QMessageBox.warning(self, "Warning", "Generate preview first.")
            return

        try:
            clipboard = QGuiApplication.clipboard()
            clipboard.setText(self.generated_copy_html, QClipboard.Clipboard)

            try:
                clipboard.setText(self.generated_copy_html, QClipboard.Selection)
            except Exception:
                pass

            self.copy_html_btn.setText("✅ HTML Copied")
            self.copy_feedback_timer.start(self.COPY_FEEDBACK_MS)
            self.set_status("HTML source copied.")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Copy failed:\n{exc}")
            self.set_status("HTML source copy failed.")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.current_input_mode = "path"
            self.current_manifest_data = None
            self.auto_refresh_checkbox.setEnabled(True)
            self.source_label.setText("Source: Folder Path")
            self.path_edit.setText(folder)

    def load_manifest_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Load Manifest", "", "JSON Files (*.json)"
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            if not isinstance(data, dict) or "items" not in data:
                raise ValueError("Invalid manifest format. Missing 'items'.")

            self.current_manifest_data = data
            self.current_input_mode = "manifest"
            self.auto_refresh_checkbox.setEnabled(False)
            self.source_label.setText("Source: Manifest")
            self.start_preview_job(data, auto=False)
            self.set_status("Manifest loaded successfully.")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to load manifest:\n{exc}")
            self.set_status("Manifest load failed.")

    def export_manifest_file(self):
        if not self.current_result_lines:
            QMessageBox.warning(self, "Warning", "Generate preview first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Manifest", "", "JSON Files (*.json)"
        )
        if not file_path:
            return

        try:
            data = MailPrepLogic.lines_to_manifest(self.current_result_lines)
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)

            self.set_status("Manifest exported.")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to export manifest:\n{exc}")
            self.set_status("Manifest export failed.")

    def open_preview_dialog(self):
        if not self.generated_preview_html:
            QMessageBox.warning(self, "Warning", "Generate preview first.")
            return

        dialog = PreviewDialog(
            self,
            html_content=self.generated_preview_html,
            theme_name=self.current_theme,
        )
        dialog.exec_()

    def closeEvent(self, event):
        try:
            self.path_debounce_timer.stop()
            self.auto_refresh_timer.stop()
            self.copy_feedback_timer.stop()
        except Exception:
            pass
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    window = MailPrepWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
