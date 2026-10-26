"""
SmartSort — Intelligent File Organizer for Windows
A professional, light-themed desktop application to analyze, categorize,
and organize files safely and efficiently.

All-in-one Python file: Application logic, UI, scanning, categorization,
smooth 60 FPS animations, progress tracking, in-place organization,
custom export to dedicated folders, undo, and collision-safe file handling.
"""

import os
import sys
import math
import time
import queue
import shutil
import datetime
import threading
import subprocess
import traceback
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

# Enable High-DPI awareness on Windows before creating any Tkinter widgets
if sys.platform.startswith("win"):
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional dependencies handled gracefully
try:
    from PIL import Image, ImageTk, ImageOps
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

try:
    import send2trash
    HAS_SEND2TRASH = True
except ImportError:
    HAS_SEND2TRASH = False


# ============================================================================
# DESIGN SYSTEM & COLOR PALETTE (Light Theme)
# ============================================================================
COLOR_BG_MAIN = "#F7F9FC"          # Clean off-white app background
COLOR_BG_CARD = "#FFFFFF"          # Pure white cards
COLOR_BORDER = "#E2E8F0"           # Subtle soft grey borders
COLOR_BORDER_LIGHT = "#F1F5F9"     # Very soft divider
COLOR_PRIMARY = "#0D9488"          # Fresh Teal accent
COLOR_PRIMARY_HOVER = "#0F766E"    # Deep Teal for hover
COLOR_PRIMARY_LIGHT = "#F0FDFA"    # Soft teal wash
COLOR_PRIMARY_TINT = "#CCFBF1"     # Selection highlight
COLOR_TEXT_MAIN = "#172B4D"        # Deep slate/navy primary text
COLOR_TEXT_MUTED = "#64748B"       # Slate secondary text
COLOR_TEXT_LIGHT = "#94A3B8"       # Light muted text
COLOR_SUCCESS = "#16A34A"          # Green
COLOR_WARNING = "#D97706"          # Amber / Orange
COLOR_DANGER = "#DC2626"           # Red
COLOR_ROW_ALT = "#F8FAFC"          # Alternating table row

# Status Badge Colors (Background, Foreground)
STATUS_COLORS = {
    "Ready": ("#DCFCE7", "#166534"),
    "Scanning": ("#E0F2FE", "#075985"),
    "Organizing": ("#FEF3C7", "#92400E"),
    "Exporting": ("#EDE9FE", "#6D28D9"),
    "Completed": ("#DCFCE7", "#166534"),
    "Cancelled": ("#FEE2E2", "#991B1B"),
    "Ready to Move": ("#DCFCE7", "#166534"),
    "Already in Place": ("#F1F5F9", "#475569"),
    "Conflict - Renamed": ("#FEF3C7", "#92400E"),
    "Excluded": ("#E2E8F0", "#64748B"),
    "Failed": ("#FEE2E2", "#991B1B"),
    "Moved": ("#DCFCE7", "#166534"),
    "Copied": ("#DCFCE7", "#166534"),
}

# Typography font choices
FONT_FAMILY = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 9)
FONT_SECTION = (FONT_FAMILY, 11, "bold")
FONT_CARD_HEADER = (FONT_FAMILY, 8, "bold")
FONT_STAT_BIG = (FONT_FAMILY, 16, "bold")
FONT_STAT_SUB = (FONT_FAMILY, 8)
FONT_BODY = (FONT_FAMILY, 9)
FONT_BODY_BOLD = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 8)
FONT_MONO = ("Consolas", 8)


# ============================================================================
# FILE CATEGORIES DEFINITION
# ============================================================================
CATEGORIES: Dict[str, List[str]] = {
    "Images": [
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
        ".tiff", ".tif", ".svg", ".ico", ".heic", ".heif", ".avif"
    ],
    "Videos": [
        ".mp4", ".mkv", ".mov", ".avi", ".wmv", ".flv",
        ".webm", ".m4v", ".mpeg", ".mpg", ".3gp"
    ],
    "Audio": [
        ".mp3", ".wav", ".aac", ".flac", ".ogg", ".m4a",
        ".wma", ".opus", ".mid", ".midi"
    ],
    "Documents": [
        ".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".tex"
    ],
    "Spreadsheets": [
        ".xls", ".xlsx", ".csv", ".ods", ".xlsm"
    ],
    "Presentations": [
        ".ppt", ".pptx", ".odp", ".key"
    ],
    "Archives": [
        ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz",
        ".tar.gz", ".tar.bz2", ".tar.xz", ".tar.zst"
    ],
    "Installers": [
        ".exe", ".msi", ".msix", ".appx"
    ],
    "Disk Images": [
        ".iso", ".img", ".vhd", ".vhdx"
    ],
    "Code": [
        ".py", ".js", ".ts", ".jsx", ".tsx", ".html", ".css",
        ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".php",
        ".go", ".rs", ".sql", ".json", ".xml", ".yaml", ".yml",
        ".sh", ".bat", ".ps1"
    ],
    "Fonts": [
        ".ttf", ".otf", ".woff", ".woff2"
    ],
    "Books": [
        ".epub", ".mobi", ".azw", ".azw3"
    ],
    "Other": []
}

CATEGORY_ICONS: Dict[str, str] = {
    "Images": "🖼️",
    "Videos": "🎬",
    "Audio": "🎵",
    "Documents": "📄",
    "Spreadsheets": "📊",
    "Presentations": "📽️",
    "Archives": "📦",
    "Installers": "⚙️",
    "Disk Images": "💿",
    "Code": "💻",
    "Fonts": "🔤",
    "Books": "📚",
    "Other": "📁"
}

# Reverse lookup table for O(1) extension category lookup
EXTENSION_TO_CATEGORY: Dict[str, str] = {}
for cat_name, ext_list in CATEGORIES.items():
    for ext in ext_list:
        EXTENSION_TO_CATEGORY[ext.lower()] = cat_name

# Multi-part compound extensions to check prior to single suffix
COMPOUND_EXTENSIONS: Tuple[str, ...] = (
    ".tar.gz", ".tar.bz2", ".tar.xz", ".tar.zst"
)


# ============================================================================
# DATA STRUCTURES
# ============================================================================
@dataclass
class FileItem:
    original_path: Path
    filename: str
    extension: str
    category: str
    size_bytes: int
    mtime: float
    relative_path: str
    proposed_destination: Optional[Path] = None
    status: str = "Ready to Move"  # "Ready to Move", "Already in Place", "Conflict - Renamed", "Excluded", "Failed", "Moved", "Copied"
    is_included: bool = True
    error_message: str = ""


@dataclass
class UndoRecord:
    original_path: Path
    destination_path: Path
    file_size: int
    category: str
    operation: str = "move"  # "move" or "copy"


# ============================================================================
# HELPER UTILITIES
# ============================================================================
def format_size(size_bytes: int) -> str:
    """Format bytes into human readable binary units."""
    if size_bytes <= 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    val = float(size_bytes)
    while val >= 1024.0 and i < len(units) - 1:
        val /= 1024.0
        i += 1
    if i == 0:
        return f"{int(val)} B"
    elif val < 10.0:
        return f"{val:.2f} {units[i]}"
    else:
        return f"{val:.1f} {units[i]}"


def format_time(seconds: float) -> str:
    """
    Format elapsed or estimated seconds into a user-friendly string.
    Example: 12 sec, 2 min 15 sec, 1 hr 8 min.
    """
    if seconds < 0 or math.isinf(seconds) or math.isnan(seconds):
        return "--"
    sec = int(round(seconds))
    if sec < 60:
        return f"{sec} sec"
    minutes = sec // 60
    rem_sec = sec % 60
    if minutes < 60:
        return f"{minutes} min {rem_sec} sec" if rem_sec else f"{minutes} min"
    hours = minutes // 60
    rem_min = minutes % 60
    return f"{hours} hr {rem_min} min" if rem_min else f"{hours} hr"


def get_file_extension(filename: str) -> str:
    """Extract file extension safely, checking compound extensions first."""
    lower = filename.lower()
    for compound in COMPOUND_EXTENSIONS:
        if lower.endswith(compound):
            return compound
    p = Path(filename)
    return p.suffix.lower()


def detect_category(filename: str) -> str:
    """Categorize file based on its extension, supporting compound and single suffixes."""
    ext = get_file_extension(filename)
    if ext in EXTENSION_TO_CATEGORY:
        return EXTENSION_TO_CATEGORY[ext]
    single_ext = Path(filename).suffix.lower()
    return EXTENSION_TO_CATEGORY.get(single_ext, "Other")


def truncate_path(path_str: str, max_chars: int = 55) -> str:
    """Truncate a path with ellipsis in the middle or end if too long."""
    if len(path_str) <= max_chars:
        return path_str
    half = (max_chars - 3) // 2
    return f"{path_str[:half]}...{path_str[-half:]}"


def is_system_path(path: Path) -> bool:
    """Check if the given path is a protected Windows system directory."""
    try:
        resolved = path.resolve()
        name_lower = resolved.name.lower()
        if name_lower in ("$recycle.bin", "system volume information", "recovery", "windows"):
            return True
        # Check against Windows root system folder
        win_dir = os.environ.get("SystemRoot", "C:\\Windows")
        if str(resolved).lower().startswith(win_dir.lower()):
            return True
    except Exception:
        pass
    return False


def get_target_subfolder(filename: str, category: str, extension: str, structure: str) -> str:
    """
    Determine subfolder path based on user's desired organizational structure:
    - 'category': Images, Videos, Documents, Audio...
    - 'extension': Separate folder for each format (PDF, DOCX, JPG, PNG, MP4, etc.)
    - 'nested': Category / Extension (e.g. Documents/PDF, Images/JPG)
    """
    ext_clean = extension.lstrip(".").upper() if extension else "NO_EXT"
    if structure == "extension":
        return ext_clean
    elif structure == "nested":
        return f"{category}/{ext_clean}"
    else:  # default 'category'
        return category


def plan_destinations(root: Path, file_items: List[FileItem], structure: str = "category") -> List[FileItem]:
    """
    Build proposed destination paths with full collision protection.
    1. Detects files already located inside their target destination folder.
    2. Generates unique safe names 'name (1).ext' for any destination conflicts.
    3. Guarantees no two files will be assigned the same destination.
    """
    reserved_names: Dict[str, Set[str]] = {}

    # Inspect current disk contents of each target folder
    for cat_name in CATEGORIES.keys():
        cat_dir = root / cat_name
        reserved_set: Set[str] = set()
        if cat_dir.is_dir():
            try:
                with os.scandir(cat_dir) as it:
                    for entry in it:
                        if entry.is_file():
                            reserved_set.add(entry.name.lower())
            except Exception:
                pass
        reserved_names[cat_name] = reserved_set

    # Pass 1: Identify files that are ALREADY in the correct category folder
    for item in file_items:
        subfolder = get_target_subfolder(item.filename, item.category, item.extension, structure)
        expected_dest = root / subfolder / item.filename
        try:
            if item.original_path.resolve() == expected_dest.resolve():
                item.proposed_destination = expected_dest
                item.status = "Already in Place"
                item.is_included = False  # By default, do not re-move already sorted files
                reserved_names.setdefault(subfolder, set()).add(item.filename.lower())
                continue
        except Exception:
            pass
        item.status = "Pending"

    # Pass 2: Calculate safe destination for files requiring movement
    for item in file_items:
        if item.status == "Already in Place":
            continue

        subfolder = get_target_subfolder(item.filename, item.category, item.extension, structure)
        dest_dir = root / subfolder
        sub_set = reserved_names.setdefault(subfolder, set())

        orig_name = item.filename
        ext = item.extension

        # Extract base stem without compound extension
        if ext and orig_name.lower().endswith(ext):
            stem = orig_name[:-len(ext)]
        else:
            stem = Path(orig_name).stem

        cand_name = orig_name
        if cand_name.lower() not in sub_set:
            sub_set.add(cand_name.lower())
            item.proposed_destination = dest_dir / cand_name
            item.status = "Ready to Move"
            item.is_included = True
        else:
            # Collision resolution: filename (1).ext
            counter = 1
            cand_name = f"{stem} ({counter}){ext}"
            while cand_name.lower() in sub_set:
                counter += 1
                cand_name = f"{stem} ({counter}){ext}"
            sub_set.add(cand_name.lower())
            item.proposed_destination = dest_dir / cand_name
            item.status = "Conflict - Renamed"
            item.is_included = True

    return file_items


# ============================================================================
# CUSTOM UI WIDGETS (Modern Light Theme & Smooth 60 FPS Animations)
# ============================================================================
class Card(tk.Frame):
    """Clean white card with subtle grey border and standard padding."""
    def __init__(self, parent, **kwargs):
        bg = kwargs.pop("bg", COLOR_BG_CARD)
        border = kwargs.pop("highlightbackground", COLOR_BORDER)
        thickness = kwargs.pop("highlightthickness", 1)
        super().__init__(
            parent,
            bg=bg,
            highlightbackground=border,
            highlightthickness=thickness,
            bd=0,
            **kwargs
        )


class ModernButton(tk.Button):
    """
    Polished flat button supporting primary (teal), secondary (soft grey),
    danger (red tint), and outline styles with active hover states.
    """
    def __init__(self, parent, text="", command=None, style="secondary", **kwargs):
        self.btn_style = style
        self.custom_cmd = command

        # Determine color definitions
        if style == "primary":
            self.bg_normal = COLOR_PRIMARY
            self.fg_normal = "#FFFFFF"
            self.bg_hover = COLOR_PRIMARY_HOVER
            self.fg_hover = "#FFFFFF"
            border_w = 0
            border_c = COLOR_PRIMARY
            font = FONT_BODY_BOLD
        elif style == "export":
            self.bg_normal = "#7C3AED"        # Vivid modern purple
            self.fg_normal = "#FFFFFF"
            self.bg_hover = "#6D28D9"         # Deep violet on hover
            self.fg_hover = "#FFFFFF"
            border_w = 0
            border_c = "#7C3AED"
            font = FONT_BODY_BOLD
        elif style == "danger":
            self.bg_normal = "#FEE2E2"
            self.fg_normal = "#991B1B"
            self.bg_hover = "#FCA5A5"
            self.fg_hover = "#7F1D1D"
            border_w = 1
            border_c = "#FECACA"
            font = FONT_BODY_BOLD
        elif style == "outline":
            self.bg_normal = "#FFFFFF"
            self.fg_normal = COLOR_PRIMARY
            self.bg_hover = COLOR_PRIMARY_LIGHT
            self.fg_hover = COLOR_PRIMARY_HOVER
            border_w = 1
            border_c = COLOR_PRIMARY
            font = FONT_BODY_BOLD
        else:  # secondary
            self.bg_normal = "#F1F5F9"
            self.fg_normal = "#1E293B"
            self.bg_hover = "#E2E8F0"
            self.fg_hover = "#0F172A"
            border_w = 1
            border_c = "#CBD5E1"
            font = FONT_BODY

        pad_x = kwargs.pop("padx", 12)
        pad_y = kwargs.pop("pady", 6)

        super().__init__(
            parent,
            text=text,
            command=command,
            bg=self.bg_normal,
            fg=self.fg_normal,
            activebackground=self.bg_hover,
            activeforeground=self.fg_hover,
            font=font,
            relief="flat",
            bd=0,
            highlightthickness=border_w,
            highlightbackground=border_c,
            cursor="hand2",
            padx=pad_x,
            pady=pad_y,
            **kwargs
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, event):
        if str(self["state"]) != "disabled":
            self.configure(bg=self.bg_hover, fg=self.fg_hover)

    def _on_leave(self, event):
        if str(self["state"]) != "disabled":
            self.configure(bg=self.bg_normal, fg=self.fg_normal)

    def set_enabled(self, enabled: bool):
        if enabled:
            self.configure(state="normal", cursor="hand2", bg=self.bg_normal, fg=self.fg_normal)
        else:
            self.configure(state="disabled", cursor="arrow", bg="#F1F5F9", fg=COLOR_TEXT_LIGHT)


class AnimatedBadge(tk.Frame):
    """
    Rounded-style colored status pill with smooth pulsing/cycling animation
    during active operations (Scanning, Organizing, Exporting).
    """
    def __init__(self, parent, text="Ready", status_key="Ready", **kwargs):
        super().__init__(parent, bg=COLOR_BG_MAIN, **kwargs)
        self.status_key = status_key
        self.base_text = text
        self.dot_counter = 0
        self.anim_job = None

        bg_col, fg_col = STATUS_COLORS.get(status_key, ("#F1F5F9", "#475569"))
        self.lbl = tk.Label(
            self,
            text=f"● {text}",
            bg=bg_col,
            fg=fg_col,
            font=FONT_SMALL,
            padx=10,
            pady=4,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER
        )
        self.lbl.pack(fill="both", expand=True)

    def update_status(self, text: str, status_key: str):
        self.status_key = status_key
        self.base_text = text
        bg_col, fg_col = STATUS_COLORS.get(status_key, ("#F1F5F9", "#475569"))
        self.lbl.configure(text=f"● {text}", bg=bg_col, fg=fg_col)

        if status_key in ("Scanning", "Organizing", "Exporting"):
            if not self.anim_job:
                self._pulse_loop()
        else:
            if self.anim_job:
                self.after_cancel(self.anim_job)
                self.anim_job = None
            self.lbl.configure(text=f"● {text}")

    def _pulse_loop(self):
        if self.status_key in ("Scanning", "Organizing", "Exporting"):
            dots = "." * (self.dot_counter % 4)
            self.lbl.configure(text=f"● {self.base_text}{dots}")
            self.dot_counter += 1
            self.anim_job = self.after(350, self._pulse_loop)
        else:
            self.anim_job = None


class SmoothProgressBar(tk.Canvas):
    """
    Smooth, anti-aliased 60 FPS animated progress bar with rounded ends
    and smooth easing interpolation toward the target progress value.
    """
    def __init__(self, parent, height=16, bg=COLOR_BG_CARD, trough_color=COLOR_BORDER, fill_color=COLOR_PRIMARY, **kwargs):
        super().__init__(parent, height=height, bg=bg, highlightthickness=0, bd=0, **kwargs)
        self.bar_height = height
        self.trough_color = trough_color
        self.fill_color = fill_color
        self.current_val = 0.0
        self.target_val = 0.0
        self.is_indeterminate = False
        self.shimmer_pos = 0.0
        self.shimmer_dir = 1
        self.anim_active = False

        self.bind("<Configure>", lambda e: self._render())
        self._start_anim()

    def _start_anim(self):
        if not self.anim_active:
            self.anim_active = True
            self._tick()

    def set_value(self, val: float):
        """Set target progress (0.0 to 100.0)."""
        self.is_indeterminate = False
        self.target_val = max(0.0, min(100.0, float(val)))

    def start_indeterminate(self):
        self.is_indeterminate = True

    def stop_indeterminate(self):
        self.is_indeterminate = False

    def _tick(self):
        need_render = False
        if self.is_indeterminate:
            self.shimmer_pos += 0.025 * self.shimmer_dir
            if self.shimmer_pos >= 0.75:
                self.shimmer_dir = -1
            elif self.shimmer_pos <= 0.0:
                self.shimmer_dir = 1
            need_render = True
        else:
            diff = self.target_val - self.current_val
            if abs(diff) > 0.15:
                # Smooth 60 FPS easing
                self.current_val += diff * 0.22
                need_render = True
            elif self.current_val != self.target_val:
                self.current_val = self.target_val
                need_render = True

        if need_render:
            self._render()

        self.after(16, self._tick)

    def _render(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.bar_height
        if w <= 10:
            return

        r = h // 2
        # Trough track
        self.create_arc(0, 0, h, h, start=90, extent=180, fill=self.trough_color, outline="")
        self.create_rectangle(r, 0, w - r, h, fill=self.trough_color, outline="")
        self.create_arc(w - h, 0, w, h, start=270, extent=180, fill=self.trough_color, outline="")

        if self.is_indeterminate:
            bar_w = max(40, w * 0.3)
            x1 = self.shimmer_pos * (w - bar_w)
            x2 = x1 + bar_w
            self.create_rectangle(x1, 0, x2, h, fill=self.fill_color, outline="")
        else:
            pct = self.current_val / 100.0
            fill_w = max(0, w * pct)
            if fill_w > h:
                self.create_arc(0, 0, h, h, start=90, extent=180, fill=self.fill_color, outline="")
                self.create_rectangle(r, 0, fill_w - r, h, fill=self.fill_color, outline="")
                self.create_arc(fill_w - h, 0, fill_w, h, start=270, extent=180, fill=self.fill_color, outline="")
            elif fill_w > 0:
                self.create_rectangle(0, 0, fill_w, h, fill=self.fill_color, outline="")


class StatCard(Card):
    """Dashboard KPI card with header, big bold value, and subtle subtitle."""
    def __init__(self, parent, title="FILES", value="0", subtitle="Discovered", accent=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.accent_color = accent or COLOR_PRIMARY

        top_bar = tk.Frame(self, bg=self.accent_color, height=3)
        top_bar.pack(fill="x", side="top")

        content = tk.Frame(self, bg=COLOR_BG_CARD, padx=10, pady=8)
        content.pack(fill="both", expand=True)

        self.lbl_title = tk.Label(content, text=title.upper(), font=FONT_CARD_HEADER, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD)
        self.lbl_title.pack(anchor="w")

        self.lbl_value = tk.Label(content, text=value, font=FONT_STAT_BIG, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD)
        self.lbl_value.pack(anchor="w", pady=(1, 1))

        self.lbl_sub = tk.Label(content, text=subtitle, font=FONT_STAT_SUB, fg=self.accent_color, bg=COLOR_BG_CARD)
        self.lbl_sub.pack(anchor="w")

    def update_card(self, value: str, subtitle: Optional[str] = None):
        self.lbl_value.configure(text=value)
        if subtitle is not None:
            self.lbl_sub.configure(text=subtitle)


# ============================================================================
# BACKGROUND THREAD WORKERS
# ============================================================================
class ScanWorker(threading.Thread):
    """
    Recursively scans the directory and gathers file metadata,
    reporting real-time progress to the UI queue safely.
    """
    def __init__(self, root_dir: Path, include_subfolders: bool, msg_queue: queue.Queue, cancel_event: threading.Event):
        super().__init__(daemon=True)
        self.root_dir = root_dir
        self.include_subfolders = include_subfolders
        self.msg_queue = msg_queue
        self.cancel_event = cancel_event

    def run(self):
        start_time = time.perf_counter()
        raw_files: List[Path] = []
        folder_count = 0
        inaccessible_count = 0

        # Protect the script file itself from being moved
        script_file = None
        try:
            script_file = Path(sys.argv[0]).resolve()
        except Exception:
            pass

        # Phase 1: Fast directory traversal
        dirs_to_visit = [self.root_dir]
        visited_dirs: Set[Path] = set()

        last_update = time.perf_counter()

        while dirs_to_visit and not self.cancel_event.is_set():
            curr_dir = dirs_to_visit.pop(0)
            try:
                resolved = curr_dir.resolve()
                if resolved in visited_dirs:
                    continue
                visited_dirs.add(resolved)

                with os.scandir(curr_dir) as it:
                    for entry in it:
                        if self.cancel_event.is_set():
                            break
                        try:
                            # Skip symlinks and junction points to prevent loops
                            if entry.is_symlink():
                                continue

                            if entry.is_dir(follow_symlinks=False):
                                # Skip hidden system directories
                                if entry.name.startswith("$") or entry.name in ("System Volume Information", "$RECYCLE.BIN", ".git"):
                                    continue
                                folder_count += 1
                                if self.include_subfolders:
                                    dirs_to_visit.append(Path(entry.path))

                            elif entry.is_file(follow_symlinks=False):
                                p = Path(entry.path)
                                # Never include SmartSort script itself
                                if script_file and p.resolve() == script_file:
                                    continue
                                if p.name.lower() in ("smartsort.py", "smartsort.exe"):
                                    continue
                                raw_files.append(p)

                        except Exception:
                            inaccessible_count += 1

                # Periodically emit discovery counts
                now = time.perf_counter()
                if now - last_update > 0.08:
                    self.msg_queue.put({
                        "type": "scan_discovery",
                        "discovered_files": len(raw_files),
                        "folders": folder_count,
                        "elapsed": now - start_time
                    })
                    last_update = now

            except Exception:
                inaccessible_count += 1

        if self.cancel_event.is_set():
            self.msg_queue.put({"type": "scan_cancelled"})
            return

        total_files = len(raw_files)
        file_items: List[FileItem] = []
        total_size = 0
        cat_counts: Dict[str, int] = {c: 0 for c in CATEGORIES.keys()}
        cat_sizes: Dict[str, int] = {c: 0 for c in CATEGORIES.keys()}

        # Phase 2: Metadata extraction & categorization with determinate progress
        for i, path in enumerate(raw_files, start=1):
            if self.cancel_event.is_set():
                self.msg_queue.put({"type": "scan_cancelled"})
                return

            try:
                stat_info = path.stat()
                sz = stat_info.st_size
                mtime = stat_info.st_mtime
            except Exception:
                sz = 0
                mtime = 0
                inaccessible_count += 1

            total_size += sz
            cat = detect_category(path.name)
            ext = get_file_extension(path.name)
            try:
                rel = path.relative_to(self.root_dir).as_posix()
            except Exception:
                rel = path.name

            item = FileItem(
                original_path=path,
                filename=path.name,
                extension=ext,
                category=cat,
                size_bytes=sz,
                mtime=mtime,
                relative_path=rel
            )
            file_items.append(item)
            cat_counts[cat] += 1
            cat_sizes[cat] += sz

            now = time.perf_counter()
            if now - last_update > 0.05 or i == total_files:
                elapsed = now - start_time
                speed = i / elapsed if elapsed > 0 else 0
                rem_files = total_files - i
                eta = (elapsed / i) * rem_files if i > 0 else 0
                self.msg_queue.put({
                    "type": "scan_progress",
                    "processed": i,
                    "total": total_files,
                    "current_file": path.name,
                    "elapsed": elapsed,
                    "speed": speed,
                    "eta": eta
                })
                last_update = now

        # Phase 3: Collision resolution and destination planning
        file_items = plan_destinations(self.root_dir, file_items)
        elapsed_total = time.perf_counter() - start_time

        self.msg_queue.put({
            "type": "scan_complete",
            "items": file_items,
            "folders": folder_count,
            "total_size": total_size,
            "inaccessible": inaccessible_count,
            "elapsed": elapsed_total,
            "cat_counts": cat_counts,
            "cat_sizes": cat_sizes
        })


class OrganizeWorker(threading.Thread):
    """
    Executes in-place file organization or dry-run simulation,
    keeping a log for session undo operations.
    """
    def __init__(self, root_dir: Path, items: List[FileItem], dry_run: bool, msg_queue: queue.Queue, cancel_event: threading.Event):
        super().__init__(daemon=True)
        self.root_dir = root_dir
        self.items = items
        self.dry_run = dry_run
        self.msg_queue = msg_queue
        self.cancel_event = cancel_event

    def run(self):
        start_time = time.perf_counter()
        to_process = [
            it for it in self.items
            if it.is_included and it.status in ("Ready to Move", "Conflict - Renamed", "Moved (Simulated)")
        ]
        total = len(to_process)

        success_count = 0
        skipped_count = len(self.items) - total
        failed_count = 0
        bytes_moved = 0
        undo_records: List[UndoRecord] = []
        errors: List[Tuple[str, str]] = []

        last_update = time.perf_counter()

        for i, item in enumerate(to_process, start=1):
            if self.cancel_event.is_set():
                break

            src = item.original_path
            dest = item.proposed_destination

            if not dest:
                failed_count += 1
                item.status = "Failed"
                item.error_message = "No destination specified"
                errors.append((src.name, "No destination specified"))
                continue

            try:
                if self.dry_run:
                    time.sleep(0.003)
                    success_count += 1
                    bytes_moved += item.size_bytes
                    item.status = "Moved (Simulated)"
                else:
                    if not src.exists():
                        failed_count += 1
                        item.status = "Failed"
                        item.error_message = "Source file no longer exists"
                        errors.append((src.name, "Source file missing"))
                        continue

                    dest.parent.mkdir(parents=True, exist_ok=True)

                    # Collision safeguard right before move
                    final_dest = dest
                    if final_dest.exists() and final_dest.resolve() != src.resolve():
                        counter = 1
                        stem = final_dest.stem
                        suffix = final_dest.suffix
                        cand = final_dest.parent / f"{stem} ({counter}){suffix}"
                        while cand.exists():
                            counter += 1
                            cand = final_dest.parent / f"{stem} ({counter}){suffix}"
                        final_dest = cand

                    shutil.move(str(src), str(final_dest))

                    undo_records.append(UndoRecord(
                        original_path=src,
                        destination_path=final_dest,
                        file_size=item.size_bytes,
                        category=item.category,
                        operation="move"
                    ))

                    item.status = "Moved"
                    success_count += 1
                    bytes_moved += item.size_bytes

            except Exception as e:
                failed_count += 1
                item.status = "Failed"
                err_msg = str(e)
                item.error_message = err_msg
                errors.append((src.name, err_msg))

            now = time.perf_counter()
            if now - last_update > 0.05 or i == total:
                elapsed = now - start_time
                speed = i / elapsed if elapsed > 0 else 0
                rem = total - i
                eta = (elapsed / i) * rem if i > 0 else 0
                self.msg_queue.put({
                    "type": "organize_progress",
                    "processed": i,
                    "total": total,
                    "current_file": src.name,
                    "success": success_count,
                    "skipped": skipped_count,
                    "failed": failed_count,
                    "elapsed": elapsed,
                    "speed": speed,
                    "eta": eta
                })
                last_update = now

        elapsed_total = time.perf_counter() - start_time
        was_cancelled = self.cancel_event.is_set()

        self.msg_queue.put({
            "type": "organize_complete",
            "cancelled": was_cancelled,
            "success": success_count,
            "skipped": skipped_count,
            "failed": failed_count,
            "bytes_moved": bytes_moved,
            "elapsed": elapsed_total,
            "undo_records": undo_records,
            "errors": errors,
            "dry_run": self.dry_run
        })


class ExportWorker(threading.Thread):
    """
    Exports and separates files into a dedicated target folder,
    supporting structure by Category, by File Extension (PDF, DOCX, MP4, etc.),
    or Nested (Category/Ext), with Copy or Move transfer modes.
    """
    def __init__(
        self,
        export_dir: Path,
        items: List[FileItem],
        structure: str,
        operation: str,
        filter_type: str,
        msg_queue: queue.Queue,
        cancel_event: threading.Event
    ):
        super().__init__(daemon=True)
        self.export_dir = export_dir
        self.items = items
        self.structure = structure
        self.operation = operation  # "copy" or "move"
        self.filter_type = filter_type
        self.msg_queue = msg_queue
        self.cancel_event = cancel_event

    def run(self):
        start_time = time.perf_counter()

        # Apply Export Filtering
        selected_items: List[FileItem] = []
        for it in self.items:
            ext_lower = it.extension.lower()
            cat = it.category

            if self.filter_type == "all":
                selected_items.append(it)
            elif self.filter_type == "selected_only" and it.is_included:
                selected_items.append(it)
            elif self.filter_type == "videos_only" and cat == "Videos":
                selected_items.append(it)
            elif self.filter_type == "images_only" and cat == "Images":
                selected_items.append(it)
            elif self.filter_type == "pdf_only" and ext_lower == ".pdf":
                selected_items.append(it)
            elif self.filter_type == "docx_only" and ext_lower in (".docx", ".doc"):
                selected_items.append(it)
            elif self.filter_type == "spreadsheets_only" and cat == "Spreadsheets":
                selected_items.append(it)
            elif self.filter_type == "archives_only" and cat == "Archives":
                selected_items.append(it)

        total = len(selected_items)
        success_count = 0
        failed_count = 0
        bytes_transferred = 0
        undo_records: List[UndoRecord] = []
        errors: List[Tuple[str, str]] = []
        reserved_names: Dict[str, Set[str]] = {}

        last_update = time.perf_counter()

        for i, item in enumerate(selected_items, start=1):
            if self.cancel_event.is_set():
                break

            src = item.original_path
            if not src.exists():
                failed_count += 1
                errors.append((item.filename, "Source file no longer exists"))
                continue

            subfolder = get_target_subfolder(item.filename, item.category, item.extension, self.structure)
            target_folder = self.export_dir / subfolder

            try:
                target_folder.mkdir(parents=True, exist_ok=True)
                sub_set = reserved_names.setdefault(subfolder, set())

                # Compute unique collision-safe destination
                stem = Path(item.filename).stem
                ext = item.extension
                cand_name = item.filename
                counter = 1

                while cand_name.lower() in sub_set or (target_folder / cand_name).exists():
                    cand_name = f"{stem} ({counter}){ext}"
                    counter += 1

                sub_set.add(cand_name.lower())
                final_dest = target_folder / cand_name

                # Execute Copy or Move
                if self.operation == "copy":
                    shutil.copy2(str(src), str(final_dest))
                    item.status = "Copied"
                else:  # move
                    shutil.move(str(src), str(final_dest))
                    item.status = "Moved"

                undo_records.append(UndoRecord(
                    original_path=src,
                    destination_path=final_dest,
                    file_size=item.size_bytes,
                    category=item.category,
                    operation=self.operation
                ))

                success_count += 1
                bytes_transferred += item.size_bytes

            except Exception as e:
                failed_count += 1
                errors.append((item.filename, str(e)))

            now = time.perf_counter()
            if now - last_update > 0.05 or i == total:
                elapsed = now - start_time
                speed = i / elapsed if elapsed > 0 else 0
                rem = total - i
                eta = (elapsed / i) * rem if i > 0 else 0
                self.msg_queue.put({
                    "type": "export_progress",
                    "processed": i,
                    "total": total,
                    "current_file": item.filename,
                    "success": success_count,
                    "failed": failed_count,
                    "elapsed": elapsed,
                    "speed": speed,
                    "eta": eta
                })
                last_update = now

        elapsed_total = time.perf_counter() - start_time
        was_cancelled = self.cancel_event.is_set()

        self.msg_queue.put({
            "type": "export_complete",
            "cancelled": was_cancelled,
            "success": success_count,
            "failed": failed_count,
            "bytes_transferred": bytes_transferred,
            "export_dir": self.export_dir,
            "operation": self.operation,
            "elapsed": elapsed_total,
            "undo_records": undo_records,
            "errors": errors
        })


class UndoWorker(threading.Thread):
    """Restores organized files back to their original locations."""
    def __init__(self, records: List[UndoRecord], root_dir: Path, msg_queue: queue.Queue):
        super().__init__(daemon=True)
        self.records = records
        self.root_dir = root_dir
        self.msg_queue = msg_queue

    def run(self):
        start_time = time.perf_counter()
        total = len(self.records)
        restored = 0
        failed = 0
        errors: List[Tuple[str, str]] = []

        # Undo in reverse order
        for i, rec in enumerate(reversed(self.records), start=1):
            src_dest = rec.destination_path
            orig_src = rec.original_path
            try:
                if src_dest.exists():
                    if rec.operation == "move":
                        orig_src.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(src_dest), str(orig_src))
                        restored += 1
                    else:  # was copied, delete exported copy safely
                        src_dest.unlink()
                        restored += 1

                    # Remove empty category folders
                    try:
                        cat_dir = src_dest.parent
                        if cat_dir.is_dir() and not any(cat_dir.iterdir()):
                            cat_dir.rmdir()
                    except Exception:
                        pass
                else:
                    failed += 1
                    errors.append((src_dest.name, "File not found at destination"))
            except Exception as e:
                failed += 1
                errors.append((src_dest.name, str(e)))

            self.msg_queue.put({
                "type": "undo_progress",
                "processed": i,
                "total": total,
                "current_file": orig_src.name
            })

        self.msg_queue.put({
            "type": "undo_complete",
            "restored": restored,
            "failed": failed,
            "total": total,
            "elapsed": time.perf_counter() - start_time,
            "errors": errors
        })


# ============================================================================
# DIALOGS & MODAL WINDOWS
# ============================================================================
class ExportDialog(tk.Toplevel):
    """
    Modern modal window allowing the user to configure custom export folder name,
    structure (By Category or By File Extension like PDF/DOCX/MP4 only),
    filters, and Copy vs Move transfer mode.
    """
    def __init__(self, parent, default_parent_dir: Path, total_files: int):
        super().__init__(parent)
        self.title("Export & Separate Files")
        self.configure(bg=COLOR_BG_MAIN)
        self.geometry("560x540")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.confirmed = False
        self.default_parent_dir = default_parent_dir
        self.target_parent = default_parent_dir

        # Auto-detect unique default folder name: SmartSort_Export, SmartSort_Export (1), etc.
        base_name = "SmartSort_Export"
        cand_name = base_name
        counter = 1
        while (default_parent_dir / cand_name).exists():
            cand_name = f"{base_name} ({counter})"
            counter += 1

        self.folder_name_var = tk.StringVar(value=cand_name)
        self.structure_var = tk.StringVar(value="extension")  # Default to user's desired "separate by type"
        self.operation_var = tk.StringVar(value="copy")       # Safe default: Copy
        self.filter_var = tk.StringVar(value="all")

        # Center
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        x = px + max(0, (pw - 560) // 2)
        y = py + max(0, (ph - 540) // 2)
        self.geometry(f"+{x}+{y}")

        # Header
        hdr = tk.Frame(self, bg=COLOR_BG_CARD, padx=20, pady=14, highlightbackground=COLOR_BORDER, highlightthickness=1)
        hdr.pack(fill="x", side="top")

        hdr_top = tk.Frame(hdr, bg=COLOR_BG_CARD)
        hdr_top.pack(anchor="w")
        tk.Label(hdr_top, text="📦", font=(FONT_FAMILY, 16), bg=COLOR_BG_CARD).pack(side="left", padx=(0, 8))
        tk.Label(hdr_top, text="Export & Separate Files", font=FONT_TITLE, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(side="left")

        tk.Label(
            hdr,
            text="Choose folder name, type grouping (PDF only, Video only, etc.), and copy/move mode.",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_CARD
        ).pack(anchor="w", pady=(3, 0))

        # Main Body Card
        body = Card(self)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        inner = tk.Frame(body, bg=COLOR_BG_CARD, padx=16, pady=12)
        inner.pack(fill="both", expand=True)

        # 1. Folder Name & Location
        tk.Label(inner, text="1. Export Folder Name & Target Location", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")

        row_folder = tk.Frame(inner, bg=COLOR_BG_CARD)
        row_folder.pack(fill="x", pady=(4, 2))

        self.entry_name = tk.Entry(
            row_folder,
            textvariable=self.folder_name_var,
            font=FONT_BODY,
            bg="#F8FAFC",
            fg=COLOR_TEXT_MAIN,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            bd=0
        )
        self.entry_name.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=3)
        self.folder_name_var.trace_add("write", lambda *args: self._update_path_preview())

        ModernButton(row_folder, text="📁 Choose Location...", style="secondary", command=self._browse_location).pack(side="left")

        self.lbl_path_preview = tk.Label(
            inner,
            text="",
            font=FONT_SMALL,
            fg=COLOR_PRIMARY,
            bg=COLOR_BG_CARD
        )
        self.lbl_path_preview.pack(anchor="w", pady=(0, 10))
        self._update_path_preview()

        # 2. Separation Structure
        tk.Label(inner, text="2. Folder Separation Structure", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")

        struct_box = tk.Frame(inner, bg=COLOR_BG_CARD)
        struct_box.pack(fill="x", pady=(2, 10))

        tk.Radiobutton(
            struct_box,
            text="📄 By Specific Extension (PDF only, DOCX only, JPG only, MP4 only, etc.)",
            variable=self.structure_var,
            value="extension",
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD
        ).pack(anchor="w", pady=1)

        tk.Radiobutton(
            struct_box,
            text="📁 By Category (Images, Videos, Documents, Audio, Archives, Code...)",
            variable=self.structure_var,
            value="category",
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD
        ).pack(anchor="w", pady=1)

        tk.Radiobutton(
            struct_box,
            text="🗂️ Nested: Category + Extension (e.g. Documents/PDF, Images/JPG)",
            variable=self.structure_var,
            value="nested",
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD
        ).pack(anchor="w", pady=1)

        # 3. Filter Scope
        tk.Label(inner, text="3. Filter Files to Export", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")

        filter_row = tk.Frame(inner, bg=COLOR_BG_CARD)
        filter_row.pack(fill="x", pady=(2, 10))

        filter_options = [
            ("All Discovered Files", "all"),
            ("🎬 Only Videos (.mp4, .mkv, .mov...)", "videos_only"),
            ("🖼️ Only Images (.jpg, .png, .gif...)", "images_only"),
            ("📄 Only PDF Documents (.pdf)", "pdf_only"),
            ("📝 Only Word Documents (.docx, .doc)", "docx_only"),
            ("📊 Only Spreadsheets (.xlsx, .csv)", "spreadsheets_only"),
            ("📦 Only Archives (.zip, .rar, .tar.gz)", "archives_only"),
            ("✔️ Only Checked / Selected in Preview Table", "selected_only")
        ]

        self.cb_filter = ttk.Combobox(
            filter_row,
            values=[opt[0] for opt in filter_options],
            state="readonly",
            font=FONT_BODY,
            width=42
        )
        self.cb_filter.current(0)
        self.cb_filter.pack(side="left", fill="x", expand=True)
        self.filter_options_map = {opt[0]: opt[1] for opt in filter_options}

        # 4. Transfer Mode: Copy vs Move
        tk.Label(inner, text="4. Transfer Operation", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")

        op_box = tk.Frame(inner, bg=COLOR_BG_CARD)
        op_box.pack(fill="x", pady=(2, 6))

        tk.Radiobutton(
            op_box,
            text="📋 Copy to Export Folder (Safe: Keeps original files untouched)",
            variable=self.operation_var,
            value="copy",
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD
        ).pack(anchor="w")

        tk.Radiobutton(
            op_box,
            text="🚚 Move to Export Folder (Transfers files into the new location)",
            variable=self.operation_var,
            value="move",
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD
        ).pack(anchor="w")

        # Footer Actions
        footer = tk.Frame(self, bg=COLOR_BG_MAIN, padx=16, pady=12)
        footer.pack(fill="x", side="bottom")

        ModernButton(footer, text="Cancel", style="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ModernButton(footer, text="🚀 Start Export", style="primary", command=self._confirm).pack(side="right")

    def _browse_location(self):
        chosen = filedialog.askdirectory(title="Choose Parent Folder for Export", initialdir=str(self.target_parent))
        if chosen:
            self.target_parent = Path(chosen)
            self._update_path_preview()

    def _update_path_preview(self):
        name = self.folder_name_var.get().strip() or "SmartSort_Export"
        full = self.target_parent / name
        self.lbl_path_preview.configure(text=f"Destination: {truncate_path(str(full), 56)}")

    def _confirm(self):
        name = self.folder_name_var.get().strip()
        if not name:
            messagebox.showwarning("Folder Name Required", "Please enter a valid export folder name.")
            return

        # Sanitize folder name characters
        for invalid_char in ('<', '>', ':', '"', '/', '\\', '|', '?', '*'):
            if invalid_char in name:
                messagebox.showerror("Invalid Characters", f"Folder name cannot contain '{invalid_char}'.")
                return

        self.export_path = self.target_parent / name
        chosen_filter_label = self.cb_filter.get()
        self.filter_type = self.filter_options_map.get(chosen_filter_label, "all")
        self.structure = self.structure_var.get()
        self.operation = self.operation_var.get()

        self.confirmed = True
        self.destroy()


class ConfirmationDialog(tk.Toplevel):
    """Modern confirmation modal before executing in-place reorganization."""
    def __init__(self, parent, root_dir: Path, to_move: int, already_in_place: int, conflicts: int, excluded: int, dry_run: bool):
        super().__init__(parent)
        self.title("Confirm In-Place Organization")
        self.configure(bg=COLOR_BG_MAIN)
        self.geometry("520x420")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.confirmed = False

        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        x = px + max(0, (pw - 520) // 2)
        y = py + max(0, (ph - 420) // 2)
        self.geometry(f"+{x}+{y}")

        hdr = tk.Frame(self, bg=COLOR_BG_CARD, padx=20, pady=16, highlightbackground=COLOR_BORDER, highlightthickness=1)
        hdr.pack(fill="x", side="top")

        icon_title = tk.Frame(hdr, bg=COLOR_BG_CARD)
        icon_title.pack(anchor="w")
        tk.Label(icon_title, text="⚡", font=(FONT_FAMILY, 16), bg=COLOR_BG_CARD).pack(side="left", padx=(0, 8))
        tk.Label(icon_title, text="Ready to Organize Files", font=FONT_TITLE, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(side="left")

        mode_badge_text = "DRY RUN (Preview Only)" if dry_run else "IN-PLACE REORGANIZATION"
        mode_badge_fg = COLOR_WARNING if dry_run else COLOR_PRIMARY
        tk.Label(hdr, text=mode_badge_text, font=FONT_BODY_BOLD, fg=mode_badge_fg, bg=COLOR_BG_CARD).pack(anchor="w", pady=(4, 0))

        body_card = Card(self)
        body_card.pack(fill="both", expand=True, padx=20, pady=16)

        inner = tk.Frame(body_card, bg=COLOR_BG_CARD, padx=16, pady=14)
        inner.pack(fill="both", expand=True)

        tk.Label(inner, text="Target Folder:", font=FONT_CARD_HEADER, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).pack(anchor="w")
        tk.Label(inner, text=truncate_path(str(root_dir), 50), font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w", pady=(2, 10))

        stats_frame = tk.Frame(inner, bg=COLOR_BG_CARD)
        stats_frame.pack(fill="x", pady=4)

        rows = [
            ("Files to be moved:", f"{to_move:,} files", COLOR_PRIMARY),
            ("Files already in place (skipped):", f"{already_in_place:,} files", COLOR_TEXT_MUTED),
            ("Conflicts renamed safely:", f"{conflicts:,} files", COLOR_WARNING if conflicts else COLOR_TEXT_MUTED),
            ("Files excluded by user:", f"{excluded:,} files", COLOR_TEXT_MUTED),
        ]

        for i, (label, val, col) in enumerate(rows):
            lbl = tk.Label(stats_frame, text=label, font=FONT_BODY, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD)
            lbl.grid(row=i, column=0, sticky="w", pady=3)
            val_lbl = tk.Label(stats_frame, text=val, font=FONT_BODY_BOLD, fg=col, bg=COLOR_BG_CARD)
            val_lbl.grid(row=i, column=1, sticky="e", padx=(20, 0), pady=3)

        stats_frame.columnconfigure(0, weight=1)

        warn_text = "Files will be safely organized into category subfolders.\nOriginal files remain intact, and unique collision-safe names are generated." if not dry_run else "Preview mode enabled: No files will be moved or modified on disk."
        tk.Label(inner, text=warn_text, font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD, justify="left").pack(anchor="w", pady=(12, 0))

        footer = tk.Frame(self, bg=COLOR_BG_MAIN, padx=20, pady=12)
        footer.pack(fill="x", side="bottom")

        ModernButton(footer, text="Cancel", style="secondary", command=self._cancel).pack(side="right", padx=(8, 0))
        btn_confirm_text = "Start Simulation" if dry_run else "Confirm & Organize"
        ModernButton(footer, text=btn_confirm_text, style="primary", command=self._confirm).pack(side="right")

    def _confirm(self):
        self.confirmed = True
        self.destroy()

    def _cancel(self):
        self.confirmed = False
        self.destroy()


class CompletionSummaryDialog(tk.Toplevel):
    """Summary modal after scanning, organization, or export finishes."""
    def __init__(self, parent, title: str, stats: List[Tuple[str, str]], open_folder_cmd=None, view_errors_cmd=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=COLOR_BG_MAIN)
        self.geometry("490x390")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        x = px + max(0, (pw - 490) // 2)
        y = py + max(0, (ph - 390) // 2)
        self.geometry(f"+{x}+{y}")

        hdr = tk.Frame(self, bg=COLOR_BG_CARD, padx=20, pady=14, highlightbackground=COLOR_BORDER, highlightthickness=1)
        hdr.pack(fill="x", side="top")
        tk.Label(hdr, text="✓ Operation Complete", font=FONT_TITLE, fg=COLOR_SUCCESS, bg=COLOR_BG_CARD).pack(anchor="w")

        body = Card(self)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        inner = tk.Frame(body, bg=COLOR_BG_CARD, padx=16, pady=14)
        inner.pack(fill="both", expand=True)

        for i, (label, val) in enumerate(stats):
            tk.Label(inner, text=label, font=FONT_BODY, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).grid(row=i, column=0, sticky="w", pady=4)
            tk.Label(inner, text=val, font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).grid(row=i, column=1, sticky="e", padx=(20, 0), pady=4)

        inner.columnconfigure(0, weight=1)

        footer = tk.Frame(self, bg=COLOR_BG_MAIN, padx=20, pady=12)
        footer.pack(fill="x", side="bottom")

        ModernButton(footer, text="Close", style="secondary", command=self.destroy).pack(side="right")

        if view_errors_cmd:
            ModernButton(footer, text="View Errors", style="danger", command=view_errors_cmd).pack(side="left", padx=(0, 8))

        if open_folder_cmd:
            ModernButton(footer, text="📂 Open Folder", style="outline", command=open_folder_cmd).pack(side="left")


class ImagePreviewDialog(tk.Toplevel):
    """Renders a thumbnail image preview and file metadata using Pillow."""
    def __init__(self, parent, file_item: FileItem):
        super().__init__(parent)
        self.title(f"Image Preview — {file_item.filename}")
        self.configure(bg=COLOR_BG_MAIN)
        self.geometry("540x520")
        self.resizable(False, False)
        self.transient(parent)

        hdr = tk.Frame(self, bg=COLOR_BG_CARD, padx=16, pady=12, highlightbackground=COLOR_BORDER, highlightthickness=1)
        hdr.pack(fill="x")
        tk.Label(hdr, text=file_item.filename, font=FONT_SECTION, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")
        tk.Label(hdr, text=f"Category: {file_item.category} • Size: {format_size(file_item.size_bytes)}", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).pack(anchor="w")

        preview_card = Card(self)
        preview_card.pack(fill="both", expand=True, padx=16, pady=12)

        self.img_label = tk.Label(preview_card, bg=COLOR_BG_CARD)
        self.img_label.pack(fill="both", expand=True, padx=12, pady=12)

        info_frame = tk.Frame(preview_card, bg=COLOR_BG_CARD)
        info_frame.pack(fill="x", padx=12, pady=(0, 10))

        if HAS_PILLOW:
            try:
                img = Image.open(file_item.original_path)
                dim_str = f"{img.width} × {img.height} px"
                mode_str = img.mode
                fmt_str = img.format or file_item.extension.upper()

                img.thumbnail((460, 320), Image.Resampling.LANCZOS)
                self.photo = ImageTk.PhotoImage(img)
                self.img_label.configure(image=self.photo)

                meta_text = f"Dimensions: {dim_str}  |  Format: {fmt_str}  |  Color Mode: {mode_str}"
                tk.Label(info_frame, text=meta_text, font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).pack(anchor="center")
            except Exception as e:
                self.img_label.configure(text=f"Could not load image preview:\n{e}", fg=COLOR_DANGER, font=FONT_BODY)
        else:
            self.img_label.configure(text="Pillow library is not installed.\nImage preview unavailable.", fg=COLOR_TEXT_MUTED, font=FONT_BODY)

        btn_box = tk.Frame(self, bg=COLOR_BG_MAIN, padx=16, pady=10)
        btn_box.pack(fill="x", side="bottom")
        ModernButton(btn_box, text="Close", style="secondary", command=self.destroy).pack(side="right")


class ErrorLogDialog(tk.Toplevel):
    """Displays error and permission failure logs."""
    def __init__(self, parent, errors: List[Tuple[str, str]]):
        super().__init__(parent)
        self.title("Operation Errors")
        self.geometry("640x440")
        self.configure(bg=COLOR_BG_MAIN)
        self.transient(parent)

        hdr = tk.Frame(self, bg=COLOR_BG_CARD, padx=16, pady=12, highlightbackground=COLOR_BORDER, highlightthickness=1)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Encountered {len(errors)} Issue(s)", font=FONT_SECTION, fg=COLOR_DANGER, bg=COLOR_BG_CARD).pack(anchor="w")

        card = Card(self)
        card.pack(fill="both", expand=True, padx=16, pady=12)

        txt_frame = tk.Frame(card, bg=COLOR_BG_CARD)
        txt_frame.pack(fill="both", expand=True, padx=8, pady=8)

        text_widget = tk.Text(txt_frame, font=FONT_MONO, bg="#F8FAFC", fg=COLOR_TEXT_MAIN, relief="flat", bd=0)
        scrollbar = ttk.Scrollbar(txt_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        text_widget.pack(side="left", fill="both", expand=True)

        for filename, err in errors:
            text_widget.insert("end", f"File: {filename}\nError: {err}\n{'-'*60}\n")
        text_widget.configure(state="disabled")

        btn_box = tk.Frame(self, bg=COLOR_BG_MAIN, padx=16, pady=10)
        btn_box.pack(fill="x", side="bottom")
        ModernButton(btn_box, text="Close", style="secondary", command=self.destroy).pack(side="right")


# ============================================================================
# MAIN APPLICATION
# ============================================================================
class SmartSortApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SmartSort — Intelligent File Organizer")
        self.root.geometry("1240x870")
        self.root.minsize(1020, 680)
        self.root.configure(bg=COLOR_BG_MAIN)

        # Application state
        self.selected_folder: Optional[Path] = None
        self.last_export_dir: Optional[Path] = None
        self.all_files: List[FileItem] = []
        self.filtered_files: List[FileItem] = []
        self.undo_stack: List[List[UndoRecord]] = []
        self.last_errors: List[Tuple[str, str]] = []

        self.msg_queue: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.current_worker: Optional[threading.Thread] = None

        # Pagination & Filtering state
        self.page_size = 100
        self.current_page = 1
        self.sort_column = "filename"
        self.sort_ascending = True

        # Setup TTK Styling
        self._setup_styles()

        # Build UI Components
        self._build_ui()

        # Periodic queue checker
        self.root.after(35, self._poll_queue)

        # Window closing protocol
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)

    def _setup_styles(self):
        """Configure modern clean TTK light-theme styles."""
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        # Treeview styling
        self.style.configure(
            "Treeview",
            background=COLOR_BG_CARD,
            foreground=COLOR_TEXT_MAIN,
            fieldbackground=COLOR_BG_CARD,
            rowheight=26,
            bordercolor=COLOR_BORDER,
            font=FONT_BODY
        )
        self.style.configure(
            "Treeview.Heading",
            background="#F8FAFC",
            foreground="#334155",
            relief="flat",
            font=FONT_BODY_BOLD,
            bordercolor=COLOR_BORDER
        )
        self.style.map("Treeview.Heading", background=[("active", COLOR_BORDER)])
        self.style.map(
            "Treeview",
            background=[("selected", COLOR_PRIMARY_TINT)],
            foreground=[("selected", COLOR_PRIMARY_HOVER)]
        )

        # Scrollbar
        self.style.configure(
            "Vertical.TScrollbar",
            troughcolor=COLOR_BG_MAIN,
            background="#CBD5E1",
            bordercolor=COLOR_BG_MAIN,
            arrowcolor=COLOR_TEXT_MUTED
        )

    def _build_ui(self):
        """Construct the entire application layout."""
        # 1. STATUS BAR (Docked at the very bottom of root)
        self._build_status_bar(self.root)

        # 2. BOTTOM ACTIONS BAR (Docked above the status bar, ALWAYS visible)
        self._build_bottom_actions(self.root)

        # 3. MAIN EXPANDABLE CONTAINER (Fills all space between header and bottom bar)
        main_container = tk.Frame(self.root, bg=COLOR_BG_MAIN)
        main_container.pack(fill="both", expand=True, padx=16, pady=(10, 4))

        # A. TOP HEADER BAR
        self._build_header(main_container)

        # B. FOLDER SELECTION CARD
        self._build_folder_card(main_container)

        # C. SUMMARY KPI CARDS
        self._build_summary_cards(main_container)

        # D. PROGRESS CARD (Smooth 60 FPS Canvas)
        self._build_progress_card(main_container)

        # E & F. SPLIT VIEW: CATEGORY OVERVIEW (LEFT) + FILE PREVIEW TABLE (RIGHT)
        self._build_preview_panes(main_container)

    def _build_header(self, parent: tk.Frame):
        """Header bar with app name, tagline, and live pulsing status badge."""
        header_frame = tk.Frame(parent, bg=COLOR_BG_MAIN)
        header_frame.pack(fill="x", pady=(0, 10))

        left_box = tk.Frame(header_frame, bg=COLOR_BG_MAIN)
        left_box.pack(side="left")

        logo_title = tk.Frame(left_box, bg=COLOR_BG_MAIN)
        logo_title.pack(anchor="w")

        tk.Label(logo_title, text="🗂️", font=(FONT_FAMILY, 17), bg=COLOR_BG_MAIN).pack(side="left", padx=(0, 6))
        tk.Label(logo_title, text="SmartSort", font=FONT_TITLE, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_MAIN).pack(side="left")

        tk.Label(
            left_box,
            text="Your files, organized intelligently",
            font=FONT_SUBTITLE,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_MAIN
        ).pack(anchor="w", padx=(30, 0))

        # Live animated status pill
        self.status_badge = AnimatedBadge(header_frame, text="Ready", status_key="Ready")
        self.status_badge.pack(side="right", pady=4)

    def _build_folder_card(self, parent: tk.Frame):
        """Folder selection card with path input, info stats, and options."""
        card = Card(parent)
        card.pack(fill="x", pady=(0, 10))

        inner = tk.Frame(card, bg=COLOR_BG_CARD, padx=14, pady=10)
        inner.pack(fill="x")

        # Row 1: Path Entry + Choose Folder Button
        row1 = tk.Frame(inner, bg=COLOR_BG_CARD)
        row1.pack(fill="x")

        tk.Label(row1, text="Selected Folder:", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(side="left", padx=(0, 8))

        self.path_var = tk.StringVar(value="")
        self.entry_path = tk.Entry(
            row1,
            textvariable=self.path_var,
            font=FONT_BODY,
            bg="#F8FAFC",
            fg=COLOR_TEXT_MAIN,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            bd=0
        )
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=4)

        self.btn_browse = ModernButton(
            row1,
            text="📁 Choose Folder",
            style="secondary",
            command=self._on_choose_folder
        )
        self.btn_browse.pack(side="left")

        # Row 2: Stats Info + Options + Action Buttons
        row2 = tk.Frame(inner, bg=COLOR_BG_CARD)
        row2.pack(fill="x", pady=(8, 0))

        self.lbl_folder_info = tk.Label(
            row2,
            text="No folder selected • Choose a folder to begin analysis",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_CARD
        )
        self.lbl_folder_info.pack(side="left")

        right_actions = tk.Frame(row2, bg=COLOR_BG_CARD)
        right_actions.pack(side="right")

        self.var_include_subfolders = tk.BooleanVar(value=True)
        chk_sub = tk.Checkbutton(
            right_actions,
            text="Include subfolders",
            variable=self.var_include_subfolders,
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD,
            fg=COLOR_TEXT_MAIN
        )
        chk_sub.pack(side="left", padx=(0, 12))

        self.var_dry_run = tk.BooleanVar(value=False)
        chk_dry = tk.Checkbutton(
            right_actions,
            text="Dry run (Preview only)",
            variable=self.var_dry_run,
            font=FONT_BODY,
            bg=COLOR_BG_CARD,
            activebackground=COLOR_BG_CARD,
            fg=COLOR_TEXT_MAIN
        )
        chk_dry.pack(side="left", padx=(0, 16))

        self.btn_analyze = ModernButton(
            right_actions,
            text="🔍 Analyze Files",
            style="primary",
            command=self._start_scan
        )
        self.btn_analyze.pack(side="left", padx=(0, 6))

        self.btn_cancel_scan = ModernButton(
            right_actions,
            text="⏹ Cancel",
            style="danger",
            command=self._cancel_current_operation
        )
        self.btn_cancel_scan.pack(side="left", padx=(0, 10))
        self.btn_cancel_scan.set_enabled(False)

        # Prominent Top Export & In-Place Organize Buttons
        self.btn_export_top = ModernButton(
            right_actions,
            text="📦 Export to Folder...",
            style="export",
            command=self._start_export_dialog
        )
        self.btn_export_top.pack(side="left", padx=(0, 6))
        self.btn_export_top.set_enabled(False)

        self.btn_organize_top = ModernButton(
            right_actions,
            text="🚀 Organize In-Place",
            style="secondary",
            command=self._start_organize
        )
        self.btn_organize_top.pack(side="left")
        self.btn_organize_top.set_enabled(False)

    def _build_summary_cards(self, parent: tk.Frame):
        """Displays 7 KPI summary cards."""
        self.summary_frame = tk.Frame(parent, bg=COLOR_BG_MAIN)
        self.summary_frame.pack(fill="x", pady=(0, 10))

        self.card_total_files = StatCard(self.summary_frame, title="Total Files", value="0", subtitle="Discovered", accent="#0D9488")
        self.card_total_files.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_total_folders = StatCard(self.summary_frame, title="Total Folders", value="0", subtitle="Directories", accent="#3B82F6")
        self.card_total_folders.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_total_size = StatCard(self.summary_frame, title="Total Size", value="0 B", subtitle="Disk Usage", accent="#6366F1")
        self.card_total_size.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_images = StatCard(self.summary_frame, title="Images", value="0", subtitle="0 B", accent="#EC4899")
        self.card_images.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_videos = StatCard(self.summary_frame, title="Videos", value="0", subtitle="0 B", accent="#8B5CF6")
        self.card_videos.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_docs = StatCard(self.summary_frame, title="Documents", value="0", subtitle="0 B", accent="#F59E0B")
        self.card_docs.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.card_other = StatCard(self.summary_frame, title="Other", value="0", subtitle="0 B", accent="#64748B")
        self.card_other.pack(side="left", fill="both", expand=True)

    def _build_progress_card(self, parent: tk.Frame):
        """Modern progress card with smooth 60 FPS progress bar, ETA, elapsed, and speed."""
        self.progress_card = Card(parent)
        self.progress_card.pack(fill="x", pady=(0, 10))

        inner = tk.Frame(self.progress_card, bg=COLOR_BG_CARD, padx=14, pady=10)
        inner.pack(fill="x")

        # Top row: Operation description & Large Percentage
        top_row = tk.Frame(inner, bg=COLOR_BG_CARD)
        top_row.pack(fill="x")

        self.lbl_operation = tk.Label(
            top_row,
            text="Ready • Select folder and click Analyze",
            font=FONT_BODY_BOLD,
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG_CARD
        )
        self.lbl_operation.pack(side="left")

        self.lbl_percentage = tk.Label(
            top_row,
            text="0%",
            font=FONT_STAT_BIG,
            fg=COLOR_PRIMARY,
            bg=COLOR_BG_CARD
        )
        self.lbl_percentage.pack(side="right")

        # Middle row: Smooth Canvas Progress Bar
        self.progressbar = SmoothProgressBar(
            inner,
            height=16,
            bg=COLOR_BG_CARD,
            trough_color=COLOR_BORDER,
            fill_color=COLOR_PRIMARY
        )
        self.progressbar.pack(fill="x", pady=(6, 8))

        # Bottom row: Current file, processed counter, elapsed, ETA, speed
        bot_row = tk.Frame(inner, bg=COLOR_BG_CARD)
        bot_row.pack(fill="x")

        self.lbl_current_file = tk.Label(
            bot_row,
            text="Current: None",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg=COLOR_BG_CARD
        )
        self.lbl_current_file.pack(side="left")

        stats_box = tk.Frame(bot_row, bg=COLOR_BG_CARD)
        stats_box.pack(side="right")

        self.lbl_processed_count = tk.Label(stats_box, text="Processed: 0 / 0", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD)
        self.lbl_processed_count.pack(side="left", padx=(0, 14))

        self.lbl_elapsed_time = tk.Label(stats_box, text="Elapsed: 0 sec", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD)
        self.lbl_elapsed_time.pack(side="left", padx=(0, 14))

        self.lbl_eta_time = tk.Label(stats_box, text="ETA: --", font=FONT_SMALL, fg=COLOR_PRIMARY, bg=COLOR_BG_CARD)
        self.lbl_eta_time.pack(side="left", padx=(0, 14))

        self.lbl_speed = tk.Label(stats_box, text="Speed: -- files/sec", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD)
        self.lbl_speed.pack(side="left")

    def _build_preview_panes(self, parent: tk.Frame):
        """Split view: Left Category Overview table, Right File Preview table."""
        split_frame = tk.Frame(parent, bg=COLOR_BG_MAIN)
        split_frame.pack(fill="both", expand=True, pady=(0, 10))

        # LEFT PANE: Category Overview Card (Width ~320px)
        left_card = Card(split_frame, width=320)
        left_card.pack(side="left", fill="y", padx=(0, 8))
        left_card.pack_propagate(False)

        cat_header = tk.Frame(left_card, bg=COLOR_BG_CARD, padx=12, pady=10)
        cat_header.pack(fill="x")
        tk.Label(cat_header, text="Category Overview", font=FONT_SECTION, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(anchor="w")
        tk.Label(cat_header, text="Click a category to filter the preview", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).pack(anchor="w")

        # Category Treeview
        cat_tree_frame = tk.Frame(left_card, bg=COLOR_BG_CARD, padx=8)
        cat_tree_frame.pack(fill="both", expand=True)

        self.cat_tree = ttk.Treeview(cat_tree_frame, columns=("name", "count", "size"), show="headings", selectmode="browse")
        self.cat_tree.heading("name", text="Category")
        self.cat_tree.heading("count", text="Files")
        self.cat_tree.heading("size", text="Size")
        self.cat_tree.column("name", width=140, anchor="w")
        self.cat_tree.column("count", width=55, anchor="e")
        self.cat_tree.column("size", width=75, anchor="e")

        cat_scroll = ttk.Scrollbar(cat_tree_frame, orient="vertical", command=self.cat_tree.yview)
        self.cat_tree.configure(yscrollcommand=cat_scroll.set)

        self.cat_tree.pack(side="left", fill="both", expand=True)
        cat_scroll.pack(side="right", fill="y")
        self.cat_tree.bind("<<TreeviewSelect>>", self._on_category_selected)

        # Example extensions box
        self.ext_info_frame = tk.Frame(left_card, bg="#F8FAFC", padx=10, pady=8, highlightbackground=COLOR_BORDER, highlightthickness=1)
        self.ext_info_frame.pack(fill="x", padx=8, pady=(8, 4))
        self.lbl_ext_info = tk.Label(
            self.ext_info_frame,
            text="Extensions: Select a category above",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg="#F8FAFC",
            wraplength=280,
            justify="left"
        )
        self.lbl_ext_info.pack(anchor="w")

        # Category Action Buttons
        cat_btn_box = tk.Frame(left_card, bg=COLOR_BG_CARD, padx=8, pady=6)
        cat_btn_box.pack(fill="x")
        ModernButton(cat_btn_box, text="Include All", style="secondary", command=lambda: self._set_category_inclusion(True)).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ModernButton(cat_btn_box, text="Exclude All", style="secondary", command=lambda: self._set_category_inclusion(False)).pack(side="left", fill="x", expand=True)

        # RIGHT PANE: File Preview Table Card
        right_card = Card(split_frame)
        right_card.pack(side="left", fill="both", expand=True)

        right_inner = tk.Frame(right_card, bg=COLOR_BG_CARD, padx=12, pady=10)
        right_inner.pack(fill="both", expand=True)

        # Top Filter Bar
        filter_bar = tk.Frame(right_inner, bg=COLOR_BG_CARD)
        filter_bar.pack(fill="x", pady=(0, 8))

        # Search Box
        tk.Label(filter_bar, text="🔍", font=FONT_BODY, bg=COLOR_BG_CARD).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._apply_filters())
        self.search_entry = tk.Entry(
            filter_bar,
            textvariable=self.search_var,
            font=FONT_BODY,
            bg="#F8FAFC",
            fg=COLOR_TEXT_MAIN,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            bd=0,
            width=22
        )
        self.search_entry.pack(side="left", padx=(0, 12), ipady=3)

        # Category Dropdown Filter
        tk.Label(filter_bar, text="Category:", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(side="left", padx=(0, 4))
        self.cat_filter_var = tk.StringVar(value="All Categories")
        cat_options = ["All Categories"] + list(CATEGORIES.keys())
        self.cb_cat_filter = ttk.Combobox(
            filter_bar,
            textvariable=self.cat_filter_var,
            values=cat_options,
            state="readonly",
            width=14
        )
        self.cb_cat_filter.pack(side="left", padx=(0, 12))
        self.cb_cat_filter.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # Status Dropdown Filter
        tk.Label(filter_bar, text="Status:", font=FONT_BODY_BOLD, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD).pack(side="left", padx=(0, 4))
        self.status_filter_var = tk.StringVar(value="All Statuses")
        status_options = ["All Statuses", "Ready to Move", "Already in Place", "Conflict - Renamed", "Excluded", "Failed", "Moved", "Copied"]
        self.cb_status_filter = ttk.Combobox(
            filter_bar,
            textvariable=self.status_filter_var,
            values=status_options,
            state="readonly",
            width=16
        )
        self.cb_status_filter.pack(side="left", padx=(0, 12))
        self.cb_status_filter.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # Reset Filter Button
        ModernButton(filter_bar, text="Clear Filters", style="secondary", command=self._clear_filters).pack(side="left")

        # Table Selection Toolbar
        toolbar = tk.Frame(right_inner, bg=COLOR_BG_CARD)
        toolbar.pack(fill="x", pady=(0, 6))

        ModernButton(toolbar, text="Select All", style="secondary", padx=8, pady=3, command=lambda: self._toggle_selection_all(True)).pack(side="left", padx=(0, 4))
        ModernButton(toolbar, text="Deselect All", style="secondary", padx=8, pady=3, command=lambda: self._toggle_selection_all(False)).pack(side="left", padx=(0, 4))
        ModernButton(toolbar, text="Invert Selection", style="secondary", padx=8, pady=3, command=self._invert_selection).pack(side="left", padx=(0, 8))

        if HAS_PILLOW:
            ModernButton(toolbar, text="👁 Preview Thumbnail", style="outline", padx=8, pady=3, command=self._preview_selected_image).pack(side="left", padx=(0, 6))

        # Table Toolbar Export Button
        self.btn_export_table = ModernButton(
            toolbar,
            text="📦 Export Files...",
            style="export",
            padx=10,
            pady=3,
            command=self._start_export_dialog
        )
        self.btn_export_table.pack(side="left", padx=(0, 8))
        self.btn_export_table.set_enabled(False)

        self.lbl_table_count = tk.Label(toolbar, text="Showing 0 of 0 files", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD)
        self.lbl_table_count.pack(side="right")

        # File Preview Treeview Table
        table_container = tk.Frame(right_inner, bg=COLOR_BG_CARD)
        table_container.pack(fill="both", expand=True)

        columns = ("include", "filename", "category", "extension", "size", "destination", "status")
        self.file_tree = ttk.Treeview(
            table_container,
            columns=columns,
            show="headings",
            selectmode="extended"
        )

        # Headings with sort callbacks
        self.file_tree.heading("include", text="✓", command=lambda: self._sort_by_column("include"))
        self.file_tree.heading("filename", text="Filename", command=lambda: self._sort_by_column("filename"))
        self.file_tree.heading("category", text="Category", command=lambda: self._sort_by_column("category"))
        self.file_tree.heading("extension", text="Ext", command=lambda: self._sort_by_column("extension"))
        self.file_tree.heading("size", text="Size", command=lambda: self._sort_by_column("size_bytes"))
        self.file_tree.heading("destination", text="Proposed Destination", command=lambda: self._sort_by_column("destination"))
        self.file_tree.heading("status", text="Status", command=lambda: self._sort_by_column("status"))

        # Column widths
        self.file_tree.column("include", width=36, anchor="center")
        self.file_tree.column("filename", width=170, anchor="w")
        self.file_tree.column("category", width=95, anchor="w")
        self.file_tree.column("extension", width=55, anchor="center")
        self.file_tree.column("size", width=70, anchor="e")
        self.file_tree.column("destination", width=220, anchor="w")
        self.file_tree.column("status", width=125, anchor="w")

        # Scrollbars
        scroll_y = ttk.Scrollbar(table_container, orient="vertical", command=self.file_tree.yview)
        scroll_x = ttk.Scrollbar(table_container, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.file_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)

        # Tags for status row styling
        self.file_tree.tag_configure("ready", foreground=COLOR_SUCCESS)
        self.file_tree.tag_configure("conflict", foreground=COLOR_WARNING)
        self.file_tree.tag_configure("inplace", foreground=COLOR_TEXT_MUTED)
        self.file_tree.tag_configure("excluded", foreground=COLOR_TEXT_LIGHT)
        self.file_tree.tag_configure("failed", foreground=COLOR_DANGER)
        self.file_tree.tag_configure("moved", foreground=COLOR_SUCCESS)
        self.file_tree.tag_configure("copied", foreground=COLOR_SUCCESS)
        self.file_tree.tag_configure("altrow", background=COLOR_ROW_ALT)

        # Bind events
        self.file_tree.bind("<Double-Button-1>", self._on_row_double_click)
        self.file_tree.bind("<space>", self._on_space_pressed)
        self.file_tree.bind("<Button-3>", self._show_context_menu)

        # Pagination Controls
        page_bar = tk.Frame(right_inner, bg=COLOR_BG_CARD)
        page_bar.pack(fill="x", pady=(6, 0))

        self.btn_prev_page = ModernButton(page_bar, text="◀ Prev", style="secondary", padx=8, pady=2, command=self._prev_page)
        self.btn_prev_page.pack(side="left")

        self.lbl_page_info = tk.Label(page_bar, text="Page 1 of 1", font=FONT_SMALL, fg=COLOR_TEXT_MAIN, bg=COLOR_BG_CARD)
        self.lbl_page_info.pack(side="left", padx=10)

        self.btn_next_page = ModernButton(page_bar, text="Next ▶", style="secondary", padx=8, pady=2, command=self._next_page)
        self.btn_next_page.pack(side="left")

        tk.Label(page_bar, text="Rows per page:", font=FONT_SMALL, fg=COLOR_TEXT_MUTED, bg=COLOR_BG_CARD).pack(side="right", padx=(0, 4))
        self.page_size_var = tk.StringVar(value="100")
        cb_page_size = ttk.Combobox(page_bar, textvariable=self.page_size_var, values=["50", "100", "250", "500", "All"], state="readonly", width=6)
        cb_page_size.pack(side="right")
        cb_page_size.bind("<<ComboboxSelected>>", self._on_page_size_changed)

    def _build_bottom_actions(self, parent: tk.Widget):
        """Bottom action buttons with In-Place Organize and Dedicated Export."""
        card = Card(parent)
        card.pack(fill="x", side="bottom", padx=16, pady=(0, 6))

        inner = tk.Frame(card, bg=COLOR_BG_CARD, padx=14, pady=10)
        inner.pack(fill="x")

        # Left action group
        left_grp = tk.Frame(inner, bg=COLOR_BG_CARD)
        left_grp.pack(side="left")

        self.btn_open_folder = ModernButton(
            left_grp,
            text="📁 Open Source Folder",
            style="secondary",
            command=self._open_selected_folder
        )
        self.btn_open_folder.pack(side="left", padx=(0, 6))

        self.btn_open_dest = ModernButton(
            left_grp,
            text="📂 Open Export Folder",
            style="secondary",
            command=self._open_export_folder
        )
        self.btn_open_dest.pack(side="left", padx=(0, 6))

        self.btn_reset = ModernButton(
            left_grp,
            text="🔄 Reset",
            style="secondary",
            command=self._reset_application
        )
        self.btn_reset.pack(side="left")

        # Right action group
        right_grp = tk.Frame(inner, bg=COLOR_BG_CARD)
        right_grp.pack(side="right")

        self.btn_undo = ModernButton(
            right_grp,
            text="↩ Undo",
            style="secondary",
            command=self._on_undo_clicked
        )
        self.btn_undo.pack(side="left", padx=(0, 8))
        self.btn_undo.set_enabled(False)

        # Prominent "Export to Dedicated Folder" Button
        self.btn_export = ModernButton(
            right_grp,
            text="📦 Export to Folder...",
            style="export",
            padx=14,
            pady=7,
            command=self._start_export_dialog
        )
        self.btn_export.pack(side="left", padx=(0, 8))
        self.btn_export.set_enabled(False)

        # In-Place Organization Button
        self.btn_organize = ModernButton(
            right_grp,
            text="🚀 Organize In-Place",
            style="primary",
            padx=16,
            pady=7,
            command=self._start_organize
        )
        self.btn_organize.pack(side="left")
        self.btn_organize.set_enabled(False)

    def _build_status_bar(self, parent: tk.Widget):
        """Bottom application status bar."""
        self.status_bar_frame = tk.Frame(parent, bg="#EDF2F7", highlightbackground=COLOR_BORDER, highlightthickness=1)
        self.status_bar_frame.pack(fill="x", side="bottom")

        self.lbl_status_bar = tk.Label(
            self.status_bar_frame,
            text="Ready • Select a folder to begin organizing",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg="#EDF2F7",
            padx=12,
            pady=3
        )
        self.lbl_status_bar.pack(side="left")

        self.lbl_app_info = tk.Label(
            self.status_bar_frame,
            text="SmartSort v1.0 • Windows File Organizer",
            font=FONT_SMALL,
            fg=COLOR_TEXT_MUTED,
            bg="#EDF2F7",
            padx=12,
            pady=3
        )
        self.lbl_app_info.pack(side="right")

    # ========================================================================
    # EVENT HANDLERS & ACTIONS
    # ========================================================================
    def _on_choose_folder(self):
        """Open native Windows directory selection dialog."""
        initial_dir = self.selected_folder if self.selected_folder else Path.home()
        chosen = filedialog.askdirectory(
            title="Select Folder to Organize",
            initialdir=str(initial_dir)
        )
        if chosen:
            p = Path(chosen)
            if is_system_path(p):
                confirm = messagebox.askyesno(
                    "Warning: System Location",
                    f"'{p}' appears to be a Windows system or protected directory.\n\nOrganizing system files can cause instability. Are you sure you want to proceed?"
                )
                if not confirm:
                    return

            self.selected_folder = p
            self.path_var.set(str(p))
            self.lbl_folder_info.configure(text=f"Folder selected: {p} • Click 'Analyze Files' to scan")
            self.lbl_status_bar.configure(text=f"Selected folder: {p}")
            self.btn_analyze.set_enabled(True)
            self._update_status_indicator("Ready", "Ready")

    def _set_export_buttons_enabled(self, enabled: bool):
        """Update enabled state across all Export buttons (top card, table bar, bottom bar)."""
        for btn in [getattr(self, "btn_export", None),
                    getattr(self, "btn_export_top", None),
                    getattr(self, "btn_export_table", None)]:
            if btn:
                btn.set_enabled(enabled)

    def _set_organize_buttons_enabled(self, enabled: bool):
        """Update enabled state across all Organize buttons (top card, bottom bar)."""
        for btn in [getattr(self, "btn_organize", None),
                    getattr(self, "btn_organize_top", None)]:
            if btn:
                btn.set_enabled(enabled)

    def _start_scan(self):
        """Initiate background thread to scan and categorize files."""
        if not self.selected_folder or not self.selected_folder.is_dir():
            messagebox.showwarning("No Folder Selected", "Please choose a valid folder before scanning.")
            return

        # Reset state
        self.cancel_event.clear()
        self.all_files.clear()
        self.filtered_files.clear()
        self._clear_tree()

        # UI state during scan
        self.btn_analyze.set_enabled(False)
        self.btn_browse.set_enabled(False)
        self.btn_cancel_scan.set_enabled(True)
        self._set_organize_buttons_enabled(False)
        self._set_export_buttons_enabled(False)

        self.progressbar.start_indeterminate()
        self.lbl_operation.configure(text="Discovering files and folders...")
        self.lbl_percentage.configure(text="--")
        self._update_status_indicator("Scanning", "Scanning")

        # Launch worker thread
        self.current_worker = ScanWorker(
            root_dir=self.selected_folder,
            include_subfolders=self.var_include_subfolders.get(),
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event
        )
        self.current_worker.start()

    def _cancel_current_operation(self):
        """Signal background worker to cancel cleanly."""
        self.cancel_event.set()
        self.lbl_operation.configure(text="Cancelling operation...")
        self.btn_cancel_scan.set_enabled(False)

    def _start_organize(self):
        """Open confirmation modal and run in-place organization worker thread."""
        if not self.selected_folder or not self.all_files:
            return

        to_move = sum(1 for f in self.all_files if f.is_included and f.status in ("Ready to Move", "Conflict - Renamed", "Moved (Simulated)"))
        already_in_place = sum(1 for f in self.all_files if f.status == "Already in Place")
        conflicts = sum(1 for f in self.all_files if f.is_included and f.status == "Conflict - Renamed")
        excluded = sum(1 for f in self.all_files if not f.is_included and f.status != "Already in Place")

        if to_move == 0:
            messagebox.showinfo("No Files to Move", "There are no files selected or eligible to be moved in-place.")
            return

        dlg = ConfirmationDialog(
            parent=self.root,
            root_dir=self.selected_folder,
            to_move=to_move,
            already_in_place=already_in_place,
            conflicts=conflicts,
            excluded=excluded,
            dry_run=self.var_dry_run.get()
        )
        self.root.wait_window(dlg)

        if not dlg.confirmed:
            return

        self.cancel_event.clear()
        self.btn_analyze.set_enabled(False)
        self.btn_browse.set_enabled(False)
        self._set_organize_buttons_enabled(False)
        self._set_export_buttons_enabled(False)
        self.btn_cancel_scan.set_enabled(True)

        self.progressbar.set_value(0)
        mode_str = "Simulating move of" if self.var_dry_run.get() else "Moving"
        self.lbl_operation.configure(text=f"{mode_str} file 1 of {to_move}...")
        self._update_status_indicator("Organizing", "Organizing")

        self.current_worker = OrganizeWorker(
            root_dir=self.selected_folder,
            items=self.all_files,
            dry_run=self.var_dry_run.get(),
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event
        )
        self.current_worker.start()

    def _start_export_dialog(self):
        """Open the Export Dialog to configure folder name, separation type, and transfer mode."""
        if not self.selected_folder or not self.all_files:
            messagebox.showinfo("Analyze Required", "Please choose and analyze a folder before exporting.")
            return

        dlg = ExportDialog(
            parent=self.root,
            default_parent_dir=self.selected_folder,
            total_files=len(self.all_files)
        )
        self.root.wait_window(dlg)

        if not dlg.confirmed:
            return

        # Prepare UI for Exporting
        self.cancel_event.clear()
        self.btn_analyze.set_enabled(False)
        self.btn_browse.set_enabled(False)
        self._set_organize_buttons_enabled(False)
        self._set_export_buttons_enabled(False)
        self.btn_cancel_scan.set_enabled(True)

        self.progressbar.set_value(0)
        self.lbl_operation.configure(text=f"Exporting files to '{dlg.export_path.name}'...")
        self._update_status_indicator("Exporting", "Exporting")

        self.current_worker = ExportWorker(
            export_dir=dlg.export_path,
            items=self.all_files,
            structure=dlg.structure,
            operation=dlg.operation,
            filter_type=dlg.filter_type,
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event
        )
        self.current_worker.start()

    def _on_undo_clicked(self):
        """Prompt and execute undo of last organization or export."""
        if not self.undo_stack:
            messagebox.showinfo("Undo", "No previous organization operations recorded in this session.")
            return

        last_batch = self.undo_stack[-1]
        confirm = messagebox.askyesno(
            "Confirm Undo",
            f"Undo will restore/revert {len(last_batch)} files.\n\nDo you want to proceed?"
        )
        if not confirm:
            return

        self.btn_undo.set_enabled(False)
        self.lbl_operation.configure(text="Restoring files...")
        self.progressbar.start_indeterminate()
        self._update_status_indicator("Organizing", "Organizing")

        worker = UndoWorker(
            records=last_batch,
            root_dir=self.selected_folder,
            msg_queue=self.msg_queue
        )
        worker.start()

    # ========================================================================
    # QUEUE MESSAGE DISPATCHER (Main UI Thread)
    # ========================================================================
    def _poll_queue(self):
        """Safely poll messages sent from background worker threads."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                msg_type = msg.get("type")

                if msg_type == "scan_discovery":
                    cnt = msg["discovered_files"]
                    fld = msg["folders"]
                    el = msg["elapsed"]
                    self.lbl_operation.configure(text=f"Discovered {cnt:,} files across {fld:,} folders...")
                    self.lbl_elapsed_time.configure(text=f"Elapsed: {format_time(el)}")

                elif msg_type == "scan_progress":
                    proc = msg["processed"]
                    tot = msg["total"]
                    cur = msg["current_file"]
                    el = msg["elapsed"]
                    speed = msg["speed"]
                    eta = msg["eta"]

                    pct = int((proc / tot) * 100) if tot > 0 else 0
                    self.progressbar.set_value(pct)
                    self.lbl_percentage.configure(text=f"{pct}%")
                    self.lbl_operation.configure(text=f"Analyzing file {proc:,} of {tot:,}...")
                    self.lbl_current_file.configure(text=f"Current: {truncate_path(cur, 35)}")
                    self.lbl_processed_count.configure(text=f"Processed: {proc:,} / {tot:,}")
                    self.lbl_elapsed_time.configure(text=f"Elapsed: {format_time(el)}")
                    self.lbl_eta_time.configure(text=f"ETA: {format_time(eta)}" if el > 1.2 and proc > 5 else "ETA: Calculating...")
                    self.lbl_speed.configure(text=f"Speed: {speed:.1f} files/sec" if el > 1.0 else "Speed: Calculating...")

                elif msg_type == "scan_complete":
                    self.progressbar.set_value(100)
                    self.lbl_percentage.configure(text="100%")
                    self.all_files = msg["items"]
                    tot_sz = msg["total_size"]
                    fld_cnt = msg["folders"]
                    inacc = msg["inaccessible"]
                    el = msg["elapsed"]
                    cat_counts = msg["cat_counts"]
                    cat_sizes = msg["cat_sizes"]

                    # Update KPI cards
                    self.card_total_files.update_card(f"{len(self.all_files):,}", "Discovered")
                    self.card_total_folders.update_card(f"{fld_cnt:,}", "Directories")
                    self.card_total_size.update_card(format_size(tot_sz), "Disk Usage")
                    self.card_images.update_card(f"{cat_counts.get('Images', 0):,}", format_size(cat_sizes.get("Images", 0)))
                    self.card_videos.update_card(f"{cat_counts.get('Videos', 0):,}", format_size(cat_sizes.get("Videos", 0)))
                    self.card_docs.update_card(f"{cat_counts.get('Documents', 0):,}", format_size(cat_sizes.get("Documents", 0)))
                    other_cnt = cat_counts.get("Other", 0) + cat_counts.get("Spreadsheets", 0) + cat_counts.get("Presentations", 0) + cat_counts.get("Archives", 0) + cat_counts.get("Code", 0)
                    other_sz = cat_sizes.get("Other", 0) + cat_sizes.get("Spreadsheets", 0) + cat_sizes.get("Presentations", 0) + cat_sizes.get("Archives", 0) + cat_sizes.get("Code", 0)
                    self.card_other.update_card(f"{other_cnt:,}", format_size(other_sz))

                    # Update Category Overview Tree
                    self._populate_category_tree(cat_counts, cat_sizes)

                    # Update Table
                    self._apply_filters()

                    # Summary status
                    self.lbl_operation.configure(text=f"Scan complete: {len(self.all_files):,} files ready for review")
                    self.lbl_folder_info.configure(text=f"Files: {len(self.all_files):,}  |  Folders: {fld_cnt:,}  |  Total Size: {format_size(tot_sz)}")
                    self.lbl_status_bar.configure(text=f"Scan complete in {format_time(el)}. {len(self.all_files):,} files analyzed.")
                    self._update_status_indicator("Ready", "Ready")

                    # Button states
                    self.btn_analyze.set_enabled(True)
                    self.btn_browse.set_enabled(True)
                    self.btn_cancel_scan.set_enabled(False)
                    has_files = len(self.all_files) > 0
                    self._set_organize_buttons_enabled(has_files)
                    self._set_export_buttons_enabled(has_files)

                    summary_stats = [
                        ("Total Files Discovered:", f"{len(self.all_files):,} files"),
                        ("Subfolders Scanned:", f"{fld_cnt:,} folders"),
                        ("Total Data Analyzed:", format_size(tot_sz)),
                        ("Inaccessible / Locked Files:", f"{inacc:,} files"),
                        ("Scan Duration:", format_time(el)),
                    ]
                    CompletionSummaryDialog(
                        parent=self.root,
                        title="Scan Summary",
                        stats=summary_stats,
                        open_folder_cmd=self._open_selected_folder
                    )

                elif msg_type == "scan_cancelled":
                    self.progressbar.set_value(0)
                    self.lbl_percentage.configure(text="0%")
                    self.lbl_operation.configure(text="Scan was cancelled by user.")
                    self._update_status_indicator("Cancelled", "Cancelled")
                    self.btn_analyze.set_enabled(True)
                    self.btn_browse.set_enabled(True)
                    self.btn_cancel_scan.set_enabled(False)

                elif msg_type == "organize_progress":
                    proc = msg["processed"]
                    tot = msg["total"]
                    cur = msg["current_file"]
                    el = msg["elapsed"]
                    speed = msg["speed"]
                    eta = msg["eta"]

                    pct = int((proc / tot) * 100) if tot > 0 else 0
                    self.progressbar.set_value(pct)
                    self.lbl_percentage.configure(text=f"{pct}%")
                    self.lbl_operation.configure(text=f"Organizing file {proc:,} of {tot:,}...")
                    self.lbl_current_file.configure(text=f"Current: {truncate_path(cur, 35)}")
                    self.lbl_processed_count.configure(text=f"Processed: {proc:,} / {tot:,}")
                    self.lbl_elapsed_time.configure(text=f"Elapsed: {format_time(el)}")
                    self.lbl_eta_time.configure(text=f"ETA: {format_time(eta)}" if el > 1.2 and proc > 5 else "ETA: Calculating...")
                    self.lbl_speed.configure(text=f"Speed: {speed:.1f} files/sec" if el > 1.0 else "Speed: Calculating...")

                elif msg_type == "organize_complete":
                    was_canc = msg["cancelled"]
                    succ = msg["success"]
                    skip = msg["skipped"]
                    fail = msg["failed"]
                    bytes_mv = msg["bytes_moved"]
                    el = msg["elapsed"]
                    records = msg["undo_records"]
                    self.last_errors = msg["errors"]
                    is_dry = msg["dry_run"]

                    if records and not is_dry:
                        self.undo_stack.append(records)
                        self.btn_undo.set_enabled(True)

                    self.progressbar.set_value(100)
                    self.lbl_percentage.configure(text="100%")

                    status_title = "Simulation Complete" if is_dry else "Organization Complete"
                    if was_canc:
                        status_title = "Organization Cancelled"
                        self._update_status_indicator("Cancelled", "Cancelled")
                    else:
                        self._update_status_indicator("Completed", "Completed")

                    self.lbl_operation.configure(text=f"{status_title}: {succ:,} files organized, {fail:,} failed")
                    self.lbl_status_bar.configure(text=f"Operation finished in {format_time(el)}. {succ:,} moved, {fail:,} failed.")

                    self.btn_analyze.set_enabled(True)
                    self.btn_browse.set_enabled(True)
                    self.btn_cancel_scan.set_enabled(False)
                    self._set_organize_buttons_enabled(True)
                    self._set_export_buttons_enabled(True)

                    self._render_current_page()

                    summary_stats = [
                        ("Successfully Organized:", f"{succ:,} files"),
                        ("Skipped (In place / Excluded):", f"{skip:,} files"),
                        ("Failed Operations:", f"{fail:,} files"),
                        ("Total Data Moved:", format_size(bytes_mv)),
                        ("Time Taken:", format_time(el)),
                        ("Mode:", "Preview (Dry Run)" if is_dry else "Live In-Place Move"),
                    ]
                    err_cmd = self._show_error_log if self.last_errors else None
                    CompletionSummaryDialog(
                        parent=self.root,
                        title="Organization Summary",
                        stats=summary_stats,
                        open_folder_cmd=self._open_selected_folder,
                        view_errors_cmd=err_cmd
                    )

                elif msg_type == "export_progress":
                    proc = msg["processed"]
                    tot = msg["total"]
                    cur = msg["current_file"]
                    el = msg["elapsed"]
                    speed = msg["speed"]
                    eta = msg["eta"]

                    pct = int((proc / tot) * 100) if tot > 0 else 0
                    self.progressbar.set_value(pct)
                    self.lbl_percentage.configure(text=f"{pct}%")
                    self.lbl_operation.configure(text=f"Exporting file {proc:,} of {tot:,}...")
                    self.lbl_current_file.configure(text=f"Current: {truncate_path(cur, 35)}")
                    self.lbl_processed_count.configure(text=f"Processed: {proc:,} / {tot:,}")
                    self.lbl_elapsed_time.configure(text=f"Elapsed: {format_time(el)}")
                    self.lbl_eta_time.configure(text=f"ETA: {format_time(eta)}" if el > 1.2 and proc > 5 else "ETA: Calculating...")
                    self.lbl_speed.configure(text=f"Speed: {speed:.1f} files/sec" if el > 1.0 else "Speed: Calculating...")

                elif msg_type == "export_complete":
                    was_canc = msg["cancelled"]
                    succ = msg["success"]
                    fail = msg["failed"]
                    bytes_tr = msg["bytes_transferred"]
                    export_dir = msg["export_dir"]
                    op = msg["operation"]
                    el = msg["elapsed"]
                    records = msg["undo_records"]
                    self.last_errors = msg["errors"]
                    self.last_export_dir = export_dir

                    if records:
                        self.undo_stack.append(records)
                        self.btn_undo.set_enabled(True)

                    self.progressbar.set_value(100)
                    self.lbl_percentage.configure(text="100%")

                    if was_canc:
                        self._update_status_indicator("Cancelled", "Cancelled")
                        self.lbl_operation.configure(text=f"Export Cancelled: {succ:,} exported before stop.")
                    else:
                        self._update_status_indicator("Completed", "Completed")
                        self.lbl_operation.configure(text=f"Export Complete: {succ:,} files exported into '{export_dir.name}'.")

                    self.lbl_status_bar.configure(text=f"Export finished in {format_time(el)}. {succ:,} files transferred.")

                    self.btn_analyze.set_enabled(True)
                    self.btn_browse.set_enabled(True)
                    self.btn_cancel_scan.set_enabled(False)
                    self._set_organize_buttons_enabled(True)
                    self._set_export_buttons_enabled(True)

                    self._render_current_page()

                    summary_stats = [
                        ("Successfully Exported:", f"{succ:,} files"),
                        ("Failed Operations:", f"{fail:,} files"),
                        ("Total Data Exported:", format_size(bytes_tr)),
                        ("Export Destination:", str(export_dir)),
                        ("Operation Mode:", "Copied (Originals Untouched)" if op == "copy" else "Moved"),
                        ("Time Taken:", format_time(el)),
                    ]
                    err_cmd = self._show_error_log if self.last_errors else None
                    CompletionSummaryDialog(
                        parent=self.root,
                        title="Export Complete",
                        stats=summary_stats,
                        open_folder_cmd=self._open_export_folder,
                        view_errors_cmd=err_cmd
                    )

                elif msg_type == "undo_progress":
                    proc = msg["processed"]
                    tot = msg["total"]
                    cur = msg["current_file"]
                    self.lbl_operation.configure(text=f"Undoing: Restored {proc:,} of {tot:,} files...")
                    self.lbl_current_file.configure(text=f"Current: {truncate_path(cur, 35)}")

                elif msg_type == "undo_complete":
                    self.progressbar.set_value(100)
                    restored = msg["restored"]
                    tot = msg["total"]
                    fail = msg["failed"]
                    el = msg["elapsed"]
                    self.last_errors = msg["errors"]

                    if self.undo_stack:
                        self.undo_stack.pop()
                    self.btn_undo.set_enabled(len(self.undo_stack) > 0)

                    self.lbl_operation.configure(text=f"Undo complete: {restored:,} of {tot:,} files restored.")
                    self._update_status_indicator("Completed", "Completed")

                    messagebox.showinfo(
                        "Undo Finished",
                        f"Undo operation complete in {format_time(el)}.\n\nRestored: {restored:,} files\nFailed: {fail:,} files"
                    )

                    self._start_scan()

        except queue.Empty:
            pass
        finally:
            self.root.after(35, self._poll_queue)

    # ========================================================================
    # TABLE & CATEGORY RENDERING
    # ========================================================================
    def _update_status_indicator(self, text: str, status_key: str):
        """Update top-right badge."""
        self.status_badge.update_status(text, status_key)

    def _populate_category_tree(self, cat_counts: Dict[str, int], cat_sizes: Dict[str, int]):
        """Render category counts in left overview pane."""
        for item in self.cat_tree.get_children():
            self.cat_tree.delete(item)

        total_files = sum(cat_counts.values())
        total_size = sum(cat_sizes.values())

        # Top row: All Categories
        self.cat_tree.insert(
            "",
            "end",
            iid="all",
            values=("📁 All Categories", f"{total_files:,}", format_size(total_size))
        )

        for cat_name, exts in CATEGORIES.items():
            cnt = cat_counts.get(cat_name, 0)
            sz = cat_sizes.get(cat_name, 0)
            icon = CATEGORY_ICONS.get(cat_name, "📁")
            self.cat_tree.insert(
                "",
                "end",
                iid=cat_name,
                values=(f"{icon} {cat_name}", f"{cnt:,}", format_size(sz))
            )

    def _on_category_selected(self, event):
        """When a category is clicked in the left pane, filter preview table."""
        selected = self.cat_tree.selection()
        if not selected:
            return
        cat_id = selected[0]
        if cat_id == "all":
            self.cat_filter_var.set("All Categories")
            self.lbl_ext_info.configure(text="Extensions: Showing all detected file types.")
        else:
            self.cat_filter_var.set(cat_id)
            exts = CATEGORIES.get(cat_id, [])
            ext_str = ", ".join(exts) if exts else "All unrecognized extensions"
            self.lbl_ext_info.configure(text=f"{CATEGORY_ICONS.get(cat_id, '')} {cat_id} Extensions:\n{ext_str}")

        self._apply_filters()

    def _set_category_inclusion(self, include: bool):
        """Include or exclude all files in currently filtered category."""
        target_cat = self.cat_filter_var.get()
        count = 0
        for f in self.all_files:
            if target_cat == "All Categories" or f.category == target_cat:
                if f.status != "Already in Place":
                    f.is_included = include
                    count += 1
        self._render_current_page()
        self.lbl_status_bar.configure(text=f"{'Included' if include else 'Excluded'} {count:,} file(s).")

    def _apply_filters(self):
        """Filter files by search query, category, and status."""
        query = self.search_var.get().strip().lower()
        selected_cat = self.cat_filter_var.get()
        selected_status = self.status_filter_var.get()

        results = []
        for it in self.all_files:
            if query and query not in it.filename.lower() and query not in it.relative_path.lower():
                continue
            if selected_cat != "All Categories" and it.category != selected_cat:
                continue
            if selected_status != "All Statuses":
                if selected_status == "Excluded" and it.is_included:
                    continue
                elif selected_status != "Excluded" and it.status != selected_status:
                    continue

            results.append(it)

        self.filtered_files = results
        self.current_page = 1
        self._sort_filtered_files()
        self._render_current_page()

    def _clear_filters(self):
        """Reset search and filter dropdowns."""
        self.search_var.set("")
        self.cat_filter_var.set("All Categories")
        self.status_filter_var.set("All Statuses")
        self.lbl_ext_info.configure(text="Extensions: Showing all detected file types.")
        self._apply_filters()

    def _sort_by_column(self, col: str):
        """Toggle column sort order and re-render."""
        if self.sort_column == col:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = col
            self.sort_ascending = True

        self._sort_filtered_files()
        self._render_current_page()

    def _sort_filtered_files(self):
        """Sort filtered in-memory list."""
        col = self.sort_column
        rev = not self.sort_ascending

        if col == "include":
            self.filtered_files.sort(key=lambda x: x.is_included, reverse=rev)
        elif col == "filename":
            self.filtered_files.sort(key=lambda x: x.filename.lower(), reverse=rev)
        elif col == "category":
            self.filtered_files.sort(key=lambda x: x.category.lower(), reverse=rev)
        elif col == "extension":
            self.filtered_files.sort(key=lambda x: x.extension.lower(), reverse=rev)
        elif col == "size_bytes":
            self.filtered_files.sort(key=lambda x: x.size_bytes, reverse=rev)
        elif col == "destination":
            self.filtered_files.sort(key=lambda x: str(x.proposed_destination or "").lower(), reverse=rev)
        elif col == "status":
            self.filtered_files.sort(key=lambda x: x.status.lower(), reverse=rev)

    def _render_current_page(self):
        """Render paginated rows into Treeview for instant, smooth UI response."""
        self._clear_tree()

        total = len(self.filtered_files)
        if total == 0:
            self.lbl_table_count.configure(text="Showing 0 of 0 files")
            self.lbl_page_info.configure(text="Page 1 of 1")
            self.btn_prev_page.set_enabled(False)
            self.btn_next_page.set_enabled(False)
            return

        if self.page_size == -1:
            start_idx = 0
            end_idx = total
            total_pages = 1
            self.current_page = 1
        else:
            total_pages = max(1, math.ceil(total / self.page_size))
            self.current_page = max(1, min(self.current_page, total_pages))
            start_idx = (self.current_page - 1) * self.page_size
            end_idx = min(start_idx + self.page_size, total)

        page_items = self.filtered_files[start_idx:end_idx]

        for i, item in enumerate(page_items):
            check_char = "✓" if item.is_included else " "
            dest_str = item.proposed_destination.name if item.proposed_destination else "--"
            if item.status == "Conflict - Renamed" and item.proposed_destination:
                dest_str = f"↳ {item.proposed_destination.name} (renamed)"

            tag_list = []
            if i % 2 == 1:
                tag_list.append("altrow")

            if not item.is_included:
                tag_list.append("excluded")
                display_status = "Excluded"
            elif item.status == "Ready to Move":
                tag_list.append("ready")
                display_status = "Ready to Move"
            elif item.status == "Conflict - Renamed":
                tag_list.append("conflict")
                display_status = "Conflict (Auto-Renamed)"
            elif item.status == "Already in Place":
                tag_list.append("inplace")
                display_status = "Already in Place"
            elif item.status == "Failed":
                tag_list.append("failed")
                display_status = "Failed"
            elif item.status == "Moved":
                tag_list.append("moved")
                display_status = "Moved"
            elif item.status == "Copied":
                tag_list.append("copied")
                display_status = "Copied"
            else:
                display_status = item.status

            self.file_tree.insert(
                "",
                "end",
                iid=str(start_idx + i),
                values=(
                    check_char,
                    item.filename,
                    f"{CATEGORY_ICONS.get(item.category, '')} {item.category}",
                    item.extension,
                    format_size(item.size_bytes),
                    dest_str,
                    display_status
                ),
                tags=tuple(tag_list)
            )

        self.lbl_table_count.configure(text=f"Showing {start_idx+1:,}–{end_idx:,} of {total:,} files")
        self.lbl_page_info.configure(text=f"Page {self.current_page} of {total_pages}")
        self.btn_prev_page.set_enabled(self.current_page > 1)
        self.btn_next_page.set_enabled(self.current_page < total_pages)

    def _clear_tree(self):
        for row in self.file_tree.get_children():
            self.file_tree.delete(row)

    def _prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._render_current_page()

    def _next_page(self):
        total_pages = math.ceil(len(self.filtered_files) / self.page_size) if self.page_size > 0 else 1
        if self.current_page < total_pages:
            self.current_page += 1
            self._render_current_page()

    def _on_page_size_changed(self, event):
        val = self.page_size_var.get()
        self.page_size = -1 if val == "All" else int(val)
        self.current_page = 1
        self._render_current_page()

    def _on_row_double_click(self, event):
        selected = self.file_tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        if 0 <= idx < len(self.filtered_files):
            item = self.filtered_files[idx]
            if item.status != "Already in Place":
                item.is_included = not item.is_included
            self._render_current_page()

    def _on_space_pressed(self, event):
        selected = self.file_tree.selection()
        if not selected:
            return
        for s in selected:
            idx = int(s)
            if 0 <= idx < len(self.filtered_files):
                item = self.filtered_files[idx]
                if item.status != "Already in Place":
                    item.is_included = not item.is_included
        self._render_current_page()

    def _toggle_selection_all(self, include: bool):
        for item in self.filtered_files:
            if item.status != "Already in Place":
                item.is_included = include
        self._render_current_page()

    def _invert_selection(self):
        for item in self.filtered_files:
            if item.status != "Already in Place":
                item.is_included = not item.is_included
        self._render_current_page()

    def _preview_selected_image(self):
        selected = self.file_tree.selection()
        if not selected:
            messagebox.showinfo("Preview", "Select an image file in the table first.")
            return
        idx = int(selected[0])
        item = self.filtered_files[idx]
        if item.category != "Images":
            messagebox.showinfo("Preview", f"'{item.filename}' is categorized as {item.category}, not an Image.")
            return
        ImagePreviewDialog(self.root, item)

    def _show_context_menu(self, event):
        row_id = self.file_tree.identify_row(event.y)
        if not row_id:
            return
        self.file_tree.selection_set(row_id)
        idx = int(row_id)
        item = self.filtered_files[idx]

        menu = tk.Menu(self.root, tearoff=0, bg=COLOR_BG_CARD, fg=COLOR_TEXT_MAIN, font=FONT_BODY)
        menu.add_command(
            label="📦 Export to Folder (Separate by Extension / Category)...",
            command=self._start_export_dialog
        )
        menu.add_separator()
        menu.add_command(
            label="Toggle Include / Exclude",
            command=lambda: self._toggle_single_item(item)
        )
        menu.add_separator()
        menu.add_command(
            label="Open Containing Folder in Explorer",
            command=lambda: self._open_path_in_explorer(item.original_path.parent)
        )
        if HAS_PILLOW and item.category == "Images":
            menu.add_command(
                label="👁 Preview Image Thumbnail",
                command=lambda: ImagePreviewDialog(self.root, item)
            )
        menu.add_separator()
        menu.add_command(
            label="Copy File Path to Clipboard",
            command=lambda: self._copy_to_clipboard(str(item.original_path))
        )
        menu.tk_popup(event.x_root, event.y_root)

    def _toggle_single_item(self, item: FileItem):
        if item.status != "Already in Place":
            item.is_included = not item.is_included
            self._render_current_page()

    def _copy_to_clipboard(self, text: str):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.lbl_status_bar.configure(text=f"Copied to clipboard: {text}")

    def _open_selected_folder(self):
        """Open the active source folder in Windows Explorer."""
        if self.selected_folder and self.selected_folder.is_dir():
            self._open_path_in_explorer(self.selected_folder)
        else:
            messagebox.showinfo("Open Folder", "No active folder selected.")

    def _open_export_folder(self):
        """Open the latest export folder or category folder in Windows Explorer."""
        if self.last_export_dir and self.last_export_dir.is_dir():
            self._open_path_in_explorer(self.last_export_dir)
        elif self.selected_folder:
            cat = self.cat_filter_var.get()
            if cat == "All Categories":
                cat_dir = self.selected_folder
            else:
                cat_dir = self.selected_folder / cat
            cat_dir.mkdir(parents=True, exist_ok=True)
            self._open_path_in_explorer(cat_dir)
        else:
            messagebox.showinfo("Open Folder", "No export folder available yet.")

    def _open_path_in_explorer(self, path: Path):
        """Safely launch Windows Explorer to show folder or select file."""
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Error Opening Explorer", f"Could not open '{path}':\n{e}")

    def _show_error_log(self):
        if self.last_errors:
            ErrorLogDialog(self.root, self.last_errors)
        else:
            messagebox.showinfo("Errors", "No errors recorded in the last operation.")

    def _reset_application(self):
        """Reset application state back to clean initial state."""
        if self.current_worker and self.current_worker.is_alive():
            confirm = messagebox.askyesno("Operation Active", "An operation is currently running. Stop and reset?")
            if not confirm:
                return
            self.cancel_event.set()

        self.selected_folder = None
        self.last_export_dir = None
        self.path_var.set("")
        self.all_files.clear()
        self.filtered_files.clear()
        self._clear_tree()

        for item in self.cat_tree.get_children():
            self.cat_tree.delete(item)

        self.card_total_files.update_card("0", "Discovered")
        self.card_total_folders.update_card("0", "Directories")
        self.card_total_size.update_card("0 B", "Disk Usage")
        self.card_images.update_card("0", "0 B")
        self.card_videos.update_card("0", "0 B")
        self.card_docs.update_card("0", "0 B")
        self.card_other.update_card("0", "0 B")

        self.progressbar.set_value(0)
        self.lbl_percentage.configure(text="0%")
        self.lbl_operation.configure(text="Ready • Select folder and click Analyze")
        self.lbl_current_file.configure(text="Current: None")
        self.lbl_processed_count.configure(text="Processed: 0 / 0")
        self.lbl_elapsed_time.configure(text="Elapsed: 0 sec")
        self.lbl_eta_time.configure(text="ETA: --")
        self.lbl_speed.configure(text="Speed: -- files/sec")
        self.lbl_folder_info.configure(text="No folder selected • Choose a folder to begin analysis")
        self.lbl_status_bar.configure(text="Ready • Select a folder to begin")
        self._update_status_indicator("Ready", "Ready")

        self.btn_analyze.set_enabled(True)
        self.btn_browse.set_enabled(True)
        self.btn_cancel_scan.set_enabled(False)
        self._set_organize_buttons_enabled(False)
        self._set_export_buttons_enabled(False)

    def _on_window_close(self):
        """Safely handle window close event when tasks are active."""
        if self.current_worker and self.current_worker.is_alive():
            confirm = messagebox.askyesno(
                "Exit Confirmation",
                "An operation is currently in progress.\n\nAre you sure you want to cancel and exit SmartSort?"
            )
            if not confirm:
                return
            self.cancel_event.set()
            self.current_worker.join(timeout=1.0)
        self.root.destroy()


# ============================================================================
# ENTRY POINT
# ============================================================================
def main():
    """Launch SmartSort application."""
    root = tk.Tk()
    app = SmartSortApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
