"""
PyCommander - Norton Commander clone dla Windows 11
v2: F2 zmiana nazwy, otwieranie plików, wybór dysku, zaznaczanie wielu, kosmetyka
"""

import sys
import os
import shutil
import datetime
import string
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QLabel, QFrame, QLineEdit,
    QDialog, QDialogButtonBox, QMessageBox, QHeaderView, QStatusBar,
    QComboBox, QSizePolicy
)
from PyQt6.QtCore import Qt, QEvent
from PyQt6.QtGui import QKeyEvent, QColor, QKeySequence, QShortcut

try:
    from win32com.shell import shell as win32shell, shellcon
    SHELL_API = True
except ImportError:
    SHELL_API = False


# ─── Kolory ──────────────────────────────────────────────────────────────────
BG               = "#001F5C"
FG               = "#E8E8E8"
SEL_BG           = "#007A7A"
SEL_FG           = "#FFFFFF"
MULTI_BG         = "#8B0000"
MULTI_FG         = "#FFFF99"
ACTIVE_BORDER    = "#FFD700"
INACTIVE_BORDER  = "#3A5A9A"
HEADER_BG        = "#005F87"
HEADER_FG        = "#FFFFFF"
PATH_BG          = "#003070"
FKEY_BG          = "#005F87"
FKEY_FG          = "#FFFFFF"
STATUS_BG        = "#001040"
STATUS_FG        = "#AAD4FF"
DIR_COLOR        = "#7FD4FF"
FILE_COLOR       = "#E8E8E8"
INPUT_BG         = "#000A30"
INPUT_FG         = "#FFD700"


def available_drives() -> list[str]:
    drives = []
    if sys.platform == "win32":
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append(drive)
    else:
        drives = ["/"]
    return drives


def fmt_size(size: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024:
            return f"{size:.0f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def free_space(path: str) -> str:
    try:
        stat = shutil.disk_usage(path)
        return f"Wolne: {fmt_size(stat.free)} / {fmt_size(stat.total)}"
    except Exception:
        return ""


# ─── Dialog z polem tekstowym ─────────────────────────────────────────────────
class InputDialog(QDialog):
    def __init__(self, parent, title: str, label: str, default: str = ""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(420)
        self.setStyleSheet(f"QDialog {{ background:{BG}; color:{FG}; }}")
        layout = QVBoxLayout(self)
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{FG}; font-size:13px;")
        self.inp = QLineEdit(default)
        self.inp.setStyleSheet(
            f"background:{INPUT_BG}; color:{INPUT_FG}; border:1px solid {INPUT_FG}; "
            f"padding:4px; font-size:13px; font-family:Consolas;"
        )
        self.inp.selectAll()
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.setStyleSheet(
            f"QPushButton {{ background:{HEADER_BG}; color:{FG}; border:1px solid {INACTIVE_BORDER}; padding:4px 12px; }}"
            f"QPushButton:hover {{ background:{SEL_BG}; }}"
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(lbl)
        layout.addWidget(self.inp)
        layout.addWidget(btns)

    def value(self) -> str:
        return self.inp.text().strip()


# ─── Panel plików ─────────────────────────────────────────────────────────────
class FilePanel(QFrame):
    def __init__(self, start_path: str, parent=None):
        super().__init__(parent)
        self.current_path = os.path.abspath(start_path)
        self.items: list[dict] = []
        self.search_filter = ""
        self.marked: set[str] = set()

        self._build_ui()
        self.refresh()

    def _build_ui(self):
        self.setFrameShape(QFrame.Shape.Box)
        self.setLineWidth(2)
        self._set_inactive_style()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Górny pasek: dysk + ścieżka
        top = QHBoxLayout()
        top.setContentsMargins(2, 2, 2, 2)
        top.setSpacing(4)

        self.drive_combo = QComboBox()
        self.drive_combo.setFixedWidth(72)
        self.drive_combo.setStyleSheet(
            f"QComboBox {{ background:{PATH_BG}; color:{INPUT_FG}; border:1px solid {INACTIVE_BORDER}; "
            f"font-weight:bold; font-size:12px; padding:1px 4px; }}"
            f"QComboBox QAbstractItemView {{ background:{BG}; color:{FG}; }}"
        )
        for d in available_drives():
            self.drive_combo.addItem(d)
        self.drive_combo.currentTextChanged.connect(self._on_drive_changed)
        top.addWidget(self.drive_combo)

        self.path_label = QLabel(self.current_path)
        self.path_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.path_label.setStyleSheet(
            f"background:{PATH_BG}; color:{INPUT_FG}; font-weight:bold; "
            f"font-size:12px; padding:2px 6px; font-family:Consolas;"
        )
        self.path_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top.addWidget(self.path_label)
        layout.addLayout(top)

        # Tabela
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["  Nazwa", "Rozmiar", "Data modyfikacji"])
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.table.verticalHeader().setDefaultSectionSize(20)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setStyleSheet(
            f"QHeaderView::section {{ background:{HEADER_BG}; color:{HEADER_FG}; "
            f"font-weight:bold; border:none; border-bottom:1px solid {INACTIVE_BORDER}; padding:3px; }}"
        )
        self.table.setStyleSheet(
            f"QTableWidget {{ background:{BG}; color:{FG}; border:none; "
            f"font-family:Consolas; font-size:13px; outline:none; }}"
            f"QTableWidget::item:selected {{ background:{SEL_BG}; color:{SEL_FG}; }}"
            f"QTableWidget::item:focus {{ background:{SEL_BG}; color:{SEL_FG}; border:none; }}"
        )
        self.table.itemActivated.connect(self._on_activate)
        self.table.installEventFilter(self)
        layout.addWidget(self.table)

        # Pasek wyszukiwania
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("  Szukaj... (Esc aby wyczyścić)")
        self.search_bar.setStyleSheet(
            f"background:{INPUT_BG}; color:{INPUT_FG}; border:1px solid {INPUT_FG}; "
            f"padding:3px; font-size:13px; font-family:Consolas;"
        )
        self.search_bar.textChanged.connect(self._on_search_changed)
        self.search_bar.hide()
        layout.addWidget(self.search_bar)

        # Pasek wolnego miejsca
        self.space_label = QLabel("")
        self.space_label.setStyleSheet(
            f"background:{STATUS_BG}; color:{STATUS_FG}; font-size:11px; padding:2px 6px;"
        )
        layout.addWidget(self.space_label)

    def _set_active_style(self):
        self.setStyleSheet(f"QFrame {{ border:2px solid {ACTIVE_BORDER}; }}")

    def _set_inactive_style(self):
        self.setStyleSheet(f"QFrame {{ border:2px solid {INACTIVE_BORDER}; }}")

    # ── Dyski ────────────────────────────────────────────────────────────────
    def _on_drive_changed(self, drive: str):
        if drive and os.path.exists(drive):
            self.navigate_to(drive)

    def _sync_drive_combo(self):
        if sys.platform == "win32":
            drive = os.path.splitdrive(self.current_path)[0] + "\\"
            idx = self.drive_combo.findText(drive)
            if idx >= 0:
                self.drive_combo.blockSignals(True)
                self.drive_combo.setCurrentIndex(idx)
                self.drive_combo.blockSignals(False)

    # ── Odświeżanie ──────────────────────────────────────────────────────────
    def refresh(self, keep_selection: str = ""):
        self.path_label.setText(self.current_path)
        self._sync_drive_combo()
        self.space_label.setText(free_space(self.current_path))
        self.items = []

        if os.path.dirname(self.current_path) != self.current_path:
            self.items.append({"name": "..", "size": "", "date": "", "is_dir": True})

        try:
            entries = sorted(
                os.scandir(self.current_path),
                key=lambda e: (not e.is_dir(), e.name.lower())
            )
            for e in entries:
                if self.search_filter and self.search_filter.lower() not in e.name.lower():
                    continue
                try:
                    stat = e.stat()
                    size = "" if e.is_dir() else fmt_size(stat.st_size)
                    date = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    size, date = "", ""
                self.items.append({"name": e.name, "size": size, "date": date, "is_dir": e.is_dir()})
        except PermissionError:
            QMessageBox.warning(self, "Brak dostępu", f"Brak uprawnień do:\n{self.current_path}")

        self.table.setRowCount(len(self.items))
        for row, item in enumerate(self.items):
            is_marked = item["name"] in self.marked
            if item["name"] == "..":
                prefix = "  \u2191  "
            elif item["is_dir"]:
                prefix = "  \u25b6  "
            else:
                prefix = "     "
            if is_marked:
                prefix = "  \u2713  "

            name_cell = QTableWidgetItem(prefix + item["name"])
            size_cell = QTableWidgetItem(item["size"])
            date_cell = QTableWidgetItem(item["date"])
            size_cell.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            if is_marked:
                for cell in (name_cell, size_cell, date_cell):
                    cell.setBackground(QColor(MULTI_BG))
                    cell.setForeground(QColor(MULTI_FG))
            elif item["is_dir"]:
                for cell in (name_cell, size_cell, date_cell):
                    cell.setForeground(QColor(DIR_COLOR))
                    f = cell.font(); f.setBold(True); cell.setFont(f)
            else:
                for cell in (name_cell, size_cell, date_cell):
                    cell.setForeground(QColor(FILE_COLOR))

            self.table.setItem(row, 0, name_cell)
            self.table.setItem(row, 1, size_cell)
            self.table.setItem(row, 2, date_cell)

        if keep_selection:
            for row, item in enumerate(self.items):
                if item["name"] == keep_selection:
                    self.table.selectRow(row)
                    return
        if self.table.rowCount() > 0:
            self.table.selectRow(0)

    # ── Nawigacja ────────────────────────────────────────────────────────────
    def selected_item(self) -> dict | None:
        row = self.table.currentRow()
        return self.items[row] if 0 <= row < len(self.items) else None

    def selected_path(self) -> str | None:
        item = self.selected_item()
        if item:
            return os.path.join(self.current_path, item["name"])
        return None

    def marked_paths(self) -> list[str]:
        if self.marked:
            return [os.path.join(self.current_path, n) for n in self.marked]
        p = self.selected_path()
        sel = self.selected_item()
        if p and sel and sel["name"] != "..":
            return [p]
        return []

    def navigate_to(self, path: str):
        self.marked.clear()
        self.current_path = os.path.abspath(path)
        self.refresh()

    def go_up(self):
        parent = os.path.dirname(self.current_path)
        if parent != self.current_path:
            old_name = os.path.basename(self.current_path)
            self.marked.clear()
            self.current_path = parent
            self.refresh(keep_selection=old_name)

    def _on_activate(self, _item):
        sel = self.selected_item()
        if not sel:
            return
        if sel["is_dir"]:
            if sel["name"] == "..":
                self.go_up()
            else:
                self.navigate_to(os.path.join(self.current_path, sel["name"]))
        else:
            path = os.path.join(self.current_path, sel["name"])
            try:
                if sys.platform == "win32":
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception as e:
                QMessageBox.warning(self, "Błąd otwierania", str(e))

    # ── Zaznaczanie ──────────────────────────────────────────────────────────
    def toggle_mark(self):
        sel = self.selected_item()
        if not sel or sel["name"] == "..":
            return
        name = sel["name"]
        row = self.table.currentRow()
        if name in self.marked:
            self.marked.discard(name)
        else:
            self.marked.add(name)
        self.refresh()
        next_row = min(row + 1, self.table.rowCount() - 1)
        self.table.selectRow(next_row)

    def clear_marks(self):
        self.marked.clear()
        self.refresh()

    # ── Wyszukiwanie ─────────────────────────────────────────────────────────
    def show_search(self):
        self.search_bar.show()
        self.search_bar.setFocus()

    def hide_search(self):
        self.search_bar.hide()
        self.search_bar.clear()
        self.search_filter = ""
        self.refresh()
        self.table.setFocus()

    def _on_search_changed(self, text: str):
        self.search_filter = text
        self.refresh()

    # ── EventFilter tabeli ───────────────────────────────────────────────────
    def eventFilter(self, source, event):
        if source is self.table and event.type() == QEvent.Type.KeyPress:
            key = event.key()
            if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                self._on_activate(None)
                return True
            elif key == Qt.Key.Key_Backspace:
                self.go_up()
                return True
            elif key == Qt.Key.Key_Escape:
                if self.search_bar.isVisible():
                    self.hide_search()
                    return True
            elif key in (Qt.Key.Key_Insert, Qt.Key.Key_Space):
                self.toggle_mark()
                return True
        return super().eventFilter(source, event)


# ─── Główne okno ──────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyCommander by HACS9")
        self.setMinimumSize(1000, 600)
        self.resize(1280, 750)
        self.setStyleSheet(f"QMainWindow, QWidget {{ background:{BG}; }}")

        start = os.path.expanduser("~")
        self.left  = FilePanel(start)
        self.right = FilePanel(start)
        self.active_panel: FilePanel = self.left

        self._build_ui()
        self._setup_shortcuts()
        self._activate_panel(self.left)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(4, 4, 4, 2)
        main_layout.setSpacing(3)

        panels = QHBoxLayout()
        panels.setSpacing(6)
        panels.addWidget(self.left)
        panels.addWidget(self.right)
        main_layout.addLayout(panels)

        # Pasek F-key
        fkeys = QHBoxLayout()
        fkeys.setSpacing(2)
        for key, label in [
            ("F2", "Zmień nazwę"), ("F5", "Kopiuj"), ("F6", "Przesuń"),
            ("F7", "Nowy folder"), ("F8", "Usuń"),
            ("Ins", "Zaznacz"), ("Tab", "Przełącz"),
            ("Ctrl+D", "Dysk"),
            ("Alt+F4", "Wyjście"),
        ]:
            btn = QLabel(f"<b style='color:{INPUT_FG}'>{key}</b> {label}")
            btn.setAlignment(Qt.AlignmentFlag.AlignCenter)
            btn.setStyleSheet(
                f"background:{FKEY_BG}; color:{FKEY_FG}; padding:4px 6px; "
                f"border:1px solid #007A9A; border-radius:2px; font-size:12px;"
            )
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            fkeys.addWidget(btn)
        main_layout.addLayout(fkeys)

        self.status = QStatusBar()
        self.status.setStyleSheet(f"background:{STATUS_BG}; color:{STATUS_FG}; font-size:12px;")
        self.setStatusBar(self.status)
        self.status.showMessage("Gotowy  │  Tab: przełącz panel  │  /: szukaj  │  Ins/Spacja: zaznacz")

        self.left.table.installEventFilter(self)
        self.right.table.installEventFilter(self)

    def eventFilter(self, source, event):
        if event.type() == QEvent.Type.FocusIn:
            if source is self.left.table:
                self._activate_panel(self.left)
            elif source is self.right.table:
                self._activate_panel(self.right)
        return super().eventFilter(source, event)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Tab"),    self, self._switch_panel)
        QShortcut(QKeySequence("F2"),     self, self._rename)
        QShortcut(QKeySequence("F5"),     self, self._copy)
        QShortcut(QKeySequence("F6"),     self, self._move)
        QShortcut(QKeySequence("F7"),     self, self._mkdir)
        QShortcut(QKeySequence("F8"),     self, self._delete)
        QShortcut(QKeySequence("Delete"), self, self._delete)
        QShortcut(QKeySequence("/"),      self, lambda: self.active_panel.show_search())
        QShortcut(QKeySequence("Ctrl+R"), self, self._refresh_both)
        QShortcut(QKeySequence("Ctrl+D"), self, lambda: self._pick_drive(self.active_panel))

    def _activate_panel(self, panel: FilePanel):
        self.active_panel = panel
        self.left._set_active_style()  if panel is self.left  else self.left._set_inactive_style()
        self.right._set_active_style() if panel is self.right else self.right._set_inactive_style()
        panel.table.setFocus()

    def _switch_panel(self):
        self._activate_panel(self.right if self.active_panel is self.left else self.left)

    def _other_panel(self) -> FilePanel:
        return self.right if self.active_panel is self.left else self.left

    def _refresh_both(self):
        self.left.refresh()
        self.right.refresh()

    # ── Alt+F1 / Alt+F2 — wybór dysku klawiaturą ────────────────────────────
    def _pick_drive(self, panel: FilePanel):
        drives = available_drives()
        if not drives:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Wybierz dysk")
        dlg.setFixedWidth(220)
        dlg.setStyleSheet(f"QDialog {{ background:{BG}; color:{FG}; }}")
        layout = QVBoxLayout(dlg)

        lbl = QLabel("Wybierz dysk:  (↑↓ Enter)")
        lbl.setStyleSheet(f"color:{STATUS_FG}; font-size:12px; padding:2px;")
        layout.addWidget(lbl)

        tbl = QTableWidget(len(drives), 1)
        tbl.setHorizontalHeaderLabels(["Dysk"])
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        tbl.verticalHeader().setVisible(False)
        tbl.setShowGrid(False)
        tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tbl.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        tbl.verticalHeader().setDefaultSectionSize(24)
        tbl.setStyleSheet(
            f"QTableWidget {{ background:{BG}; color:{INPUT_FG}; border:none; "
            f"font-family:Consolas; font-size:14px; font-weight:bold; }}"
            f"QTableWidget::item:selected {{ background:{SEL_BG}; color:{SEL_FG}; }}"
        )
        tbl.horizontalHeader().setStyleSheet(
            f"QHeaderView::section {{ background:{HEADER_BG}; color:{HEADER_FG}; border:none; padding:2px; }}"
        )

        current_drive = os.path.splitdrive(panel.current_path)[0] + "\\" if sys.platform == "win32" else "/"
        sel_row = 0
        for i, d in enumerate(drives):
            try:
                usage = shutil.disk_usage(d)
                free = fmt_size(usage.free)
                label = f"  {d}   wolne: {free}"
            except Exception:
                label = f"  {d}"
            item = QTableWidgetItem(label)
            tbl.setItem(i, 0, item)
            if d == current_drive:
                sel_row = i

        tbl.selectRow(sel_row)
        tbl.setFocus()
        layout.addWidget(tbl)

        def accept_selection():
            row = tbl.currentRow()
            if 0 <= row < len(drives):
                panel.navigate_to(drives[row])
                self._activate_panel(panel)
            dlg.accept()

        tbl.itemActivated.connect(lambda _: accept_selection())

        # Enter akceptuje, Esc zamyka
        def key_handler(event):
            if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                accept_selection()
            elif event.key() == Qt.Key.Key_Escape:
                dlg.reject()
            else:
                QTableWidget.keyPressEvent(tbl, event)
        tbl.keyPressEvent = key_handler

        dlg.exec()

    def _err(self, e: Exception) -> str:
        """Zwróć czytelny komunikat błędu — ze wskazówką przy braku uprawnień."""
        if isinstance(e, PermissionError):
            return (
                f"Brak uprawnień dostępu.\n\n"
                f"Aby operować na folderach systemowych (np. C:\\)\n"
                f"uruchom PyCommander jako administrator:\n"
                f"kliknij prawym na .exe → 'Uruchom jako administrator'.\n\n"
                f"Szczegóły: {e}"
            )
        return str(e)

    # ── F2 Zmiana nazwy ──────────────────────────────────────────────────────
    def _rename(self):
        sel = self.active_panel.selected_item()
        if not sel or sel["name"] == "..":
            return
        dlg = InputDialog(self, "Zmień nazwę", "Nowa nazwa:", sel["name"])
        if dlg.exec() and dlg.value() and dlg.value() != sel["name"]:
            src = os.path.join(self.active_panel.current_path, sel["name"])
            dst = os.path.join(self.active_panel.current_path, dlg.value())
            try:
                os.rename(src, dst)
                self.active_panel.refresh(keep_selection=dlg.value())
                self.status.showMessage(f"Zmieniono nazwę: {sel['name']} → {dlg.value()}")
            except Exception as e:
                QMessageBox.critical(self, "Błąd", self._err(e))

    def _shell_op(self, sources: list[str], dst_dir: str, move: bool) -> list[str]:
        """
        Kopiuj lub przenieś pliki przez Windows Shell API.
        Automatycznie pokazuje okno UAC jeśli potrzebne.
        Zwraca listę błędów (pusta = sukces).
        """
        if SHELL_API:
            try:
                flags = (
                    shellcon.FOF_NOCONFIRMMKDIR |
                    shellcon.FOF_ALLOWUNDO
                )
                operation = shellcon.FO_MOVE if move else shellcon.FO_COPY
                # Shell API przyjmuje źródła oddzielone \0, cel to folder
                src_str = "\0".join(sources)
                result = win32shell.SHFileOperation((
                    0,           # hwnd
                    operation,
                    src_str,
                    dst_dir,
                    flags,
                    None,
                    None
                ))
                # result[0] == 0 oznacza sukces
                if result[0] != 0:
                    return [f"Operacja anulowana lub błąd (kod {result[0]})"]
                return []
            except Exception as e:
                return [str(e)]
        else:
            # Fallback — shutil gdy pywin32 niedostępne
            errors = []
            for src in sources:
                dst = os.path.join(dst_dir, os.path.basename(src))
                try:
                    if move:
                        shutil.move(src, dst)
                    else:
                        shutil.copytree(src, dst) if os.path.isdir(src) else shutil.copy2(src, dst)
                except Exception as e:
                    errors.append(self._err(e))
            return errors

    # ── F5 Kopiuj ────────────────────────────────────────────────────────────
    def _copy(self):
        sources = self.active_panel.marked_paths()
        if not sources:
            return
        dst_dir = self._other_panel().current_path
        names = "\n".join(os.path.basename(s) for s in sources)
        reply = QMessageBox.question(
            self, "Kopiuj",
            f"Kopiować {len(sources)} element(ów):\n{names}\n\n→ {dst_dir}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            errors = self._shell_op(sources, dst_dir, move=False)
            self.active_panel.clear_marks()
            self._other_panel().refresh()
            if errors:
                QMessageBox.critical(self, "Błędy", "\n\n".join(errors))
            else:
                self.status.showMessage(f"Skopiowano {len(sources)} element(ów)")

    # ── F6 Przesuń ───────────────────────────────────────────────────────────
    def _move(self):
        sources = self.active_panel.marked_paths()
        if not sources:
            return
        dst_dir = self._other_panel().current_path
        names = "\n".join(os.path.basename(s) for s in sources)
        reply = QMessageBox.question(
            self, "Przesuń",
            f"Przenieść {len(sources)} element(ów):\n{names}\n\n→ {dst_dir}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            errors = self._shell_op(sources, dst_dir, move=True)
            self.active_panel.clear_marks()
            self._fix_panel_path(self.active_panel)
            self._other_panel().refresh()
            if errors:
                QMessageBox.critical(self, "Błędy", "\n\n".join(errors))
            else:
                self.status.showMessage(f"Przeniesiono {len(sources)} element(ów)")

    # ── F7 Nowy folder ───────────────────────────────────────────────────────
    def _mkdir(self):
        dlg = InputDialog(self, "Nowy folder", "Nazwa nowego folderu:")
        if dlg.exec() and dlg.value():
            path = os.path.join(self.active_panel.current_path, dlg.value())
            try:
                os.makedirs(path)
                self.active_panel.refresh(keep_selection=dlg.value())
                self._other_panel().refresh()
                self.status.showMessage(f"Utworzono folder: {dlg.value()}")
            except Exception as e:
                QMessageBox.critical(self, "Błąd", self._err(e))

    # ── F8 Usuń ──────────────────────────────────────────────────────────────
    def _fix_panel_path(self, panel: FilePanel):
        path = panel.current_path
        while path and not os.path.exists(path):
            parent = os.path.dirname(path)
            if parent == path:
                path = os.path.expanduser("~")
                break
            path = parent
        panel.current_path = path
        panel.refresh()

    def _delete(self):
        sources = self.active_panel.marked_paths()
        if not sources:
            return
        names = "\n".join(os.path.basename(s) for s in sources)
        reply = QMessageBox.question(
            self, "Usuń",
            f"Usunąć {len(sources)} element(ów)?\n{names}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            errors = []
            for src in sources:
                try:
                    shutil.rmtree(src) if os.path.isdir(src) else os.remove(src)
                except Exception as e:
                    errors.append(self._err(e))
            self.active_panel.clear_marks()
            self._fix_panel_path(self.active_panel)
            self._fix_panel_path(self._other_panel())
            if errors:
                QMessageBox.critical(self, "Błędy", "\n\n".join(errors))
            else:
                self.status.showMessage(f"Usunięto {len(sources)} element(ów)")


# ─── Uruchomienie ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
