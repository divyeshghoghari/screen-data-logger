import sys, os, re, time, json
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QRect, Signal
from PySide6.QtGui import QPainter, QColor, QPen, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QDoubleSpinBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QFormLayout, QHeaderView, QAbstractItemView
)

try:
    import pytesseract
    from PIL import ImageGrab, ImageOps, ImageEnhance, ImageFilter
    _tesseract_cmd = os.environ.get("TESSERACT_CMD")
    if not _tesseract_cmd:
        candidate = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(candidate):
            _tesseract_cmd = candidate
    if _tesseract_cmd and os.path.exists(_tesseract_cmd):
        pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd
    OCR_IMPORT_ERROR = None
except Exception as exc:
    pytesseract = None
    ImageGrab = ImageOps = ImageEnhance = ImageFilter = None
    OCR_IMPORT_ERROR = str(exc)

APP_NAME = "Screen Data Logger"
CONFIG_PATH = Path.home() / ".screen_data_logger_config.json"
OCR_WHITELIST = "0123456789.,+-"
PSM_SEQUENCE = (7, 8, 13, 6, 7, 8, 13, 6, 7, 8)

class RegionOverlay(QWidget):
    selected = Signal(tuple)
    cancelled = Signal()

    def __init__(self, screen_geometry):
        super().__init__()
        self.setGeometry(screen_geometry)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setCursor(Qt.CursorShape.CrossCursor)
        self.origin = None
        self.current = None
        self.show()
        self.raise_()
        self.activateWindow()

    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, 100))
        if self.origin and self.current:
            r = QRect(self.origin, self.current).normalized()
            p.fillRect(r, QColor(40, 170, 255, 65))
            p.setPen(QPen(QColor(30, 180, 255), 2))
            p.drawRect(r)
        p.setPen(QColor(255, 255, 255))
        p.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        p.drawText(24, 42, "Drag around the displayed number. Press Esc to cancel.")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.origin = event.position().toPoint()
            self.current = self.origin
            self.update()

    def mouseMoveEvent(self, event):
        if self.origin:
            self.current = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event):
        if self.origin:
            self.current = event.position().toPoint()
            r = QRect(self.origin, self.current).normalized()
            if r.width() >= 5 and r.height() >= 5:
                g = self.geometry()
                self.selected.emit((g.x() + r.x(), g.y() + r.y(), r.width(), r.height()))
            else:
                self.cancelled.emit()
            self.close()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.cancelled.emit()
            self.close()


class ScreenLogger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1180, 760)
        self.regions = []
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.log_once)
        self.logging = False
        self.output_path = str(Path.home() / "Documents" / "screen_log.xlsx")
        self.workbook = None
        self.worksheet = None
        self.next_row = 2
        self.overlay = None
        self._build_ui()
        self._load_config()
        self._refresh_table()
        self._update_status("Ready. Add parameters and select a screen rectangle for each.")

    def _build_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)
        intro = QLabel(
            "<h2>Screen Data Logger</h2>"
            "Reads numeric values from selected SCADA screen areas and records them in Excel."
        )
        intro.setWordWrap(True)
        root.addWidget(intro)

        cfg = QFormLayout()
        self.interval = QDoubleSpinBox()
        self.interval.setRange(0.5, 86400)
        self.interval.setDecimals(1)
        self.interval.setSingleStep(1)
        self.interval.setValue(30)
        self.interval.setSuffix(" seconds")
        cfg.addRow("Logging interval:", self.interval)

        self.retry_count = QSpinBox()
        self.retry_count.setRange(0, 10)
        self.retry_count.setValue(5)
        self.retry_count.setSuffix(" retries after the first attempt")
        cfg.addRow("OCR retries:", self.retry_count)

        pathrow = QHBoxLayout()
        self.path_edit = QLineEdit(self.output_path)
        browse = QPushButton("Choose Excel file…")
        browse.clicked.connect(self.choose_output)
        pathrow.addWidget(self.path_edit, 1)
        pathrow.addWidget(browse)
        cfg.addRow("Output file:", pathrow)
        root.addLayout(cfg)

        actions = QHBoxLayout()
        self.add_btn = QPushButton("＋ Add parameter")
        self.add_btn.clicked.connect(self.add_parameter)
        self.select_btn = QPushButton("Select rectangle for selected row")
        self.select_btn.clicked.connect(self.select_region)
        self.test_btn = QPushButton("Test OCR (with retries)")
        self.test_btn.clicked.connect(self.test_ocr)
        self.remove_btn = QPushButton("Remove selected")
        self.remove_btn.clicked.connect(self.remove_parameter)
        for btn in (self.add_btn, self.select_btn, self.test_btn, self.remove_btn):
            actions.addWidget(btn)
        root.addLayout(actions)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "Enabled", "Parameter name", "Unit", "Screen rectangle (x, y, w, h)",
            "Latest value", "OCR status", "Decimals"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        for col, mode in enumerate([
            QHeaderView.ResizeMode.ResizeToContents, QHeaderView.ResizeMode.Stretch,
            QHeaderView.ResizeMode.ResizeToContents, QHeaderView.ResizeMode.Stretch,
            QHeaderView.ResizeMode.ResizeToContents, QHeaderView.ResizeMode.Stretch,
            QHeaderView.ResizeMode.ResizeToContents
        ]):
            self.table.horizontalHeader().setSectionResizeMode(col, mode)
        root.addWidget(self.table, 1)

        bottom = QHBoxLayout()
        self.start_btn = QPushButton("▶ Start logging")
        self.start_btn.clicked.connect(self.start_logging)
        self.stop_btn = QPushButton("■ Stop")
        self.stop_btn.clicked.connect(self.stop_logging)
        self.stop_btn.setEnabled(False)
        self.status = QLabel("Status: Ready")
        self.status.setWordWrap(True)
        bottom.addWidget(self.start_btn)
        bottom.addWidget(self.stop_btn)
        bottom.addWidget(self.status, 1)
        root.addLayout(bottom)

        note = QLabel(
            "Parameters are processed sequentially. Each parameter gets up to the configured number "
            "of OCR attempts. Failed values are left blank and marked OCR FAILED; old values are never reused."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #555;")
        root.addWidget(note)
        self.setCentralWidget(central)

    def _update_status(self, msg):
        self.status.setText("Status: " + msg)

    def _selected_row(self):
        row = self.table.currentRow()
        return row if row >= 0 else None

    def add_parameter(self):
        self.regions.append({
            "enabled": True, "name": f"Parameter {len(self.regions)+1}", "unit": "",
            "rect": None, "latest": "", "status": "Not tested", "decimals": 2
        })
        self._refresh_table()
        self.table.selectRow(len(self.regions) - 1)
        self._save_config()

    def remove_parameter(self):
        row = self._selected_row()
        if row is None:
            QMessageBox.information(self, APP_NAME, "Select a parameter row first.")
            return
        del self.regions[row]
        self._refresh_table()
        self._save_config()

    def _refresh_table(self):
        self.table.setRowCount(len(self.regions))
        for i, item in enumerate(self.regions):
            enabled = QTableWidgetItem()
            enabled.setFlags(enabled.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            enabled.setCheckState(Qt.CheckState.Checked if item.get("enabled", True) else Qt.CheckState.Unchecked)
            self.table.setItem(i, 0, enabled)
            self.table.setItem(i, 1, QTableWidgetItem(item.get("name", "")))
            self.table.setItem(i, 2, QTableWidgetItem(item.get("unit", "")))
            rect = item.get("rect")
            self.table.setItem(i, 3, QTableWidgetItem(", ".join(map(str, rect)) if rect else "Not selected"))
            self.table.setItem(i, 4, QTableWidgetItem(item.get("latest", "")))
            self.table.setItem(i, 5, QTableWidgetItem(item.get("status", "Not tested")))
            self.table.setItem(i, 6, QTableWidgetItem(str(item.get("decimals", 2))))
        self.table.resizeRowsToContents()

    def _sync_from_table(self):
        for i, item in enumerate(self.regions):
            item["enabled"] = self.table.item(i, 0).checkState() == Qt.CheckState.Checked
            item["name"] = self.table.item(i, 1).text().strip() or f"Parameter {i+1}"
            item["unit"] = self.table.item(i, 2).text().strip()
            try:
                item["decimals"] = max(0, min(8, int(self.table.item(i, 6).text())))
            except Exception:
                item["decimals"] = 2
        self._save_config()

    def select_region(self):
        row = self._selected_row()
        if row is None:
            QMessageBox.information(self, APP_NAME, "Add and select a parameter row first.")
            return
        self._sync_from_table()
        self.hide()
        QApplication.processEvents()
        screen = QApplication.primaryScreen()
        if not screen:
            self.show()
            QMessageBox.warning(self, APP_NAME, "No screen detected.")
            return
        self.overlay = RegionOverlay(screen.geometry())
        self.overlay.selected.connect(lambda rect, r=row: self._region_selected(r, rect))
        self.overlay.cancelled.connect(self._region_cancelled)

    def _region_selected(self, row, rect):
        self.show()
        self.raise_()
        self.activateWindow()
        self.regions[row]["rect"] = list(rect)
        self.regions[row]["status"] = "Region selected; test OCR"
        self._refresh_table()
        self.table.selectRow(row)
        self._save_config()
        self._update_status(f"Selected rectangle for {self.regions[row]['name']}.")

    def _region_cancelled(self):
        self.show()
        self._update_status("Rectangle selection cancelled.")

    def _make_ocr_image(self, img, attempt):
        # Different retry attempts deliberately use different preprocessing strategies.
        scale = 4 if attempt in (0, 2, 4, 6, 8) else 3
        img = img.resize((max(1, img.width * scale), max(1, img.height * scale)))
        gray = ImageOps.grayscale(img)
        if attempt % 5 == 0:
            return ImageEnhance.Contrast(gray).enhance(2.4).filter(ImageFilter.SHARPEN)
        if attempt % 5 == 1:
            return gray.point(lambda p: 255 if p > 145 else 0)
        if attempt % 5 == 2:
            return ImageOps.invert(ImageEnhance.Contrast(gray).enhance(2.0)).point(lambda p: 255 if p > 130 else 0)
        if attempt % 5 == 3:
            return ImageEnhance.Sharpness(ImageEnhance.Contrast(gray).enhance(3.0)).enhance(2.0)
        return gray.point(lambda p: 255 if p > 185 else 0)

    def _capture_text(self, rect, attempt):
        if not pytesseract or not ImageGrab:
            raise RuntimeError("OCR dependencies missing. Install requirements and Tesseract OCR; see README.")
        x, y, w, h = rect
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
        processed = self._make_ocr_image(img, attempt)
        psm = PSM_SEQUENCE[min(attempt, len(PSM_SEQUENCE) - 1)]
        config = f"--oem 3 --psm {psm} -c tessedit_char_whitelist={OCR_WHITELIST}"
        return pytesseract.image_to_string(processed, config=config).strip()

    def _parse_number(self, raw):
        cleaned = raw.replace(" ", "").replace("\n", "")
        match = re.search(r"[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)", cleaned)
        if not match:
            return None
        token = match.group(0)
        if "," in token and "." not in token:
            after = token.rsplit(",", 1)[-1]
            token = token.replace(",", ".") if len(after) in (1, 2) else token.replace(",", "")
        else:
            token = token.replace(",", "")
        try:
            value = float(token)
            if value != value or value in (float("inf"), float("-inf")):
                return None
            return value
        except ValueError:
            return None

    def _read_parameter(self, item):
        if not item.get("rect"):
            return None, "No region"
        attempts = self.retry_count.value() + 1
        errors = []
        last_raw = ""
        for attempt in range(attempts):
            try:
                raw = self._capture_text(item["rect"], attempt)
                last_raw = raw
                value = self._parse_number(raw)
                if value is not None:
                    if attempt == 0:
                        return value, "OK (1 attempt)"
                    return value, f"OK (retry {attempt + 1}/{attempts})"
                errors.append(f"Attempt {attempt + 1}: no number")
            except Exception as exc:
                errors.append(f"Attempt {attempt + 1}: {str(exc)[:90]}")
            if attempt < attempts - 1:
                time.sleep(0.20)
        detail = last_raw if last_raw else (errors[-1] if errors else "no readable text")
        return None, f"OCR FAILED ({attempts} attempts): {detail[:70]}"

    def test_ocr(self):
        row = self._selected_row()
        if row is None:
            QMessageBox.information(self, APP_NAME, "Select a parameter row first.")
            return
        self._sync_from_table()
        item = self.regions[row]
        try:
            value, status = self._read_parameter(item)
            item["latest"] = "" if value is None else f"{value:.{item.get('decimals', 2)}f}"
            item["status"] = status
            self._refresh_table()
            self.table.selectRow(row)
            self._save_config()
            self._update_status(f"{item['name']}: {item['latest'] or '—'} ({status})")
        except Exception as exc:
            QMessageBox.critical(self, APP_NAME, str(exc))
            self._update_status("OCR test failed.")

    def choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Choose Excel output", self.path_edit.text(), "Excel Workbook (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.path_edit.setText(path)
            self.output_path = path
            self._save_config()

    def start_logging(self):
        self._sync_from_table()
        enabled = [x for x in self.regions if x.get("enabled", True)]
        if not enabled:
            QMessageBox.warning(self, APP_NAME, "Add and enable at least one parameter.")
            return
        missing = [x["name"] for x in enabled if not x.get("rect")]
        if missing:
            QMessageBox.warning(self, APP_NAME, "Select a rectangle for: " + ", ".join(missing))
            return
        self.output_path = self.path_edit.text().strip()
        if not self.output_path:
            QMessageBox.warning(self, APP_NAME, "Choose an Excel output file.")
            return
        if not self.output_path.lower().endswith(".xlsx"):
            self.output_path += ".xlsx"
            self.path_edit.setText(self.output_path)
        try:
            from openpyxl import Workbook, load_workbook
            Path(self.output_path).parent.mkdir(parents=True, exist_ok=True)
            if os.path.exists(self.output_path):
                self.workbook = load_workbook(self.output_path)
                self.worksheet = self.workbook["Screen Log"] if "Screen Log" in self.workbook.sheetnames else self.workbook.active
                if self.worksheet.max_row < 1 or self.worksheet.cell(1, 1).value is None:
                    self._write_headers()
                else:
                    self.next_row = self.worksheet.max_row + 1
            else:
                self.workbook = Workbook()
                self.worksheet = self.workbook.active
                self.worksheet.title = "Screen Log"
                self._write_headers()
            self.workbook.save(self.output_path)
        except Exception as exc:
            QMessageBox.critical(self, APP_NAME, f"Could not open/create Excel file. Close it in Excel if it is already open.\n\n{exc}")
            return
        self.logging = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self._update_status(f"Logging started. {len(enabled)} parameters are read sequentially; up to {self.retry_count.value() + 1} total OCR attempts each.")
        # Take first sample immediately, then continue at the selected interval.
        self.log_once()
        if self.logging:
            self.timer.start(max(500, int(self.interval.value() * 1000)))

    def _write_headers(self):
        enabled = [x for x in self.regions if x.get("enabled", True)]
        self.worksheet.cell(1, 1, "Timestamp")
        for col, item in enumerate(enabled, start=2):
            self.worksheet.cell(1, col, f"{item['name']}" + (f" ({item['unit']})" if item["unit"] else ""))
        offset = 2 + len(enabled)
        for col, item in enumerate(enabled, start=offset):
            self.worksheet.cell(1, col, f"{item['name']} OCR Status")
        self.next_row = 2

    def log_once(self):
        if not self.logging:
            return
        self._sync_from_table()
        active = [x for x in self.regions if x.get("enabled", True)]
        timestamp = datetime.now().astimezone().isoformat(timespec="milliseconds")
        values, statuses = [], []
        for item in active:
            try:
                value, status = self._read_parameter(item)
                values.append(value)
                statuses.append(status)
                item["latest"] = "" if value is None else f"{value:.{item.get('decimals', 2)}f}"
                item["status"] = status
            except Exception as exc:
                values.append(None)
                statuses.append(f"OCR ERROR: {str(exc)[:70]}")
                item["latest"] = ""
                item["status"] = statuses[-1]
        self._refresh_table()
        self.worksheet.cell(self.next_row, 1, timestamp)
        for col, value in enumerate(values, start=2):
            self.worksheet.cell(self.next_row, col, value)
        for col, status in enumerate(statuses, start=2 + len(active)):
            self.worksheet.cell(self.next_row, col, status)
        try:
            self.workbook.save(self.output_path)
            self.next_row += 1
            failed = sum(v is None for v in values)
            self._update_status(
                f"Logged {timestamp} | {len(active)} parameters | {failed} failed | {Path(self.output_path).name}"
            )
        except Exception as exc:
            self.stop_logging()
            QMessageBox.critical(self, APP_NAME, f"Saving failed; logging stopped. Close the workbook in Excel and restart.\n{exc}")

    def stop_logging(self):
        self.logging = False
        self.timer.stop()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        if self.workbook:
            try:
                self.workbook.save(self.output_path)
            except Exception:
                pass
        self._update_status("Stopped. Data saved where possible.")

    def _save_config(self):
        try:
            self.output_path = self.path_edit.text().strip() if hasattr(self, "path_edit") else self.output_path
            data = {
                "interval": self.interval.value() if hasattr(self, "interval") else 30,
                "retry_count": self.retry_count.value() if hasattr(self, "retry_count") else 5,
                "output_path": self.output_path,
                "regions": self.regions
            }
            CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_config(self):
        try:
            if CONFIG_PATH.exists():
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                self.regions = data.get("regions", [])
                self.interval.setValue(float(data.get("interval", 30)))
                self.retry_count.setValue(int(data.get("retry_count", 5)))
                self.output_path = data.get("output_path", self.output_path)
                self.path_edit.setText(self.output_path)
        except Exception:
            self.regions = []

    def closeEvent(self, event):
        if self.logging:
            reply = QMessageBox.question(self, APP_NAME, "Logging is active. Stop logging and exit?")
            if reply != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.stop_logging()
        self._save_config()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = ScreenLogger()
    win.show()
    if OCR_IMPORT_ERROR:
        QMessageBox.information(
            win, APP_NAME,
            "OCR dependencies may be missing. Install packages from requirements.txt and install Tesseract OCR. See README.md."
        )
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
