import sys
import pandas as pd
import matplotlib.pyplot as plt

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QComboBox, QListWidget, QListWidgetItem, QMessageBox
)
from PySide6.QtGui import QFont

class ExcelVizApp(QWidget):
    def __init__(self):
        super().__init__()
        # ---------- Window ----------
        self.setWindowTitle("Excel Data Visualization Tool👀")
        # Slim & tall, but a bit shorter; preview will be smaller
        self.resize(420, 640)
        self.setMinimumWidth(380)

        # ---------- App-wide cute font + size ----------
        # Use a friendly font if available; fallback to Segoe UI/Arial
        base_font = QFont("Comic Sans MS")  # Windows usually has this
        if not base_font.exactMatch():
            base_font = QFont("Segoe UI")
        base_font.setPointSize(10)
        self.setFont(base_font)

        # Extra styling (bigger section titles, centered)
        self.setStyleSheet("""
            QLabel#SectionTitle {
                font-weight: 700;
                font-size: 14px;
                qproperty-alignment: 'AlignHCenter';
            }
            QLabel#PathLabel {
                font-size: 10.5px;
            }
            QComboBox, QPushButton {
                font-size: 11px;
            }
            QListWidget {
                font-size: 10.5px;
            }
        """)

        # ---------- State ----------
        self.df: pd.DataFrame | None = None
        self.loaded_path: str | None = None

        # ---------- Layout ----------
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # File path label
        self.path_label = QLabel("👀:Please choose an Excel file.")
        self.path_label.setObjectName("PathLabel")
        self.path_label.setWordWrap(True)
        root.addWidget(self.path_label)

        # Open file button
        self.btn_open = QPushButton("🌈Open Excel File")
        self.btn_open.clicked.connect(self.open_excel)
        root.addWidget(self.btn_open)

        # Section: column
        root.addWidget(self._make_section_label("🌈Select a column"))
        self.col_box = QComboBox()
        self._set_combobox_placeholder(self.col_box, "Select column…")
        root.addWidget(self.col_box)

        # Section: chart type
        root.addWidget(self._make_section_label("🌈Choose a chart type"))
        self.chart_box = QComboBox()
        self._set_combobox_placeholder(self.chart_box, "Select chart type…")
        self.chart_box.addItems([
            "Find Max",
            "Find Min",
            "Histogram (vertical bars)",
            "Bar Chart (horizontal)",
            "Line Chart",
            "Pie Chart"
        ])
        root.addWidget(self.chart_box)

        # Preview button
        self.btn_run = QPushButton("Preview Visualization")
        self.btn_run.clicked.connect(self.run_viz)
        root.addWidget(self.btn_run)

        # Preview area (smaller height)
        root.addWidget(self._make_section_label("Preview"))
        self.preview = QListWidget()
        # Reduced height as requested
        self.preview.setMinimumHeight(160)
        self.preview.setMaximumHeight(220)
        self.preview.setAlternatingRowColors(True)
        root.addWidget(self.preview)

        # Initial guidance
        self._log("👀:Please select an Excel file.")
        self._log("After selecting a file, choose a column (or row) to visualize.")
        self._log("Then choose a chart type and click 'Preview Visualization'.")

    # ---------- Helpers ----------
    def _make_section_label(self, text: str) -> QLabel:
        lab = QLabel(text)
        lab.setObjectName("SectionTitle")
        lab.setAlignment(Qt.AlignHCenter)  # center
        return lab

    def _set_combobox_placeholder(self, box: QComboBox, placeholder: str) -> None:
        # Editable only to show placeholder; keep read-only so users can't type new items
        box.setEditable(True)
        box.setInsertPolicy(QComboBox.NoInsert)
        box.lineEdit().setPlaceholderText(placeholder)
        box.lineEdit().setReadOnly(True)
        box.setFocusPolicy(Qt.StrongFocus)

    def _log(self, msg: str) -> None:
        """Append a line to preview, keeping it short (max ~60 lines)."""
        self.preview.addItem(QListWidgetItem(msg))
        # Keep preview length modest
        if self.preview.count() > 60:
            self.preview.takeItem(0)
        self.preview.scrollToBottom()

    # ---------- Actions ----------
    def open_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)  # needs openpyxl for .xlsx
        except Exception as e:
            QMessageBox.critical(self, "👀:Failed to read Excel", str(e))
            return

        self.df = df
        self.loaded_path = path
        self.path_label.setText(f"Selected: {path}")

        # Populate columns
        self.col_box.clear()
        self._set_combobox_placeholder(self.col_box, "Select column…")
        self.col_box.addItems([str(c) for c in df.columns])

        # Update preview guidance
        self.preview.clear()
        self._log("👀:File loaded successfully")
        self._log("Please select a column (or row) to visualize.")
        self._log("Next, choose a chart type.")

    def run_viz(self):
        if self.df is None:
            QMessageBox.information(self, "No Data", "Please select a file first.")
            self._log("Please select an Excel file.")
            return

        col = self.col_box.currentText().strip()
        if not col:
            QMessageBox.information(self, "Select Column", "Please select a column.")
            self._log("Please select a column first.")
            return
        if col not in self.df.columns:
            QMessageBox.warning(self, "Invalid Column", f"Column '{col}' not found.")
            return

        chart = self.chart_box.currentText().strip()
        if not chart:
            QMessageBox.information(self, "Select Chart", "Please choose a chart type.")
            self._log("Please choose a chart type.")
            return

        series = self.df[col].dropna()
        is_numeric = pd.api.types.is_numeric_dtype(series)

        # Text-only results
        if chart == "Find Max":
            try:
                self._log(f"[Max] {col} = {series.max()}")
            except Exception:
                self._log(f"[Max] Unable to compute max for '{col}'.")
            return

        if chart == "Find Min":
            try:
                self._log(f"[Min] {col} = {series.min()}")
            except Exception:
                self._log(f"[Min] Unable to compute min for '{col}'.")
            return

        # Plot branches
        try:
            if chart == "Histogram (vertical bars)":
                if not is_numeric:
                    self._log(f"'{col}' is not numeric. Converting to numeric where possible.")
                    series = pd.to_numeric(series, errors="coerce").dropna()
                plt.figure(figsize=(8, 5))
                plt.hist(series, bins=20)
                plt.title(f"Histogram of {col}")
                plt.xlabel(col); plt.ylabel("Frequency"); plt.grid(True)
                plt.tight_layout(); plt.show()
                self._log("Histogram shown.")

            elif chart == "Bar Chart (horizontal)":
                if is_numeric:
                    data = series.nlargest(30)
                    y = data.values
                    x = data.index.astype(str)
                    plt.figure(figsize=(8, 8))
                    plt.barh(x, y)
                    plt.title(f"Top 30 Largest Values in {col}")
                    plt.xlabel(col); plt.ylabel("Row Index")
                    plt.gca().invert_yaxis()
                else:
                    vc = series.astype(str).value_counts().head(20)
                    plt.figure(figsize=(10, 8))
                    plt.barh(vc.index[::-1], vc.values[::-1])
                    plt.title(f"Top 20 Categories in {col}")
                    plt.xlabel("Count"); plt.ylabel(col)
                plt.tight_layout(); plt.grid(False); plt.show()
                self._log("Horizontal bar chart shown.")

            elif chart == "Line Chart":
                if not is_numeric:
                    self._log(f"'{col}' is not numeric. Converting to numeric where possible.")
                    series = pd.to_numeric(series, errors="coerce").dropna()
                plt.figure(figsize=(9, 5))
                plt.plot(series.index, series.values, marker="o")
                plt.title(f"Line Chart of {col}")
                plt.xlabel("Row Index"); plt.ylabel(col); plt.grid(True)
                plt.tight_layout(); plt.show()
                self._log("Line chart shown.")

            elif chart == "Pie Chart":
                if is_numeric:
                    cats = pd.cut(series, bins=5).astype(str)
                    counts = cats.value_counts().sort_index()
                else:
                    counts = series.astype(str).value_counts().head(10)
                plt.figure(figsize=(7, 7))
                plt.pie(counts.values, labels=counts.index, autopct="%1.1f%%")
                plt.title(f"Pie Chart of {col}")
                plt.tight_layout(); plt.show()
                self._log("Pie chart shown.")

            else:
                QMessageBox.information(self, "Not Implemented", f"'{chart}' is not implemented.")
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", str(e))
            self._log(f"Plot error: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelVizApp()
    window.show()
    sys.exit(app.exec())
