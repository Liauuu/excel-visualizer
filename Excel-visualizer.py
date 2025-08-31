import sys
import pandas as pd
import matplotlib.pyplot as plt

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog,
    QComboBox, QListWidget, QListWidgetItem, QMessageBox, QGroupBox
)
from PySide6.QtGui import QFont


# --------- Helper: robust column resolver (ignores case/whitespace/underscores) ---------
def _norm(s: str) -> str:
    return "".join(str(s).lower().replace("_", "").split())

def resolve_columns(df: pd.DataFrame, want: list[str]) -> dict:
    """Match desired normalized keys to actual df columns. Returns {want_key: real_col or None}."""
    norm_map = {_norm(c): c for c in df.columns}
    out = {}
    for w in want:
        if w in norm_map:
            out[w] = norm_map[w]
            continue
        candidates = [k for k in norm_map.keys() if w in k or k in w]
        out[w] = norm_map[candidates[0]] if candidates else None
    return out


class ExcelVizApp(QWidget):
    def __init__(self):
        super().__init__()
        # ---------- Window ----------
        self.setWindowTitle("🌈Excel Visualization👀")
        self.resize(520, 720)
        self.setMinimumWidth(460)

        # ---------- App-wide font ----------
        base_font = QFont("Segoe UI")
        base_font.setPointSize(10)
        self.setFont(base_font)

        self.df: pd.DataFrame | None = None
        self.colmap: dict[str, str] = {}   # normalized_key -> real column name
        self.loaded_path: str | None = None

        # ---------- Layout ----------
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # File area
        self.path_label = QLabel("👀:Please select an Excel file.")
        root.addWidget(self.path_label)

        open_row = QHBoxLayout()
        self.btn_open = QPushButton("🌈Open File")
        self.btn_open.clicked.connect(self.open_excel)
        open_row.addWidget(self.btn_open)

        self.btn_preview_cols = QPushButton("🌈Preview Columns")
        self.btn_preview_cols.clicked.connect(self.preview_columns)
        self.btn_preview_cols.setEnabled(False)
        open_row.addWidget(self.btn_preview_cols)
        root.addLayout(open_row)

        # ---- Chart section (pick chart type first) ----
        chart_box = QGroupBox("1) Chart type")
        chart_l = QVBoxLayout(chart_box)

        self.chart_type = QComboBox()
        self.chart_type.addItems([
            "Pie Chart (percent)",
            "Bar Chart (vertical)",
            "Bar Chart (horizontal)",
            "Line Chart (single numeric column)",
        ])
        self.chart_type.currentIndexChanged.connect(self._on_chart_type_changed)
        chart_l.addWidget(self.chart_type)

        root.addWidget(chart_box)

        # ---- Inputs section (per-chart parameters) ----
        input_box = QGroupBox("2) Chart inputs")
        input_l = QVBoxLayout(input_box)

        # Pie-only: category column
        self.pie_label = QLabel("Pie target column (Region/Product/StoreLocation/CustomerType/PaymentMethod)")
        self.pie_col = QComboBox()
        input_l.addWidget(self.pie_label)
        input_l.addWidget(self.pie_col)

        # Bar-only: allowed X,Y pairs
        self.bar_label = QLabel("Bar chart X, Y (choose from these pairs)")
        self.bar_pair = QComboBox()
        self.allowed_pairs = [
            ("Region", "Quantity"),
            ("Region", "TotalPrice"),
            ("Product", "Quantity"),
            ("Product", "TotalPrice"),
        ]
        for x, y in self.allowed_pairs:
            self.bar_pair.addItem(f"{x}  →  {y}")
        input_l.addWidget(self.bar_label)
        input_l.addWidget(self.bar_pair)

        # Line-only: single numeric column
        self.line_label = QLabel("Line chart: numeric column")
        self.line_col = QComboBox()
        input_l.addWidget(self.line_label)
        input_l.addWidget(self.line_col)

        root.addWidget(input_box)

        # ---- Metrics section (sum/max/min) ----
        metric_box = QGroupBox("3) Aggregates (sum / max / min)")
        metric_l = QVBoxLayout(metric_box)
        self.metric_sel = QComboBox()
        self.metric_sel.addItems([
            "Salesperson – TotalPrice (SUM / MAX / MIN)",
            "StoreLocation – Returned (SUM / MAX / MIN)",
        ])
        metric_l.addWidget(self.metric_sel)

        btns_row = QHBoxLayout()
        self.btn_sum = QPushButton("SUM")
        self.btn_sum.clicked.connect(lambda: self.run_metrics("sum"))
        btns_row.addWidget(self.btn_sum)
        self.btn_max = QPushButton("MAX")
        self.btn_max.clicked.connect(lambda: self.run_metrics("max"))
        btns_row.addWidget(self.btn_max)
        self.btn_min = QPushButton("MIN")
        self.btn_min.clicked.connect(lambda: self.run_metrics("min"))
        btns_row.addWidget(self.btn_min)
        metric_l.addLayout(btns_row)

        root.addWidget(metric_box)

        # ---- Run buttons ----
        run_row = QHBoxLayout()
        self.btn_draw = QPushButton("🌈Draw chart")
        self.btn_draw.clicked.connect(self.draw_chart)
        run_row.addWidget(self.btn_draw)
        root.addLayout(run_row)

        # ---- Preview / logs ----
        root.addWidget(QLabel("Log / preview"))
        self.preview = QListWidget()
        self.preview.setMinimumHeight(200)
        self.preview.setAlternatingRowColors(True)
        root.addWidget(self.preview)

        self._log("👀:Select an Excel file to begin.")

        # initial visibility
        self._on_chart_type_changed()

    # ---------- Helpers ----------
    def _log(self, msg: str):
        self.preview.addItem(QListWidgetItem(msg))
        if self.preview.count() > 120:
            self.preview.takeItem(0)
        self.preview.scrollToBottom()

    def _on_chart_type_changed(self):
        t = self.chart_type.currentText()
        # pie
        self.pie_label.setVisible("Pie" in t)
        self.pie_col.setVisible("Pie" in t)
        # bar
        is_bar = ("Bar" in t)
        self.bar_label.setVisible(is_bar)
        self.bar_pair.setVisible(is_bar)
        # line
        is_line = ("Line" in t)
        self.line_label.setVisible(is_line)
        self.line_col.setVisible(is_line)

    # ---------- Actions ----------
    def open_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "Read failed", str(e))
            return

        self.df = df
        self.loaded_path = path
        self.path_label.setText(f"Selected: {path}")
        self.btn_preview_cols.setEnabled(True)

        # expected headers (normalized keys)
        want_keys = [
            _norm("Date"),
            _norm("Region"),
            _norm("Product"),
            _norm("Quantity"),
            _norm("UnitPrice"),
            _norm("StoreLocation"),
            _norm("CustomerType"),
            _norm("Discount"),
            _norm("Salesperson"),
            _norm("TotalPrice"),
            _norm("PaymentMethod"),
            _norm("Promotion"),
            _norm("Returned"),
        ]
        self.colmap = resolve_columns(df, want_keys)

        # pie candidates
        pie_candidates = ["Region", "Product", "StoreLocation", "CustomerType", "PaymentMethod"]
        self.pie_col.clear()
        for key in pie_candidates:
            real = self.colmap.get(_norm(key))
            if real:
                self.pie_col.addItem(real)

        # line candidates (numeric)
        self.line_col.clear()
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                self.line_col.addItem(str(c))

        self._log("👀:File loaded. Choose a chart type and then the required column(s).")

    def preview_columns(self):
        if self.df is None:
            return
        cols = ", ".join(map(str, self.df.columns))
        self._log(f"Columns: {cols}")

    # ---- Metrics ----
    def run_metrics(self, mode: str):
        if self.df is None:
            QMessageBox.information(self, "No data", "Please open a file first.")
            return

        choice = self.metric_sel.currentText()
        if "Salesperson" in choice:
            gkey = self.colmap.get(_norm("Salesperson"))
            val = self.colmap.get(_norm("TotalPrice"))
            if not gkey or not val:
                QMessageBox.warning(self, "Missing columns", "Salesperson/TotalPrice not found.")
                return
            g = self.df.groupby(gkey, dropna=False)[val].sum(min_count=1)
            if mode == "sum":
                self._log("[SUM] TotalPrice by Salesperson:")
                self._log(str(g.sort_values(ascending=False).head(20)))
            elif mode == "max":
                idx = g.idxmax()
                self._log(f"[MAX] Salesperson with highest total: {idx} = {g.max()}")
            elif mode == "min":
                idx = g.idxmin()
                self._log(f"[MIN] Salesperson with lowest total: {idx} = {g.min()}")

        else:  # StoreLocation – Returned
            gkey = self.colmap.get(_norm("StoreLocation"))
            ret = self.colmap.get(_norm("Returned"))
            if not gkey or not ret:
                QMessageBox.warning(self, "Missing columns", "StoreLocation/Returned not found.")
                return
            series = self.df[ret]
            if not pd.api.types.is_numeric_dtype(series):
                series = pd.to_numeric(series, errors="coerce").fillna(0)
            g = self.df.assign(_ret=series).groupby(gkey, dropna=False)["_ret"].sum()
            if mode == "sum":
                self._log("[SUM] Returned by StoreLocation:")
                self._log(str(g.sort_values(ascending=False).head(20)))
            elif mode == "max":
                idx = g.idxmax()
                self._log(f"[MAX] StoreLocation with highest returns: {idx} = {g.max()}")
            elif mode == "min":
                idx = g.idxmin()
                self._log(f"[MIN] StoreLocation with lowest returns: {idx} = {g.min()}")

    # ---- Charts ----
    def draw_chart(self):
        if self.df is None:
            QMessageBox.information(self, "No data", "Please open a file first.")
            return

        t = self.chart_type.currentText()

        try:
            if "Pie" in t:
                cat_col = self.pie_col.currentText().strip()
                if not cat_col:
                    QMessageBox.information(self, "Select column", "Choose a pie target column.")
                    return
                counts = self.df[cat_col].astype(str).value_counts(dropna=False)
                plt.figure(figsize=(7, 7))
                plt.pie(counts.values, labels=counts.index, autopct="%1.1f%%")
                plt.title(f"Pie – {cat_col} (%)")
                plt.tight_layout(); plt.show()
                self._log("Pie chart shown.")

            elif "Bar" in t:
                pair_text = self.bar_pair.currentText()
                x_label, y_label = [s.strip() for s in pair_text.split("→")]
                x_key = self.colmap.get(_norm(x_label))
                y_key = self.colmap.get(_norm(y_label))
                if not x_key or not y_key:
                    QMessageBox.warning(self, "Missing columns", f"{x_label}/{y_label} not found.")
                    return

                y_series = self.df[y_key]
                if pd.api.types.is_numeric_dtype(y_series):
                    agg = self.df.groupby(x_key, dropna=False)[y_key].sum(min_count=1)
                else:
                    agg = self.df.groupby(x_key, dropna=False)[y_key].count()
                agg = agg.sort_values(ascending=False).head(30)

                plt.figure(figsize=(10, 6))
                if "horizontal" in t:
                    plt.barh(agg.index.astype(str), agg.values)
                    plt.xlabel(y_label); plt.ylabel(x_label)
                    plt.title(f"Bar (horizontal) – {x_label} vs {y_label}")
                    plt.gca().invert_yaxis()
                else:
                    plt.bar(agg.index.astype(str), agg.values)
                    plt.ylabel(y_label); plt.xlabel(x_label)
                    plt.title(f"Bar (vertical) – {x_label} vs {y_label}")
                    plt.xticks(rotation=45, ha="right")
                plt.tight_layout(); plt.show()
                self._log("Bar chart shown.")

            elif "Line" in t:
                col = self.line_col.currentText().strip()
                if not col:
                    QMessageBox.information(self, "Select column", "Choose a numeric column.")
                    return
                series = pd.to_numeric(self.df[col], errors="coerce").dropna()
                plt.figure(figsize=(10, 5))
                plt.plot(series.index, series.values, marker="o")
                plt.title(f"Line – {col}")
                plt.xlabel("Row Index"); plt.ylabel(col); plt.grid(True)
                plt.tight_layout(); plt.show()
                self._log("Line chart shown.")

            else:
                QMessageBox.information(self, "Not Implemented", f"{t} is not implemented yet.")
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", str(e))
            self._log(f"Plot error: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelVizApp()
    window.show()
    sys.exit(app.exec())
