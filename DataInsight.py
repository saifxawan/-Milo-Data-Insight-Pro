"""
Data Insight Pro v2 (Milo)
Improved UI, Dashboard, Preview, Plots and robust file handling.
Renamed & themed as "📊 Milo – Data Insight Pro" with emoji-enhanced sidebar.

Save as DataInsight_Milo.py and run:
python DataInsight_Milo.py
"""

import os
import tempfile
import threading
from tkinter import filedialog, ttk, messagebox
import tkinter as tk
import customtkinter as ctk
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from fpdf import FPDF

# ---------- Appearance ----------
# set a modern dark theme by default; user may change via the Appearance menu
ctk.set_appearance_mode("Dark")         # Options: "System", "Light", "Dark"
ctk.set_default_color_theme("dark-blue")  # built-in: "blue", "green", "dark-blue"

APP_TITLE = "📊 Milo – Data Insight Pro"

# ---------- Helper functions ----------
def try_read_data(path):
    """Try common separators and formats to read a file robustly."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    if ext == ".json":
        return pd.read_json(path)
    # For text-based files, try multiple separators
    for sep in [",", ";", "\t", "|"]:
        try:
            df = pd.read_csv(path, sep=sep, engine="python")
            # simple heuristic: must have more than 1 column
            if df.shape[1] > 1:
                return df
        except Exception:
            continue
    # fallback to plain read_csv (may raise)
    return pd.read_csv(path)


# ---------- App ----------
class DataInsightPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window
        self.title(APP_TITLE)
        self.geometry("1200x760")
        self.minsize(1000, 680)
        # subtle background color for a premium look
        try:
            self.configure(fg_color="#121212")
        except Exception:
            # some older CTk versions may not support configure on root in same way; ignore safely
            pass

        # Data
        self.df = None
        self.current_file = None

        # Layout frames
        self.sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")

        self.main = ctk.CTkFrame(self)
        self.main.pack(side="right", expand=True, fill="both")

        # build UI
        self._build_sidebar()
        self._build_main()

    # ---------- appearance handler ----------
    def change_appearance(self, new_mode):
        """Change app appearance mode (System, Light, Dark)."""
        try:
            if new_mode in ("System", "Light", "Dark"):
                ctk.set_appearance_mode(new_mode)
            else:
                ctk.set_appearance_mode(new_mode)
        except Exception:
            pass

    # ---------- sidebar ----------
    def _build_sidebar(self):
        ctk.CTkLabel(self.sidebar, text=APP_TITLE, font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(18, 12))

        # File controls (with emojis)
        ctk.CTkButton(self.sidebar, text="📂  Load Data", command=self.load_file).pack(padx=20, pady=8, fill="x")
        ctk.CTkButton(self.sidebar, text="🔁  Reload Last", command=self.reload_last).pack(padx=20, pady=6, fill="x")
        ctk.CTkButton(self.sidebar, text="🔍  Preview / Refresh", command=self.preview_data).pack(padx=20, pady=6, fill="x")

        # separator
        ctk.CTkFrame(self.sidebar, height=1, corner_radius=1, fg_color="#2b2b2b").pack(fill="x", padx=16, pady=(12,12))

        # Cleaning
        ctk.CTkLabel(self.sidebar, text="🧹  Data Cleaning").pack(anchor="w", padx=20)
        ctk.CTkButton(self.sidebar, text="✨  Clean Data Options", command=self.clean_data).pack(padx=20, pady=6, fill="x")

        ctk.CTkFrame(self.sidebar, height=1, corner_radius=1, fg_color="#2b2b2b").pack(fill="x", padx=16, pady=(12,12))

        # EDA / Plots
        ctk.CTkLabel(self.sidebar, text="📈  Exploratory Analysis").pack(anchor="w", padx=20)
        ctk.CTkButton(self.sidebar, text="🧾  Dashboard / Summary", command=self.show_dashboard).pack(padx=20, pady=6, fill="x")
        ctk.CTkButton(self.sidebar, text="🧭  Correlation Heatmap", command=self.show_correlation).pack(padx=20, pady=6, fill="x")

        ctk.CTkLabel(self.sidebar, text="📊  Visualizations").pack(anchor="w", padx=20, pady=(8,0))
        ctk.CTkButton(self.sidebar, text="🔢  Plot: Column vs Column", command=self.plot_two_columns).pack(padx=20, pady=6, fill="x")
        ctk.CTkButton(self.sidebar, text="📚  Histogram (Column)", command=self.plot_histogram).pack(padx=20, pady=6, fill="x")

        ctk.CTkFrame(self.sidebar, height=1, corner_radius=1, fg_color="#2b2b2b").pack(fill="x", padx=16, pady=(12,12))

        # Export
        ctk.CTkLabel(self.sidebar, text="📦  Export").pack(anchor="w", padx=20)
        ctk.CTkButton(self.sidebar, text="💾  Export Cleaned CSV", command=self.export_csv).pack(padx=20, pady=6, fill="x")
        ctk.CTkButton(self.sidebar, text="📄  Export PDF Report", command=self.export_pdf_report).pack(padx=20, pady=6, fill="x")

        ctk.CTkLabel(self.sidebar, text="🎨  Appearance").pack(anchor="w", padx=20, pady=(12,0))
        self.appearance_option = ctk.CTkOptionMenu(self.sidebar, values=["System","Light","Dark"], command=self.change_appearance)
        self.appearance_option.set("Dark")
        self.appearance_option.pack(padx=20, pady=6, fill="x")

        ctk.CTkLabel(self.sidebar, text="Status:").pack(anchor="w", padx=20, pady=(12,2))
        self.status_label = ctk.CTkLabel(self.sidebar, text="No file loaded", anchor="w")
        self.status_label.pack(padx=20, pady=(0,10), fill="x")

    # ---------- main (tabs) ----------
    def _build_main(self):
        self.tabview = ctk.CTkTabview(self.main)
        self.tabview.pack(expand=True, fill="both", padx=12, pady=12)
        self.tabview.add("Preview")
        self.tabview.add("Dashboard")
        self.tabview.add("Plot")

        # Preview tab
        self.preview_frame = self.tabview.tab("Preview")
        self._build_preview()

        # Dashboard tab
        self.dashboard_frame = self.tabview.tab("Dashboard")
        self._build_dashboard()

        # Plot tab
        self.plot_frame = self.tabview.tab("Plot")
        self._build_plot_area()

    # ---------- Preview ----------
    def _build_preview(self):
        top = ctk.CTkFrame(self.preview_frame)
        top.pack(fill="x", padx=10, pady=(8,6))

        ctk.CTkLabel(top, text="Preview & Search:").pack(side="left", padx=(6,8))
        self.search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(top, textvariable=self.search_var, placeholder_text="type to filter rows (substring across all columns)")
        search_entry.pack(side="left", padx=(0,8), fill="x", expand=True)
        search_entry.bind("<KeyRelease>", lambda e: self.preview_data())

        btn_frame = ctk.CTkFrame(top, fg_color="transparent")
        btn_frame.pack(side="right")
        ctk.CTkButton(btn_frame, text="Show Head", width=90, command=lambda: self.preview_data(head_only=True)).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Show All (200 rows)", width=130, command=lambda: self.preview_data(head_only=False)).pack(side="left", padx=6)

        self.table_container = ctk.CTkFrame(self.preview_frame)
        self.table_container.pack(expand=True, fill="both", padx=10, pady=(6,10))

        self.empty_label = ctk.CTkLabel(self.table_container, text="Load a dataset to preview", anchor="center")
        self.empty_label.pack(expand=True)

    def preview_data(self, head_only=True):
        for w in self.table_container.winfo_children():
            w.destroy()

        if self.df is None or self.df.empty:
            self.empty_label = ctk.CTkLabel(self.table_container, text="Load a dataset to preview", anchor="center")
            self.empty_label.pack(expand=True)
            return

        df = self.df.copy()
        q = self.search_var.get().strip().lower()
        if q:
            try:
                mask = df.apply(lambda row: row.astype(str).str.lower().str.contains(q).any(), axis=1)
                df = df[mask]
            except Exception:
                pass

        display_rows = 200 if not head_only else min(50, len(df))
        if display_rows == 0:
            self.empty_label = ctk.CTkLabel(self.table_container, text="No rows to display after filtering", anchor="center")
            self.empty_label.pack(expand=True)
            return
        display_df = df.head(display_rows)

        container = ttk.Frame(self.table_container)
        container.pack(expand=True, fill="both")

        cols = list(display_df.columns)
        if not cols:
            self.empty_label = ctk.CTkLabel(self.table_container, text="No columns to display", anchor="center")
            self.empty_label.pack(expand=True)
            return

        tree = ttk.Treeview(container, columns=cols, show="headings", height=20)
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(expand=True, fill="both")

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor="w")

        for _, row in display_df.iterrows():
            vals = [self._format_value(row[c]) for c in cols]
            tree.insert("", "end", values=vals)

    # ---------- Dashboard ----------
    def _build_dashboard(self):
        top = ctk.CTkFrame(self.dashboard_frame)
        top.pack(fill="x", padx=10, pady=10)

        self.card_frames = []
        for i in range(4):
            card = ctk.CTkFrame(top, width=1)
            card.pack(side="left", expand=True, fill="both", padx=8)
            self.card_frames.append(card)

        self.summary_text = ctk.CTkTextbox(self.dashboard_frame, height=220, wrap="word")
        self.summary_text.pack(fill="both", padx=10, pady=(6,12), expand=False)

        self.dashboard_plot_container = ctk.CTkFrame(self.dashboard_frame)
        self.dashboard_plot_container.pack(expand=True, fill="both", padx=10, pady=(0,10))

    def show_dashboard(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Please load a file first.")
            return

        rows, cols = self.df.shape
        missing_pct = round(self.df.isna().sum().sum() / (rows * cols) * 100, 2) if rows * cols > 0 else 0
        numeric_count = len(self.df.select_dtypes(include=[np.number]).columns)

        card_values = [
            ("Rows", str(rows)),
            ("Columns", str(cols)),
            ("Missing %", f"{missing_pct}%"),
            ("Numeric cols", str(numeric_count)),
        ]
        for frame, (title, value) in zip(self.card_frames, card_values):
            for w in frame.winfo_children():
                w.destroy()
            ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=12)).pack(pady=(14,6))
            ctk.CTkLabel(frame, text=value, font=ctk.CTkFont(size=20, weight="bold")).pack()

        buf = []
        buf.append("=== Top 10 rows preview ===\n")
        buf.append(self.df.head(10).to_string())
        buf.append("\n\n=== Summary statistics (numeric) ===\n")
        buf.append(self.df.describe().transpose().to_string())
        buf.append("\n\n=== Top categorical counts ===\n")
        cats = self.df.select_dtypes(include=["object", "category"]).columns
        for c in cats[:6]:
            buf.append(f"\n-- {c} --")
            buf.append(self.df[c].value_counts().head(3).to_string())

        self.summary_text.delete("0.0", "end")
        self.summary_text.insert("0.0", "\n".join(buf))

        for w in self.dashboard_plot_container.winfo_children():
            w.destroy()

        num_df = self.df.select_dtypes(include=[np.number])
        if num_df.shape[1] >= 2:
            fig = Figure(figsize=(6,3), dpi=100)
            ax = fig.add_subplot(111)
            c = num_df.corr()
            im = ax.matshow(c, fignum=False)
            ax.set_xticks(range(len(c.columns)))
            ax.set_xticklabels(c.columns, rotation=45, fontsize=8)
            ax.set_yticks(range(len(c.columns)))
            ax.set_yticklabels(c.columns, fontsize=8)
            ax.set_title("Correlation (preview)")
            fig.colorbar(im, ax=ax)
            canvas = FigureCanvasTkAgg(fig, master=self.dashboard_plot_container)
            canvas.draw()
            canvas.get_tk_widget().pack(expand=True, fill="both")
        else:
            ctk.CTkLabel(self.dashboard_plot_container, text="Not enough numeric columns for correlation preview").pack(expand=True)

        self.tabview.set("Dashboard")

    # ---------- Cleaning ----------
    def clean_data(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load file first.")
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Clean Data")
        dialog.geometry("420x260")

        ctk.CTkLabel(dialog, text="Cleaning Options", font=ctk.CTkFont(size=15, weight="bold")).pack(pady=12)

        dropna = ctk.BooleanVar(value=False)
        fillna = ctk.BooleanVar(value=False)
        dup = ctk.BooleanVar(value=False)

        ctk.CTkCheckBox(dialog, text="Drop rows with any NA", variable=dropna).pack(anchor="w", padx=20, pady=6)
        ctk.CTkCheckBox(dialog, text="Fill NA: numeric->0, others->''", variable=fillna).pack(anchor="w", padx=20, pady=6)
        ctk.CTkCheckBox(dialog, text="Remove duplicate rows", variable=dup).pack(anchor="w", padx=20, pady=6)

        def apply_clean():
            df = self.df.copy()
            if dropna.get():
                df = df.dropna()
            if fillna.get():
                num_cols = df.select_dtypes(include=[np.number]).columns
                obj_cols = df.select_dtypes(exclude=[np.number]).columns
                df[num_cols] = df[num_cols].fillna(0)
                df[obj_cols] = df[obj_cols].fillna("")
            if dup.get():
                df = df.drop_duplicates()
            self.df = df.reset_index(drop=True)
            self.status_label.configure(text=f"Cleaned | Rows: {len(self.df)} | Cols: {len(self.df.columns)}")
            dialog.destroy()
            self.preview_data()
            messagebox.showinfo("Clean", "Data cleaning applied.")

        btn = ctk.CTkFrame(dialog)
        btn.pack(fill="x", pady=8, padx=10)
        ctk.CTkButton(btn, text="Apply", command=apply_clean).pack(side="right", padx=8)
        ctk.CTkButton(btn, text="Cancel", command=dialog.destroy).pack(side="right")

    # ---------- Plot area ----------
    def _build_plot_area(self):
        controls = ctk.CTkFrame(self.plot_frame)
        controls.pack(fill="x", padx=10, pady=(8,6))

        ctk.CTkLabel(controls, text="Plot Controls").pack(side="left", padx=6)
        self.plot_info = ctk.CTkLabel(controls, text="")
        self.plot_info.pack(side="right", padx=6)

        self.plot_container = ctk.CTkFrame(self.plot_frame)
        self.plot_container.pack(expand=True, fill="both", padx=10, pady=(6,10))

    def prompt_two_columns(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load a file first.")
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select Columns")
        dialog.geometry("420x260")

        cols = list(self.df.columns)
        ctk.CTkLabel(dialog, text="X column:").pack(pady=(12,4))
        xmenu = ctk.CTkOptionMenu(dialog, values=cols)
        xmenu.pack()
        ctk.CTkLabel(dialog, text="Y column:").pack(pady=(12,4))
        ymenu = ctk.CTkOptionMenu(dialog, values=cols)
        ymenu.pack()

        plot_type = ctk.CTkOptionMenu(dialog, values=["Line","Scatter","Bar"])
        plot_type.set("Line")
        ctk.CTkLabel(dialog, text="Plot type:").pack(pady=(12,4))
        plot_type.pack()

        def do_plot():
            x = xmenu.get()
            y = ymenu.get()
            t = plot_type.get()
            dialog.destroy()
            self._plot_columns(x, y, t)

        ctk.CTkButton(dialog, text="Plot", command=do_plot).pack(pady=12)

    def plot_two_columns(self):
        self.prompt_two_columns()

    def _plot_columns(self, xcol, ycol, ptype="Line"):
        for w in self.plot_container.winfo_children():
            w.destroy()
        fig = Figure(figsize=(8,5), dpi=100)
        ax = fig.add_subplot(111)
        try:
            x = self.df[xcol]
            y = self.df[ycol]
            if ptype == "Line":
                ax.plot(x, y, linewidth=1.5)
            elif ptype == "Scatter":
                ax.scatter(x, y, alpha=0.8)
            elif ptype == "Bar":
                if not np.issubdtype(x.dtype, np.number):
                    agg = self.df.groupby(xcol)[ycol].mean().sort_values(ascending=False)
                    ax.bar(agg.index.astype(str), agg.values)
                    ax.set_xticklabels(agg.index.astype(str), rotation=45, ha="right")
                else:
                    ax.bar(x, y)
            ax.set_xlabel(xcol)
            ax.set_ylabel(ycol)
            ax.set_title(f"{ptype}: {ycol} vs {xcol}")
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=self.plot_container)
            canvas.draw()
            canvas.get_tk_widget().pack(expand=True, fill="both")
            toolbar = NavigationToolbar2Tk(canvas, self.plot_container)
            toolbar.update()
            toolbar.pack()
            self.plot_info.configure(text=f"{ptype} plotted")
            self.tabview.set("Plot")
        except Exception as e:
            messagebox.showerror("Plot error", str(e))

    def plot_histogram(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load a file first.")
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Histogram")
        dialog.geometry("360x220")

        cols = list(self.df.columns)
        ctk.CTkLabel(dialog, text="Select column:").pack(pady=(12,6))
        colmenu = ctk.CTkOptionMenu(dialog, values=cols)
        colmenu.pack()

        bins_entry = ctk.CTkEntry(dialog)
        bins_entry.insert(0, "20")
        ctk.CTkLabel(dialog, text="Bins:").pack(pady=(8,2))
        bins_entry.pack()

        def do_hist():
            col = colmenu.get()
            try:
                bins = int(bins_entry.get())
            except Exception:
                bins = 20
            dialog.destroy()
            for w in self.plot_container.winfo_children():
                w.destroy()
            fig = Figure(figsize=(8,5), dpi=100)
            ax = fig.add_subplot(111)
            try:
                ax.hist(self.df[col].dropna(), bins=bins)
                ax.set_title(f"Histogram of {col}")
                ax.set_xlabel(col)
                fig.tight_layout()
                canvas = FigureCanvasTkAgg(fig, master=self.plot_container)
                canvas.draw()
                canvas.get_tk_widget().pack(expand=True, fill="both")
                toolbar = NavigationToolbar2Tk(canvas, self.plot_container)
                toolbar.update()
                toolbar.pack()
                self.plot_info.configure(text=f"Histogram ({col})")
                self.tabview.set("Plot")
            except Exception as e:
                messagebox.showerror("Histogram error", str(e))

        ctk.CTkButton(dialog, text="Plot", command=do_hist).pack(pady=12)

    # ---------- Correlation ----------
    def show_correlation(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load file first.")
            return
        num_df = self.df.select_dtypes(include=[np.number])
        if num_df.shape[1] < 2:
            messagebox.showwarning("No numeric columns", "Need at least two numeric columns for correlation.")
            return

        for w in self.plot_container.winfo_children():
            w.destroy()
        fig = Figure(figsize=(8,6), dpi=100)
        ax = fig.add_subplot(111)
        c = num_df.corr()
        im = ax.matshow(c, fignum=False)
        ax.set_xticks(range(len(c.columns)))
        ax.set_xticklabels(c.columns, rotation=45, fontsize=8)
        ax.set_yticks(range(len(c.columns)))
        ax.set_yticklabels(c.columns, fontsize=8)
        ax.set_title("Correlation matrix")
        fig.colorbar(im, ax=ax)
        canvas = FigureCanvasTkAgg(fig, master=self.plot_container)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill="both")
        toolbar = NavigationToolbar2Tk(canvas, self.plot_container)
        toolbar.update()
        toolbar.pack()
        self.tabview.set("Plot")

    # ---------- Export ----------
    def export_csv(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load file first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            self.df.to_csv(path, index=False)
            messagebox.showinfo("Export", f"Saved cleaned CSV to {path}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    def export_pdf_report(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No data", "Load file first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        # start background thread and show simple "working" modal
        progress = ctk.CTkToplevel(self)
        progress.geometry("360x120")
        progress.title("Generating PDF")
        ctk.CTkLabel(progress, text="Generating PDF report...\nPlease wait", anchor="center").pack(expand=True, pady=12)
        progress.grab_set()

        def worker():
            try:
                self._create_pdf_report(path)
                progress.destroy()
                messagebox.showinfo("PDF Report", f"Report saved to {path}")
            except Exception as e:
                progress.destroy()
                messagebox.showerror("PDF Error", str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _create_pdf_report(self, path):
        tempdir = tempfile.mkdtemp()
        images = []

        # summary text
        summary_txt = os.path.join(tempdir, "summary.txt")
        with open(summary_txt, "w", encoding="utf-8") as f:
            f.write("📊 Milo – Data Insight Pro - Report\n")
            f.write(f"File: {os.path.basename(self.current_file) if self.current_file else '(in-memory)'}\n")
            f.write(f"Rows: {len(self.df)} | Columns: {len(self.df.columns)}\n\n")
            f.write(self.df.describe(include="all").transpose().to_string())

        # correlation
        num_df = self.df.select_dtypes(include=[np.number])
        if num_df.shape[1] >= 2:
            fig = Figure(figsize=(6,4), dpi=120)
            ax = fig.add_subplot(111)
            c = num_df.corr()
            im = ax.matshow(c, fignum=False)
            ax.set_xticks(range(len(c.columns)))
            ax.set_xticklabels(c.columns, rotation=45, fontsize=8)
            ax.set_yticks(range(len(c.columns)))
            ax.set_yticklabels(c.columns, fontsize=8)
            ax.set_title("Correlation matrix")
            fig.colorbar(im, ax=ax)
            corr_path = os.path.join(tempdir, "corr.png")
            fig.savefig(corr_path, bbox_inches="tight")
            images.append(corr_path)

        # histograms for first up-to-3 numeric columns
        for col in num_df.columns[:3]:
            fig = Figure(figsize=(6,3), dpi=120)
            ax = fig.add_subplot(111)
            ax.hist(self.df[col].dropna(), bins=20)
            ax.set_title(f"Histogram of {col}")
            hist_path = os.path.join(tempdir, f"hist_{col}.png")
            fig.savefig(hist_path, bbox_inches="tight")
            images.append(hist_path)

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=12)
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "📊 Milo – Data Insight Pro - Report", ln=True)
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 8, f"File: {os.path.basename(self.current_file) if self.current_file else '(in-memory)'}", ln=True)
        pdf.cell(0, 8, f"Rows: {len(self.df)} | Columns: {len(self.df.columns)}", ln=True)
        pdf.ln(6)

        with open(summary_txt, "r", encoding="utf-8") as st:
            lines = st.read().splitlines()
        pdf.set_font("Arial", size=9)
        for line in lines:
            # wrap long lines
            for chunk in [line[i:i+90] for i in range(0, len(line), 90)]:
                pdf.cell(0, 5, chunk, ln=True)
        pdf.ln(6)

        for img in images:
            pdf.add_page()
            try:
                pdf.image(img, x=10, y=20, w=pdf.w - 20)
            except Exception:
                continue

        pdf.output(path)

    # ---------- File handling ----------
    def load_file(self):
        filetypes = [
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("JSON files", "*.json"),
            ("All files", "*.*"),
        ]
        path = filedialog.askopenfilename(title="Open data file", filetypes=filetypes)
        if not path:
            return
        try:
            df = try_read_data(path)
            self.df = df
            self.current_file = path
            self.status_label.configure(text=f"Loaded: {os.path.basename(path)} | Rows: {len(df)} | Cols: {len(df.columns)}")
            self.preview_data()
            messagebox.showinfo("Loaded", "File loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))

    def reload_last(self):
        if not self.current_file:
            messagebox.showinfo("No file", "No previously loaded file.")
            return
        try:
            df = try_read_data(self.current_file)
            self.df = df
            self.status_label.configure(text=f"Reloaded: {os.path.basename(self.current_file)} | Rows: {len(df)} | Cols: {len(df.columns)}")
            self.preview_data()
            messagebox.showinfo("Reloaded", "File reloaded successfully.")
        except Exception as e:
            messagebox.showerror("Reload error", str(e))

    # ---------- utilities ----------
    def _format_value(self, v):
        if pd.isna(v):
            return ""
        if isinstance(v, (float, np.floating)):
            return f"{v:.4f}"
        return str(v)


if __name__ == "__main__":
    app = DataInsightPro()
    app.mainloop()
