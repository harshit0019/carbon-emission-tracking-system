import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import logging
import os
import shutil
import subprocess
import json
import pandas as pd
import numpy as np
from openpyxl import Workbook
from scipy.interpolate import make_interp_spline
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# ---------------- Global Constants (Theming) ----------------
BACKGROUND_COLOR = "#F8F9FA"
CARD_COLOR = "white"
PRIMARY_COLOR = "#3498DB"
PRIMARY_HOVER = "#2980B9"
DANGER_COLOR = "#E74C3C"
DANGER_HOVER = "#c0392b"
TEXT_COLOR = "#2C3E50"
SHADOW_COLOR = "#d3d3d3"
FONT_FAMILY = "Arial"

# ---------------- Logging Setup ----------------
logging.basicConfig(filename="app.log", level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Global data storage
# Each emission record: [Email, Entry Date, Month, Unit, Emission Category, Emission Name, Factor, Amount, Total, Document, RecordID]
emission_records = []
document_logs = []
DATA_FILE = "emission_records.json"
record_id_counter = 0

# ---------------- Persistence Functions ----------------
def save_emission_records():
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(emission_records, f)
        logging.info("Emission records saved to disk.")
    except Exception as e:
        logging.error(f"Error saving emission records: {e}")

def load_emission_records():
    global emission_records, record_id_counter
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as f:
                emission_records[:] = json.load(f)
            if emission_records:
                record_id_counter = max(int(r[10]) for r in emission_records) + 1
            logging.info("Emission records loaded from disk.")
    except Exception as e:
        logging.error(f"Error loading emission records: {e}")

# ---------------- Helper Functions ----------------
def update_total_value(factor, amount_str):
    try:
        amount = float(amount_str)
        total = factor * amount
        return f"{total:.2f}"
    except Exception as e:
        logging.error(f"Error calculating total: {e}")
        return "0.00"

# ---------------- Document Management System (DMS) ----------------
def get_user_role(email):
    if email == "manager@gmail.com":
        return "Manager"
    elif email == "employee@gmail.com":
        return "Employee"
    else:
        return "Employee"

class DocumentManagementSystem:
    BASE_DIR = "CarbonData"
    @staticmethod
    def generate_unique_code(unit_name, upload_date, emission_name, emission_type):
        dt = datetime.strptime(upload_date, "%Y-%m-%d")
        return f"{unit_name}_{dt.strftime('%d_%m_%Y')}_{emission_name}_{emission_type}"
    @staticmethod
    def get_storage_path(unit_name, upload_date):
        dt = datetime.strptime(upload_date, "%Y-%m-%d")
        folder_path = os.path.join(DocumentManagementSystem.BASE_DIR, unit_name, dt.strftime("%Y"), dt.strftime("%m_%B"))
        os.makedirs(folder_path, exist_ok=True)
        return folder_path
    @staticmethod
    def save_document(file_path, unit_name, upload_date, emission_name, emission_type, uploader, role):
        unique_code = DocumentManagementSystem.generate_unique_code(unit_name, upload_date, emission_name, emission_type)
        storage_path = DocumentManagementSystem.get_storage_path(unit_name, upload_date)
        ext = os.path.splitext(file_path)[1]
        new_file_name = f"{unique_code}{ext}"
        new_file_path = os.path.join(storage_path, new_file_name)
        version = 1
        final_file_path = new_file_path
        while os.path.exists(final_file_path):
            version += 1
            final_file_name = f"{unique_code}_v{version}{ext}"
            final_file_path = os.path.join(storage_path, final_file_name)
        shutil.copy(file_path, final_file_path)
        metadata = {
            "unique_code": unique_code,
            "file_path": final_file_path,
            "upload_date_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "uploader": uploader,
            "role": role,
            "unit_name": unit_name,
            "upload_date": upload_date,
            "emission_name": emission_name,
            "emission_type": emission_type,
            "file_status": "Pending",
            "version": version
        }
        document_logs.append(metadata)
        logging.info(f"Document uploaded: {metadata}")
        return metadata

def upload_document(var, unit, upload_date, emission_name, emission_type, uploader):
    file_path = filedialog.askopenfilename(
        filetypes=[("All files", "*.*"), ("PDF", "*.pdf"), ("Excel Files", "*.xlsx;*.xls"), ("Images", "*.png;*.jpg;*.jpeg")],
        title="Select a document to upload"
    )
    if file_path:
        role = get_user_role(uploader)
        metadata = DocumentManagementSystem.save_document(file_path, unit, upload_date, emission_name, emission_type, uploader, role)
        var.set(metadata["file_path"])
        messagebox.showinfo("File Uploaded", f"File uploaded and saved as:\n{metadata['file_path']}")

# ---------------- Custom Numeric Entry ----------------
class NumericEntry(tk.Entry):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        vcmd = (self.register(self.validate_numeric), '%P')
        self.config(validate="key", validatecommand=vcmd)
    def validate_numeric(self, new_value):
        if new_value == "":
            self.config(bg="white")
            return True
        try:
            float(new_value)
            self.config(bg="white")
            return True
        except ValueError:
            self.config(bg="#ffcccc")
            return False

# ---------------- Hover and Focus Effects ----------------
def add_hover(widget, normal_bg, hover_bg):
    widget.bind("<Enter>", lambda e: widget.config(bg=hover_bg))
    widget.bind("<Leave>", lambda e: widget.config(bg=normal_bg))

def add_focus_effect(entry, normal_bg="white", focus_bg="#e0f7fa"):
    entry.bind("<FocusIn>", lambda e: entry.config(bg=focus_bg))
    entry.bind("<FocusOut>", lambda e: entry.config(bg=normal_bg))

# ---------------- Scrollable Frame Class ----------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
         super().__init__(container, *args, **kwargs)
         canvas = tk.Canvas(self, borderwidth=0, background=BACKGROUND_COLOR)
         scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
         self.scrollable_frame = tk.Frame(canvas, background=BACKGROUND_COLOR)
         self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
         canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
         canvas.configure(yscrollcommand=scrollbar.set)
         canvas.pack(side="left", fill="both", expand=True)
         scrollbar.pack(side="right", fill="y")

# ---------------- Utility: Create a "Card" ----------------
def create_card(parent, pady=15, padx=20, fill="x"):
    shadow = tk.Frame(parent, bg=SHADOW_COLOR)
    shadow.pack(pady=pady, padx=padx, fill=fill)
    card = tk.Frame(shadow, bg=CARD_COLOR)
    card.pack(padx=3, pady=3, fill=fill)
    return card

# ---------------- MultiSelectDropdown Widget ----------------
class MultiSelectDropdown(tk.Menubutton):
    def __init__(self, parent, options, **kwargs):
        super().__init__(parent, text="All", relief="raised", indicatoron=True, borderwidth=1, **kwargs)
        self.var_dict = {}
        self.menu = tk.Menu(self, tearoff=0)
        for option in options:
            var = tk.BooleanVar(value=False)
            self.var_dict[option] = var
            self.menu.add_checkbutton(label=option, variable=var, command=self.update_text)
        self.config(menu=self.menu)
    def update_text(self):
        selected = [opt for opt, var in self.var_dict.items() if var.get()]
        self.config(text=", ".join(selected) if selected else "All")
    def get_selected(self):
        return [opt for opt, var in self.var_dict.items() if var.get()]

# ---------------- Analysis Page ----------------
class AnalysisPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        # Title and Analysis Box
        title_label = tk.Label(self, text="RMX Joss Carbon Tracking System", font=(FONT_FAMILY, 24, "bold"),
                               bg=BACKGROUND_COLOR, fg=TEXT_COLOR)
        title_label.pack(pady=10)
        analysis_box = tk.Frame(self, bg=CARD_COLOR, bd=2, relief="groove")
        analysis_box.pack(padx=10, pady=5, fill="x")
        analysis_label = tk.Label(analysis_box, text="Analysis", font=(FONT_FAMILY, 18, "bold"),
                                  bg=CARD_COLOR, fg=TEXT_COLOR)
        analysis_label.pack(pady=10)
        # Global Filters (except view mode)
        self.filter_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        self.filter_frame.pack(pady=10)
        tk.Label(self.filter_frame, text="Unit:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=0, column=0, padx=5)
        self.unit_filter = MultiSelectDropdown(self.filter_frame, options=["C-49", "B-37", "C-91", "2B-4"])
        self.unit_filter.grid(row=0, column=1, padx=5)
        tk.Label(self.filter_frame, text="Year:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=0, column=2, padx=5)
        self.analysis_year = tk.StringVar(value="All")
        self.year_combobox = ttk.Combobox(self.filter_frame, textvariable=self.analysis_year, state="readonly", width=10)
        self.year_combobox.grid(row=0, column=3, padx=5)
        tk.Label(self.filter_frame, text="Month:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=0, column=4, padx=5)
        self.analysis_month = tk.StringVar(value="All")
        month_options = ["All", "January", "February", "March", "April", "May", "June",
                         "July", "August", "September", "October", "November", "December"]
        ttk.Combobox(self.filter_frame, textvariable=self.analysis_month, values=month_options, state="readonly", width=10).grid(row=0, column=5, padx=5)
        btn_update = tk.Button(self.filter_frame, text="Update Analysis", command=self.update_analysis,
                               bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10, "bold"))
        btn_update.grid(row=0, column=6, padx=10)
        # KPI Cards
        self.kpi_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        self.kpi_frame.pack(fill="x", padx=10, pady=5)
        self.kpi_total = self.create_kpi_card(self.kpi_frame, "Total Emissions", "0.00 tons")
        self.kpi_scope1 = self.create_kpi_card(self.kpi_frame, "Scope 1 Emissions", "0.00 tons")
        self.kpi_scope2 = self.create_kpi_card(self.kpi_frame, "Scope 2 Emissions", "0.00 tons")
        self.kpi_top_unit = self.create_kpi_card(self.kpi_frame, "Highest Emitting Unit", "N/A")
        self.kpi_top_gas = self.create_kpi_card(self.kpi_frame, "Most Emitted Gas", "N/A")
        self.kpi_top_category = self.create_kpi_card(self.kpi_frame, "Highest Emission Category", "N/A")
        for card in [self.kpi_total, self.kpi_scope1, self.kpi_scope2, self.kpi_top_unit, self.kpi_top_gas, self.kpi_top_category]:
            card.pack(side="left", padx=10, pady=10)
        # Charts Area
        self.charts_container = tk.Frame(self, bg=BACKGROUND_COLOR, bd=2, relief="groove")
        self.charts_container.pack(fill="both", expand=True, padx=10, pady=10)
        # Left Chart Box: Line Chart with View Mode control inside
        self.left_chart_frame = tk.Frame(self.charts_container, bg=BACKGROUND_COLOR, bd=1, relief="solid")
        self.left_chart_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.line_chart_control_frame = tk.Frame(self.left_chart_frame, bg=BACKGROUND_COLOR)
        self.line_chart_control_frame.pack(side="top", fill="x", padx=5, pady=5)
        tk.Label(self.line_chart_control_frame, text="View Mode:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).pack(side="left", padx=5)
        self.view_mode = tk.StringVar(value="Monthly")
        view_mode_cb = ttk.Combobox(self.line_chart_control_frame, textvariable=self.view_mode, values=["Monthly", "Yearly"], state="readonly", width=10)
        view_mode_cb.pack(side="left", padx=5)
        self.view_mode.trace_add("write", lambda *args: self.update_analysis())
        self.fig_line = Figure(figsize=(6,4), dpi=100)
        self.ax_line = self.fig_line.add_subplot(111)
        self.canvas_line = FigureCanvasTkAgg(self.fig_line, master=self.left_chart_frame)
        self.canvas_line.get_tk_widget().pack(fill="both", expand=True)
        # Right Chart Box: Donut Chart
        self.right_chart_frame = tk.Frame(self.charts_container, bg=BACKGROUND_COLOR, bd=1, relief="solid")
        self.right_chart_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.fig_donut = Figure(figsize=(4,4), dpi=100)
        self.ax_donut = self.fig_donut.add_subplot(111)
        self.canvas_donut = FigureCanvasTkAgg(self.fig_donut, master=self.right_chart_frame)
        self.canvas_donut.get_tk_widget().pack(fill="both", expand=True)
        # Navigation Buttons
        nav_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        nav_frame.pack(pady=10)
        btn_home = tk.Button(nav_frame, text="Go to Home", command=lambda: self.controller.show_frame("HomePage"),
                             bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"))
        btn_home.pack(side="left", padx=10)
        btn_emission = tk.Button(nav_frame, text="Go to Emission Data", command=lambda: self.controller.show_frame("EmissionDataPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"))
        btn_emission.pack(side="left", padx=10)
        btn_refresh = tk.Button(nav_frame, text="Refresh", command=self.update_analysis,
                                bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"))
        btn_refresh.pack(side="left", padx=10)
        self.update_analysis()
    
    def create_kpi_card(self, parent, title, value):
        frame = tk.Frame(parent, bg="white", bd=2, relief="solid", padx=10, pady=10)
        lbl_title = tk.Label(frame, text=title, font=(FONT_FAMILY, 10, "bold"), bg="white", fg=TEXT_COLOR)
        lbl_value = tk.Label(frame, text=value, font=(FONT_FAMILY, 12), bg="white", fg=TEXT_COLOR)
        lbl_title.pack()
        lbl_value.pack()
        frame.lbl_value = lbl_value
        return frame
    
    def update_year_options(self, df):
        if not df.empty:
            years = sorted(set(df["Entry Date"].dt.year))
            options = ["All"] + [str(y) for y in years]
        else:
            options = ["All"]
        self.year_combobox['values'] = options
    
    def update_analysis(self):
        if not emission_records:
            return
        df = pd.DataFrame(emission_records, columns=["Email", "Entry Date", "Month", "Unit",
                                                     "Emission Category", "Emission Name",
                                                     "Factor", "Amount", "Total", "Document", "RecordID"])
        df["Entry Date"] = pd.to_datetime(df["Entry Date"])
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce")
        self.update_year_options(df)
        # For line chart:
        if self.view_mode.get() == "Monthly":
            if self.analysis_year.get() != "All":
                selected_year = int(self.analysis_year.get())
                df_line = df[df["Entry Date"].dt.year == selected_year].copy()
            else:
                df_line = df.copy()
        else:
            df_line = df.copy()
        # For KPIs and donut chart, apply year filter if not "All"
        df_donut = df.copy()
        if self.analysis_year.get() != "All":
            selected_year_donut = int(self.analysis_year.get())
            df_donut = df_donut[df_donut["Entry Date"].dt.year == selected_year_donut]
        if self.analysis_month.get() != "All":
            df_donut = df_donut[df_donut["Month"] == self.analysis_month.get()]
        # Apply unit filter for both
        selected_units = self.unit_filter.get_selected()
        if selected_units:
            df_line = df_line[df_line["Unit"].isin(selected_units)]
            df_donut = df_donut[df_donut["Unit"].isin(selected_units)]
        # Compute KPIs from df_donut
        total_emissions = df_donut["Total"].sum()
        scope1 = df_donut[df_donut["Emission Category"].isin(["Fuel", "Refrigerants"])]["Total"].sum()
        scope2 = df_donut[df_donut["Emission Category"] == "Electricity"]["Total"].sum()
        top_unit_group = df_donut.groupby("Unit")["Total"].sum()
        top_unit = top_unit_group.idxmax() if not top_unit_group.empty else "N/A"
        top_unit_value = top_unit_group.max() if not top_unit_group.empty else 0
        top_gas_group = df_donut.groupby("Emission Name")["Total"].sum()
        top_gas = top_gas_group.idxmax() if not top_gas_group.empty else "N/A"
        top_category_group = df_donut.groupby("Emission Category")["Total"].sum()
        top_category = top_category_group.idxmax() if not top_category_group.empty else "N/A"
        self.kpi_total.lbl_value.config(text=f"{total_emissions:.2f} tons")
        self.kpi_scope1.lbl_value.config(text=f"{scope1:.2f} tons")
        self.kpi_scope2.lbl_value.config(text=f"{scope2:.2f} tons")
        self.kpi_top_unit.lbl_value.config(text=f"{top_unit} ({top_unit_value:.2f} tons)")
        self.kpi_top_gas.lbl_value.config(text=f"{top_gas}")
        self.kpi_top_category.lbl_value.config(text=f"{top_category}")
        # ---------------- Line Chart ----------------
        self.ax_line.clear()
        if self.view_mode.get() == "Monthly":
            month_order = {"January":1, "February":2, "March":3, "April":4, "May":5, "June":6,
                           "July":7, "August":8, "September":9, "October":10, "November":11, "December":12}
            df_line["MonthNum"] = df_line["Month"].map(month_order)
            pivot = df_line.pivot_table(index="Month", columns="Unit", values="Total", aggfunc="sum", fill_value=0)
            pivot = pivot.reindex(sorted(pivot.index, key=lambda x: month_order.get(x, 0)))
            x_orig = np.array(range(len(pivot.index)))
            for unit in pivot.columns:
                y_orig = pivot[unit].values
                if len(x_orig) > 1:
                    k = min(3, len(x_orig)-1)
                    x_new = np.linspace(x_orig.min(), x_orig.max(), 300)
                    spl = make_interp_spline(x_orig, y_orig, k=k)
                    y_smooth = spl(x_new)
                    self.ax_line.plot(x_new, y_smooth, label=unit, clip_on=True)
                else:
                    self.ax_line.plot(x_orig, y_orig, marker="o", label=unit, clip_on=True)
            self.ax_line.set_xticks(x_orig)
            self.ax_line.set_xticklabels(pivot.index, rotation=45, ha="right")
            title_str = "Monthly Emissions by Unit"
            if self.analysis_year.get() != "All":
                title_str += f" for {self.analysis_year.get()}"
            self.ax_line.set_title(title_str)
            self.ax_line.set_xlabel("Month")
            self.ax_line.set_ylabel("Emissions (tons)")
            data = pivot.values.flatten()
            if len(data) > 0:
                y_min, y_max = np.min(data), np.max(data)
                margin = (y_max - y_min)*0.1 if y_max > y_min else 10
                self.ax_line.set_ylim(y_min - margin, y_max + margin)
            else:
                self.ax_line.set_ylim(700, 1300)
        else:
            df_line["Year"] = df_line["Entry Date"].dt.year
            years = sorted(df_line["Year"].unique())
            if not years:
                years = [2023, 2033]
            pivot = df_line.pivot_table(index="Year", columns="Unit", values="Total", aggfunc="sum", fill_value=0)
            pivot = pivot.reindex(years, fill_value=0)
            x_orig = np.array(years)
            for unit in pivot.columns:
                y = pivot[unit].values
                if len(x_orig) > 1:
                    k = min(3, len(x_orig)-1)
                    x_new = np.linspace(x_orig.min(), x_orig.max(), 300)
                    spl = make_interp_spline(x_orig, y, k=k)
                    y_smooth = spl(x_new)
                    self.ax_line.plot(x_new, y_smooth, label=unit, clip_on=True)
                else:
                    self.ax_line.plot(x_orig, y, marker="o", label=unit, clip_on=True)
            self.ax_line.set_title("Yearly Emissions by Unit")
            self.ax_line.set_xlabel("Year")
            self.ax_line.set_ylabel("Emissions (tons)")
            self.ax_line.set_xticks(x_orig)
            self.ax_line.set_xticklabels(x_orig)
            data = pivot.values.flatten()
            if len(data) > 0:
                y_min, y_max = np.min(data), np.max(data)
                margin = (y_max - y_min)*0.1 if y_max > y_min else 10
                self.ax_line.set_ylim(y_min - margin, y_max + margin)
            else:
                self.ax_line.set_ylim(500, 14500)
        self.ax_line.grid(True)
        self.ax_line.legend(bbox_to_anchor=(1.05, 1), loc="upper left", borderaxespad=0)
        self.fig_line.subplots_adjust(bottom=0.35, right=0.75)
        self.canvas_line.draw()
        # ---------------- Donut Chart ----------------
        self.ax_donut.clear()
        df_donut_filtered = df.copy()
        if self.analysis_year.get() != "All":
            selected_year_donut = int(self.analysis_year.get())
            df_donut_filtered = df_donut_filtered[df_donut_filtered["Entry Date"].dt.year == selected_year_donut]
        if self.analysis_month.get() != "All":
            df_donut_filtered = df_donut_filtered[df_donut_filtered["Month"] == self.analysis_month.get()]
        if selected_units:
            df_donut_filtered = df_donut_filtered[df_donut_filtered["Unit"].isin(selected_units)]
        unit_group = df_donut_filtered.groupby("Unit")["Total"].sum()
        if not unit_group.empty:
            labels = unit_group.index.tolist()
            totals = unit_group.values.tolist()
            def make_autopct(allvals):
                def my_autopct(pct):
                    total = sum(allvals)
                    absolute = int(round(pct/100.*total))
                    return f"{pct:.1f}%\n({absolute} tons)"
                return my_autopct
            wedges, texts, autotexts = self.ax_donut.pie(totals, labels=labels,
                                                          autopct=make_autopct(totals),
                                                          startangle=90, wedgeprops=dict(width=0.4))
            self.ax_donut.set_title("Emission Distribution by Unit")
        else:
            self.ax_donut.text(0.5, 0.5, "No Data", horizontalalignment='center', verticalalignment='center')
        self.canvas_donut.draw()
    
    def create_kpi_card(self, parent, title, value):
        frame = tk.Frame(parent, bg="white", bd=2, relief="solid", padx=10, pady=10)
        lbl_title = tk.Label(frame, text=title, font=(FONT_FAMILY, 10, "bold"), bg="white", fg=TEXT_COLOR)
        lbl_value = tk.Label(frame, text=value, font=(FONT_FAMILY, 12), bg="white", fg=TEXT_COLOR)
        lbl_title.pack()
        lbl_value.pack()
        frame.lbl_value = lbl_value
        return frame
    
    def tkraise(self, aboveThis=None):
        self.update_analysis()
        super().tkraise(aboveThis)

# ---------------- Main Application (Single Window) ----------------
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RMX Joss Carbon Tracking System")
        self.geometry("1100x900")
        self.email = None
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.frames = {}
        for F in (LoginPage, HomePage, DataEntryPage, EmissionDataPage, AnalysisPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        load_emission_records()
        self.show_frame("LoginPage")
    
    def show_frame(self, page_name):
        frame = self.frames[page_name]
        if hasattr(frame, "update_role_buttons"):
            frame.update_role_buttons()
        if hasattr(frame, "user_label"):
            frame.user_label.config(text=f"User: {self.email}")
        if page_name == "EmissionDataPage":
            frame.refresh_table()
        frame.tkraise()
    
    def logout(self):
        self.email = None
        self.frames["LoginPage"].reset()
        self.show_frame("LoginPage")

class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        frame = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=300, height=250)
        tk.Label(frame, text="Login to RMX Joss System", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 14, "bold")).pack(pady=10)
        tk.Label(frame, text="Email:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).pack(pady=5)
        self.email_entry = tk.Entry(frame, width=30)
        self.email_entry.pack(pady=5)
        tk.Label(frame, text="Password:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).pack(pady=5)
        self.password_entry = tk.Entry(frame, show="*", width=30)
        self.password_entry.pack(pady=5)
        btn_login = tk.Button(frame, text="Login", command=self.login, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10, "bold"))
        btn_login.pack(pady=10)
        add_hover(btn_login, PRIMARY_COLOR, PRIMARY_HOVER)
    
    def login(self):
        email = self.email_entry.get().strip()
        password = self.password_entry.get().strip()
        if email in ["employee@gmail.com", "manager@gmail.com"] and ((email == "employee@gmail.com" and password == "1234") or (email == "manager@gmail.com" and password == "admin")):
            logging.info(f"User {email} logged in successfully.")
            self.controller.email = email
            self.controller.show_frame("HomePage")
        else:
            messagebox.showerror("Login Failed", "Invalid credentials.")
    
    def reset(self):
        self.email_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)

class HomePage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        card = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
        card.place(relx=0.5, rely=0.5, anchor="center", width=500, height=400)
        tk.Label(card, text="Welcome to RMX Joss Carbon Tracking System", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 16, "bold")).pack(pady=20)
        self.user_label = tk.Label(card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 12))
        self.user_label.pack(pady=10)
        btn_data = tk.Button(card, text="Data Entry", command=lambda: controller.show_frame("DataEntryPage"),
                             bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
        btn_data.pack(pady=10)
        add_hover(btn_data, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_emission = tk.Button(card, text="Emission Data", command=lambda: controller.show_frame("EmissionDataPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
        btn_emission.pack(pady=10)
        add_hover(btn_emission, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_analysis = tk.Button(card, text="Analysis", command=lambda: controller.show_frame("AnalysisPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
        btn_analysis.pack(pady=10)
        add_hover(btn_analysis, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_logout = tk.Button(card, text="Logout", command=lambda: controller.logout(),
                               bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
        btn_logout.pack(pady=10)
        add_hover(btn_logout, DANGER_COLOR, DANGER_HOVER)
    
    def tkraise(self, aboveThis=None):
        self.user_label.config(text=f"Logged in as: {self.controller.email}")
        super().tkraise(aboveThis)

class EmissionDataPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        self.sort_ascending = True
        self.main_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        
        header_label = tk.Label(self.main_frame, text="RMX Joss Carbon Emission Tracking System", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 20, "bold"))
        header_label.pack(pady=10)
        header_card = create_card(self.main_frame)
        tk.Label(header_card, text="Emission Data Records", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 16, "bold")).pack(pady=10)
        self.user_label = tk.Label(header_card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 12))
        self.user_label.pack(pady=5)
        
        # Filter Controls
        filter_frame = tk.Frame(self.main_frame, bg=BACKGROUND_COLOR)
        filter_frame.pack(pady=10)
        tk.Label(filter_frame, text="Filter By:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 12, "bold")).grid(row=0, column=0, padx=5)
        tk.Label(filter_frame, text="Unit:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=1, padx=5)
        self.filter_unit = tk.StringVar(value="All")
        unit_options = ["All", "C-49", "B-37", "C-91", "2B-4"]
        ttk.Combobox(filter_frame, textvariable=self.filter_unit, values=unit_options, state="readonly", width=10).grid(row=0, column=2, padx=5)
        tk.Label(filter_frame, text="Month:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=3, padx=5)
        self.filter_month = tk.StringVar(value="All")
        month_options = ["All", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        ttk.Combobox(filter_frame, textvariable=self.filter_month, values=month_options, state="readonly", width=10).grid(row=0, column=4, padx=5)
        tk.Label(filter_frame, text="Year:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=5, padx=5)
        self.filter_year = tk.StringVar(value="All")
        year_options = ["All"] + [str(year) for year in range(2020, 2031)]
        ttk.Combobox(filter_frame, textvariable=self.filter_year, values=year_options, state="readonly", width=10).grid(row=0, column=6, padx=5)
        tk.Label(filter_frame, text="Emission Type:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=7, padx=5)
        self.filter_emission_type = tk.StringVar(value="All")
        type_options = ["All", "Fuel", "Refrigerants", "Electricity"]
        ttk.Combobox(filter_frame, textvariable=self.filter_emission_type, values=type_options, state="readonly", width=12).grid(row=0, column=8, padx=5)
        btn_apply = tk.Button(filter_frame, text="Apply Filters", command=self.apply_filters, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10))
        btn_apply.grid(row=0, column=9, padx=5)
        add_hover(btn_apply, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_clear = tk.Button(filter_frame, text="Clear Filters", command=self.clear_filters, bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 10))
        btn_clear.grid(row=0, column=10, padx=5)
        add_hover(btn_clear, DANGER_COLOR, DANGER_HOVER)
        btn_sort = tk.Button(filter_frame, text="Sort by Date", command=self.sort_by_date, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10))
        btn_sort.grid(row=0, column=11, padx=5)
        add_hover(btn_sort, PRIMARY_COLOR, PRIMARY_HOVER)
        
        # Manager-only buttons
        self.btn_edit = tk.Button(filter_frame, text="Edit Record", command=self.edit_record, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10), padx=5, pady=2)
        self.btn_delete = tk.Button(filter_frame, text="Delete Record", command=self.delete_record, bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 10), padx=5, pady=2)
        
        # Emission Data Table
        table_card = create_card(self.main_frame)
        columns = ("Gmail", "Entry Date", "Month", "Unit", "Emission Category", "Emission Name", "Factor", "Amount", "Total", "Document")
        self.tree = ttk.Treeview(table_card, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=100)
        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        table_card.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_treeview_double_click)
        
        btn_frame = tk.Frame(self.main_frame, bg=BACKGROUND_COLOR)
        btn_frame.pack(pady=20)
        btn_refresh = tk.Button(btn_frame, text="Refresh", command=self.refresh_table, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), bd=0, padx=20, pady=10)
        btn_refresh.pack(side="left", padx=10)
        add_hover(btn_refresh, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_export = tk.Button(btn_frame, text="Export to Excel", command=self.export_to_excel, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), bd=0, padx=20, pady=10)
        btn_export.pack(side="left", padx=10)
        add_hover(btn_export, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_go_data = tk.Button(btn_frame, text="Go to Data Entry", command=lambda: self.controller.show_frame("DataEntryPage"), bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), bd=0, padx=20, pady=10)
        btn_go_data.pack(side="left", padx=10)
        add_hover(btn_go_data, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_analysis = tk.Button(btn_frame, text="Go to Analysis", command=lambda: self.controller.show_frame("AnalysisPage"), bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), bd=0, padx=20, pady=10)
        btn_analysis.pack(side="left", padx=10)
        add_hover(btn_analysis, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_back = tk.Button(btn_frame, text="Back to Home", command=lambda: self.controller.show_frame("HomePage"), bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), bd=0, padx=20, pady=10)
        btn_back.pack(side="left", padx=10)
        add_hover(btn_back, DANGER_COLOR, DANGER_HOVER)
        self.refresh_table()
    
    def update_role_buttons(self):
        role = get_user_role(self.controller.email)
        if role == "Manager":
            self.btn_edit.grid(row=0, column=12, padx=5)
            self.btn_delete.grid(row=0, column=13, padx=5)
        else:
            self.btn_edit.grid_forget()
            self.btn_delete.grid_forget()
    
    def refresh_table(self, records=None):
        if records is None:
            records = emission_records
        for item in self.tree.get_children():
            self.tree.delete(item)
        for record in records:
            self.tree.insert("", "end", iid=str(record[10]), values=record[:10])
        logging.info("Emission table refreshed.")
    
    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                 title="Save as")
        if file_path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Emission Data"
                headers = ("Gmail", "Entry Date", "Month", "Unit", "Emission Category", "Emission Name", "Factor", "Amount", "Total", "Document")
                ws.append(headers)
                for record in emission_records:
                    ws.append(record[:10])
                wb.save(file_path)
                messagebox.showinfo("Export Successful", f"Data exported successfully to:\n{file_path}")
                logging.info("Data exported to Excel.")
            except Exception as e:
                logging.error(f"Export to Excel failed: {e}")
                messagebox.showerror("Export Failed", f"An error occurred: {e}")
    
    def apply_filters(self):
        filtered = []
        unit_filter = self.filter_unit.get()
        month_filter = self.filter_month.get()
        year_filter = self.filter_year.get()
        type_filter = self.filter_emission_type.get()
        for record in emission_records:
            if ((unit_filter == "All" or record[3] == unit_filter) and
                (month_filter == "All" or record[2] == month_filter) and
                (year_filter == "All" or record[1][:4] == year_filter) and
                (type_filter == "All" or record[4] == type_filter)):
                filtered.append(record)
        self.refresh_table(filtered)
    
    def clear_filters(self):
        self.filter_unit.set("All")
        self.filter_month.set("All")
        self.filter_year.set("All")
        self.filter_emission_type.set("All")
        self.refresh_table(emission_records)
    
    def sort_by_date(self):
        unit_filter = self.filter_unit.get()
        month_filter = self.filter_month.get()
        year_filter = self.filter_year.get()
        type_filter = self.filter_emission_type.get()
        filtered = []
        for record in emission_records:
            if ((unit_filter == "All" or record[3] == unit_filter) and
                (month_filter == "All" or record[2] == month_filter) and
                (year_filter == "All" or record[1][:4] == year_filter) and
                (type_filter == "All" or record[4] == type_filter)):
                filtered.append(record)
        filtered.sort(key=lambda x: x[1], reverse=not self.sort_ascending)
        self.sort_ascending = not self.sort_ascending
        self.refresh_table(filtered)
    
    def on_treeview_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            if col == "#10":
                item = self.tree.identify_row(event.y)
                if item:
                    values = self.tree.item(item, "values")
                    file_path = values[9]
                    if file_path != "No File" and os.path.exists(file_path):
                        try:
                            if os.name == 'nt':
                                os.startfile(file_path)
                            else:
                                subprocess.call(['open', file_path])
                        except Exception as e:
                            messagebox.showerror("File Error", f"Unable to open file:\n{e}")
                    else:
                        messagebox.showerror("File Error", "File not found or cannot be opened.")
    
    def edit_record(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("No Selection", "Please select a record to edit.")
            return
        record_id = selected[0]
        record = None
        for i, rec in enumerate(emission_records):
            if str(rec[10]) == record_id:
                record = rec
                rec_index = i
                break
        if record is None:
            messagebox.showerror("Error", "Record not found.")
            return
        EditDialog(self, record, rec_index)
    
    def delete_record(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("No Selection", "Please select a record to delete.")
            return
        record_id = selected[0]
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record?"):
            return
        global emission_records
        for i, rec in enumerate(emission_records):
            if str(rec[10]) == record_id:
                file_path = rec[9]
                if file_path != "No File" and os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        logging.info(f"Deleted file from storage: {file_path}")
                    except Exception as e:
                        logging.error(f"Error deleting file {file_path}: {e}")
                del emission_records[i]
                break
        save_emission_records()
        self.refresh_table()
        messagebox.showinfo("Deleted", "Record deleted successfully.")

class EditDialog(tk.Toplevel):
    def __init__(self, parent_page, record, rec_index):
        super().__init__(parent_page)
        self.title("Edit Record")
        self.parent_page = parent_page
        self.rec_index = rec_index
        tk.Label(self, text="Unit:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.unit_var = tk.StringVar(value=record[3])
        tk.Entry(self, textvariable=self.unit_var).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(self, text="Month:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.month_var = tk.StringVar(value=record[2])
        tk.Entry(self, textvariable=self.month_var).grid(row=1, column=1, padx=5, pady=5)
        tk.Label(self, text="Emission Category:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.cat_var = tk.StringVar(value=record[4])
        tk.Entry(self, textvariable=self.cat_var).grid(row=2, column=1, padx=5, pady=5)
        tk.Label(self, text="Emission Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.name_var = tk.StringVar(value=record[5])
        tk.Entry(self, textvariable=self.name_var).grid(row=3, column=1, padx=5, pady=5)
        tk.Label(self, text="Factor:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.factor_var = tk.StringVar(value=record[6])
        tk.Entry(self, textvariable=self.factor_var).grid(row=4, column=1, padx=5, pady=5)
        tk.Label(self, text="Amount:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.amount_var = tk.StringVar(value=record[7])
        tk.Entry(self, textvariable=self.amount_var).grid(row=5, column=1, padx=5, pady=5)
        tk.Label(self, text="Document:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
        self.doc_var = tk.StringVar(value=record[9])
        tk.Entry(self, textvariable=self.doc_var, width=40).grid(row=6, column=1, padx=5, pady=5)
        btn_save = tk.Button(self, text="Save Changes", command=self.save_changes, bg=PRIMARY_COLOR, fg="white")
        btn_save.grid(row=7, column=0, columnspan=2, pady=10)
    
    def save_changes(self):
        total = update_total_value(float(self.factor_var.get()), self.amount_var.get())
        original = emission_records[self.rec_index]
        updated = (original[0], original[1], self.month_var.get(), self.unit_var.get(),
                   self.cat_var.get(), self.name_var.get(), self.factor_var.get(),
                   self.amount_var.get(), total, self.doc_var.get(), original[10])
        emission_records[self.rec_index] = updated
        save_emission_records()
        self.parent_page.refresh_table()
        messagebox.showinfo("Updated", "Record updated successfully!")
        self.destroy()

class DataEntryPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        self.fuel_file_vars = {}
        self.refrig_file_vars = {}
        self.elec_file_var = tk.StringVar()
        # Use ScrollableFrame for Data Entry Page
        self.main_frame = ScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True)
        # Header
        header_label = tk.Label(self.main_frame.scrollable_frame, text="RMX Joss Carbon Emission Tracking System",
                                bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 20, "bold"))
        header_label.pack(pady=10)
        # Top Card with common fields
        top_card = create_card(self.main_frame.scrollable_frame)
        tk.Label(top_card, text="Choose Unit:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.unit_var = tk.StringVar()
        unit_dropdown = ttk.Combobox(top_card, textvariable=self.unit_var, state="readonly", width=12)
        unit_dropdown['values'] = ("C-49", "B-37", "C-91", "2B-4")
        unit_dropdown.grid(row=0, column=1, padx=10, pady=10)
        unit_dropdown.current(0)
        self.unit_var.trace('w', self.on_unit_change)
        tk.Label(top_card, text="Month:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.month_var = tk.StringVar()
        month_dropdown = ttk.Combobox(top_card, textvariable=self.month_var, state="readonly", width=12)
        month_dropdown['values'] = ("January", "February", "March", "April", "May", "June",
                                    "July", "August", "September", "October", "November", "December")
        month_dropdown.grid(row=0, column=3, padx=10, pady=10)
        month_dropdown.current(0)
        tk.Label(top_card, text="Current Date:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=4, padx=10, pady=10, sticky="w")
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.current_date_label = tk.Label(top_card, text=current_date, width=12, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10))
        self.current_date_label.grid(row=0, column=5, padx=10, pady=10)
        
        # Define fuel and refrigerant types BEFORE using them
        self.fuel_types = [
            {"name": "Diesel", "unit": "Liters", "factor": 2.54603},
            {"name": "Petrol", "unit": "Liters", "factor": 2.296},
            {"name": "PNG", "unit": "SCM", "factor": 2.02266},
            {"name": "LPG", "unit": "Liters", "factor": 1.55537}
        ]
        self.refrig_types = [
            {"name": "R-22", "unit": "kg", "factor": 1810},
            {"name": "R-410A", "unit": "kg", "factor": 2088}
        ]
        
        # Scope 1 Card: Fuel & Refrigerants
        scope1_card = create_card(self.main_frame.scrollable_frame, fill="both")
        tk.Label(scope1_card, text="Scope 1: Fuel & Refrigerant Entries", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 14, "bold")).pack(pady=10)
        scope1_container = tk.Frame(scope1_card, bg=CARD_COLOR)
        scope1_container.pack(pady=10, padx=10, fill="both")
        # Fuel Data Frame
        fuel_frame = tk.LabelFrame(scope1_container, text="Fuel Data", bg=CARD_COLOR, fg=TEXT_COLOR,
                                   font=(FONT_FAMILY, 12, "bold"), padx=10, pady=10)
        fuel_frame.pack(side="left", padx=10, pady=10, fill="both", expand=True)
        fuel_headers = ["Category", "Unit", "Factor", "Amount", "Total", "Upload Document"]
        for col, header in enumerate(fuel_headers):
            tk.Label(fuel_frame, text=header, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=col, padx=8, pady=8)
        self.fuel_amount_vars = {}
        self.fuel_total_labels = {}
        for i, fuel in enumerate(self.fuel_types, start=1):
            tk.Label(fuel_frame, text=fuel["name"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=i, column=0, padx=8, pady=8)
            tk.Label(fuel_frame, text=fuel["unit"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=i, column=1, padx=8, pady=8)
            factor_entry = tk.Entry(fuel_frame, width=10, font=(FONT_FAMILY, 10))
            factor_entry.insert(0, str(fuel["factor"]))
            factor_entry.config(state="readonly", readonlybackground=CARD_COLOR, fg=TEXT_COLOR)
            factor_entry.grid(row=i, column=2, padx=8, pady=8)
            amount_var = tk.StringVar()
            self.fuel_amount_vars[fuel["name"]] = amount_var
            num_entry = NumericEntry(fuel_frame, textvariable=amount_var, width=10, font=(FONT_FAMILY, 10))
            num_entry.grid(row=i, column=3, padx=8, pady=8)
            add_focus_effect(num_entry)
            total_label = tk.Label(fuel_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10))
            total_label.grid(row=i, column=4, padx=8, pady=8)
            self.fuel_total_labels[fuel["name"]] = total_label
            def callback_fuel(*args, fuel_name=fuel["name"], factor=fuel["factor"]):
                new_total = update_total_value(factor, self.fuel_amount_vars[fuel_name].get())
                self.fuel_total_labels[fuel_name].config(text=new_total)
            amount_var.trace("w", callback_fuel)
            file_var = tk.StringVar()
            self.fuel_file_vars[fuel["name"]] = file_var
            btn = tk.Button(fuel_frame, text="Upload",
                            command=lambda var=file_var, f=fuel: upload_document(var,
                                                     self.unit_var.get(),
                                                     self.current_date_label.cget("text"),
                                                     f["name"],
                                                     "Fuel",
                                                     self.controller.email),
                            bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10),
                            relief="raised", bd=2, padx=10, pady=4)
            btn.grid(row=i, column=5, padx=8, pady=8)
            add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)
        # Scope 1 Refrigerants
        refrig_frame = tk.LabelFrame(scope1_container, text="Refrigerants", bg=CARD_COLOR, fg=TEXT_COLOR,
                                     font=(FONT_FAMILY, 12, "bold"), padx=10, pady=10)
        refrig_frame.pack(side="right", padx=10, pady=10, fill="both", expand=True)
        refrig_headers = ["Category", "Unit", "Factor", "Amount", "Total", "Upload Document"]
        for col, header in enumerate(refrig_headers):
            tk.Label(refrig_frame, text=header, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=col, padx=8, pady=8)
        self.refrig_amount_vars = {}
        self.refrig_total_labels = {}
        for i, refrig in enumerate(self.refrig_types, start=1):
            tk.Label(refrig_frame, text=refrig["name"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=i, column=0, padx=8, pady=8)
            tk.Label(refrig_frame, text=refrig["unit"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=i, column=1, padx=8, pady=8)
            refrig_factor_var = tk.StringVar()
            refrig_factor_var.set(str(refrig["factor"]))
            factor_entry = tk.Entry(refrig_frame, textvariable=refrig_factor_var, width=10, font=(FONT_FAMILY, 10))
            factor_entry.grid(row=i, column=2, padx=8, pady=8)
            amount_var = tk.StringVar()
            self.refrig_amount_vars[refrig["name"]] = amount_var
            num_entry = NumericEntry(refrig_frame, textvariable=amount_var, width=10, font=(FONT_FAMILY, 10))
            num_entry.grid(row=i, column=3, padx=8, pady=8)
            add_focus_effect(num_entry)
            total_label = tk.Label(refrig_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10))
            total_label.grid(row=i, column=4, padx=8, pady=8)
            self.refrig_total_labels[refrig["name"]] = total_label
            def callback_refrig(*args, refrig_name=refrig["name"], factor_var=refrig_factor_var):
                try:
                    factor_val = float(factor_var.get())
                except:
                    factor_val = 0
                new_total = update_total_value(factor_val, self.refrig_amount_vars[refrig_name].get())
                self.refrig_total_labels[refrig_name].config(text=new_total)
            amount_var.trace("w", callback_refrig)
            file_var = tk.StringVar()
            self.refrig_file_vars[refrig["name"]] = file_var
            btn = tk.Button(refrig_frame, text="Upload",
                            command=lambda var=file_var, r=refrig: upload_document(var,
                                                      self.unit_var.get(),
                                                      self.current_date_label.cget("text"),
                                                      r["name"],
                                                      "Refrigerants",
                                                      self.controller.email),
                            bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10),
                            relief="raised", bd=2, padx=10, pady=4)
            btn.grid(row=i, column=5, padx=8, pady=8)
            add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)
        # Scope 2: Electricity
        scope2_card = create_card(self.main_frame.scrollable_frame)
        tk.Label(scope2_card, text="Scope 2: Electricity Entries", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 14, "bold")).pack(pady=10)
        elec_frame = tk.LabelFrame(scope2_card, text="Electricity Data", bg=CARD_COLOR, fg=TEXT_COLOR,
                                   font=(FONT_FAMILY, 12, "bold"), padx=10, pady=10)
        elec_frame.pack(pady=10, padx=10, fill="x")
        elec_headers = ["Category", "Type", "Unit", "Factor", "Amount", "Total", "Upload Document"]
        for col, header in enumerate(elec_headers):
            tk.Label(elec_frame, text=header, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=col, padx=8, pady=8)
        tk.Label(elec_frame, text="Electricity", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=1, column=0, padx=8, pady=8)
        tk.Label(elec_frame, text="India", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=1, column=1, padx=8, pady=8)
        tk.Label(elec_frame, text="kWh", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).grid(row=1, column=2, padx=8, pady=8)
        self.elec_factor = 0.6727
        factor_entry = tk.Entry(elec_frame, width=10, font=(FONT_FAMILY, 10))
        factor_entry.insert(0, str(self.elec_factor))
        factor_entry.config(state="readonly", readonlybackground=CARD_COLOR, fg=TEXT_COLOR)
        factor_entry.grid(row=1, column=3, padx=8, pady=8)
        self.elec_amount_var = tk.StringVar()
        elec_entry = NumericEntry(elec_frame, textvariable=self.elec_amount_var, width=10, font=(FONT_FAMILY, 10))
        elec_entry.grid(row=1, column=4, padx=8, pady=8)
        add_focus_effect(elec_entry)
        elec_total_label = tk.Label(elec_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10))
        elec_total_label.grid(row=1, column=5, padx=8, pady=8)
        def callback_elec(*args):
            new_total = update_total_value(self.elec_factor, self.elec_amount_var.get())
            elec_total_label.config(text=new_total)
        self.elec_amount_var.trace("w", callback_elec)
        self.elec_file_var = tk.StringVar()
        btn = tk.Button(elec_frame, text="Upload", command=lambda var=self.elec_file_var: upload_document(var,
                                             self.unit_var.get(),
                                             self.current_date_label.cget("text"),
                                             "Electricity",
                                             "Electricity",
                                             self.controller.email),
                        bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10),
                        relief="raised", bd=2, padx=10, pady=4)
        btn.grid(row=1, column=6, padx=8, pady=8)
        add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)
        # Scope 3: Reserved for Future
        scope3_card = create_card(self.main_frame.scrollable_frame)
        tk.Label(scope3_card, text="Scope 3: Reserved for Future Edits", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 14, "bold")).pack(pady=10)
        tk.Label(scope3_card, text="Reserved for future enhancements", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 12)).pack(pady=10)
        # Navigation Buttons
        btn_frame = tk.Frame(self.main_frame.scrollable_frame, bg=BACKGROUND_COLOR)
        btn_frame.pack(pady=20)
        btn_submit = tk.Button(btn_frame, text="Submit", command=self.submit_data_handler,
                               bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"),
                               relief="raised", bd=2, padx=20, pady=10)
        btn_submit.pack(side="left", padx=10)
        add_hover(btn_submit, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_go_emission = tk.Button(btn_frame, text="Go to Emission Data", command=lambda: self.controller.show_frame("EmissionDataPage"),
                                    bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"),
                                    relief="raised", bd=2, padx=20, pady=10)
        btn_go_emission.pack(side="left", padx=10)
        add_hover(btn_go_emission, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_back = tk.Button(btn_frame, text="Back to Home", command=lambda: self.controller.show_frame("HomePage"),
                             bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"),
                             relief="raised", bd=2, padx=20, pady=10)
        btn_back.pack(side="left", padx=10)
        add_hover(btn_back, DANGER_COLOR, DANGER_HOVER)
    
    def on_unit_change(self, *args):
        self.reset_input_fields()
    
    def reset_input_fields(self):
        for key in self.fuel_amount_vars:
            self.fuel_amount_vars[key].set("")
            self.fuel_total_labels[key].config(text="0.00")
            self.fuel_file_vars[key].set("")
        for key in self.refrig_amount_vars:
            self.refrig_amount_vars[key].set("")
            self.refrig_total_labels[key].config(text="0.00")
            self.refrig_file_vars[key].set("")
        self.elec_amount_var.set("")
    
    def submit_data_handler(self):
        try:
            unit = self.unit_var.get().strip()
            month = self.month_var.get().strip()
            entry_date = self.current_date_label.cget("text")
            user_email = self.controller.email
            if not unit or not month or not entry_date:
                messagebox.showerror("Mandatory Fields Missing", "Please fill out all common fields.")
                return
            for fuel in self.fuel_types:
                amount = self.fuel_amount_vars[fuel["name"]].get().strip()
                file_path = self.fuel_file_vars.get(fuel["name"], tk.StringVar()).get()
                if amount and (file_path == "" or file_path == "No File"):
                    messagebox.showerror("Document Missing", f"Please upload a document for {fuel['name']}.")
                    return
            for refrig in self.refrig_types:
                amount = self.refrig_amount_vars[refrig["name"]].get().strip()
                file_path = self.refrig_file_vars.get(refrig["name"], tk.StringVar()).get()
                if amount and (file_path == "" or file_path == "No File"):
                    messagebox.showerror("Document Missing", f"Please upload a document for {refrig['name']}.")
                    return
            elec_amount = self.elec_amount_var.get().strip()
            if elec_amount and (self.elec_file_var.get() == "" or self.elec_file_var.get() == "No File"):
                messagebox.showerror("Document Missing", "Please upload a document for Electricity.")
                return
            new_records = []
            global record_id_counter
            for fuel in self.fuel_types:
                amount = self.fuel_amount_vars[fuel["name"]].get().strip()
                if amount:
                    total = self.fuel_total_labels[fuel["name"]].cget("text")
                    file_path = self.fuel_file_vars.get(fuel["name"], tk.StringVar()).get()
                    record = (user_email, entry_date, month, unit, "Fuel", fuel["name"],
                              f"{fuel['factor']}", amount, total, file_path if file_path else "No File", record_id_counter)
                    record_id_counter += 1
                    new_records.append(record)
            for refrig in self.refrig_types:
                amount = self.refrig_amount_vars[refrig["name"]].get().strip()
                if amount:
                    total = self.refrig_total_labels[refrig["name"]].cget("text")
                    file_path = self.refrig_file_vars.get(refrig["name"], tk.StringVar()).get()
                    record = (user_email, entry_date, month, unit, "Refrigerants", refrig["name"],
                              f"{refrig['factor']}", amount, total, file_path if file_path else "No File", record_id_counter)
                    record_id_counter += 1
                    new_records.append(record)
            if elec_amount:
                total = update_total_value(self.elec_factor, elec_amount)
                file_path = self.elec_file_var.get()
                record = (user_email, entry_date, month, unit, "Electricity", "Electricity",
                          f"{self.elec_factor}", elec_amount, total, file_path if file_path else "No File", record_id_counter)
                record_id_counter += 1
                new_records.append(record)
            if new_records:
                emission_records.extend(new_records)
                save_emission_records()
                logging.info(f"Data submitted for user {user_email}: {new_records}")
                messagebox.showinfo("Data Submitted", "Data submitted successfully!")
                self.reset_input_fields()
                if "EmissionDataPage" in self.controller.frames:
                    self.controller.frames["EmissionDataPage"].refresh_table()
            else:
                messagebox.showwarning("No Data", "No emission data entered. Please enter some values before submitting.")
        except Exception as e:
            logging.error(f"Error in data submission: {e}")
            messagebox.showerror("Submission Error", f"An error occurred during submission: {e}")

if __name__ == "__main__":
    class MainApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("RMX Joss Carbon Tracking System")
            self.geometry("1100x900")
            self.email = None
            container = tk.Frame(self)
            container.pack(side="top", fill="both", expand=True)
            container.grid_rowconfigure(0, weight=1)
            container.grid_columnconfigure(0, weight=1)
            self.frames = {}
            for F in (LoginPage, HomePage, DataEntryPage, EmissionDataPage, AnalysisPage):
                page_name = F.__name__
                frame = F(parent=container, controller=self)
                self.frames[page_name] = frame
                frame.grid(row=0, column=0, sticky="nsew")
            load_emission_records()
            self.show_frame("LoginPage")
        
        def show_frame(self, page_name):
            frame = self.frames[page_name]
            if hasattr(frame, "update_role_buttons"):
                frame.update_role_buttons()
            if hasattr(frame, "user_label"):
                frame.user_label.config(text=f"User: {self.email}")
            if page_name == "EmissionDataPage":
                frame.refresh_table()
            frame.tkraise()
        
        def logout(self):
            self.email = None
            self.frames["LoginPage"].reset()
            self.show_frame("LoginPage")
    
    class LoginPage(tk.Frame):
        def __init__(self, parent, controller):
            super().__init__(parent, bg=BACKGROUND_COLOR)
            self.controller = controller
            frame = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
            frame.place(relx=0.5, rely=0.5, anchor="center", width=300, height=250)
            tk.Label(frame, text="Login to RMX Joss System", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 14, "bold")).pack(pady=10)
            tk.Label(frame, text="Email:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).pack(pady=5)
            self.email_entry = tk.Entry(frame, width=30)
            self.email_entry.pack(pady=5)
            tk.Label(frame, text="Password:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 10)).pack(pady=5)
            self.password_entry = tk.Entry(frame, show="*", width=30)
            self.password_entry.pack(pady=5)
            btn_login = tk.Button(frame, text="Login", command=self.login, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 10, "bold"))
            btn_login.pack(pady=10)
            add_hover(btn_login, PRIMARY_COLOR, PRIMARY_HOVER)
        
        def login(self):
            email = self.email_entry.get().strip()
            password = self.password_entry.get().strip()
            if email in ["employee@gmail.com", "manager@gmail.com"] and ((email == "employee@gmail.com" and password == "1234") or (email == "manager@gmail.com" and password == "admin")):
                logging.info(f"User {email} logged in successfully.")
                self.controller.email = email
                self.controller.show_frame("HomePage")
            else:
                messagebox.showerror("Login Failed", "Invalid credentials.")
        
        def reset(self):
            self.email_entry.delete(0, tk.END)
            self.password_entry.delete(0, tk.END)
    
    class HomePage(tk.Frame):
        def __init__(self, parent, controller):
            super().__init__(parent, bg=BACKGROUND_COLOR)
            self.controller = controller
            card = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
            card.place(relx=0.5, rely=0.5, anchor="center", width=500, height=400)
            tk.Label(card, text="Welcome to RMX Joss Carbon Tracking System", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 16, "bold")).pack(pady=20)
            self.user_label = tk.Label(card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY, 12))
            self.user_label.pack(pady=10)
            btn_data = tk.Button(card, text="Data Entry", command=lambda: controller.show_frame("DataEntryPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
            btn_data.pack(pady=10)
            add_hover(btn_data, PRIMARY_COLOR, PRIMARY_HOVER)
            btn_emission = tk.Button(card, text="Emission Data", command=lambda: controller.show_frame("EmissionDataPage"),
                                     bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
            btn_emission.pack(pady=10)
            add_hover(btn_emission, PRIMARY_COLOR, PRIMARY_HOVER)
            btn_analysis = tk.Button(card, text="Analysis", command=lambda: controller.show_frame("AnalysisPage"),
                                     bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
            btn_analysis.pack(pady=10)
            add_hover(btn_analysis, PRIMARY_COLOR, PRIMARY_HOVER)
            btn_logout = tk.Button(card, text="Logout", command=lambda: controller.logout(),
                                   bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY, 12, "bold"), width=20)
            btn_logout.pack(pady=10)
            add_hover(btn_logout, DANGER_COLOR, DANGER_HOVER)
        
        def tkraise(self, aboveThis=None):
            self.user_label.config(text=f"Logged in as: {self.controller.email}")
            super().tkraise(aboveThis)
    
    app = MainApp()
    app.mainloop()
