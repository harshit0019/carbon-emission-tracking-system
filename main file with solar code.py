import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import logging
import os
import shutil
import json
import pandas as pd
import numpy as np
from openpyxl import Workbook
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import pyodbc
from tkcalendar import DateEntry

# ---------------- Global Constants (Theming) ----------------
BACKGROUND_COLOR = "#F8F9FA"
CARD_COLOR       = "white"
PRIMARY_COLOR    = "#3498DB"
PRIMARY_HOVER    = "#2980B9"
DANGER_COLOR     = "#E74C3C"
DANGER_HOVER     = "#c0392b"
TEXT_COLOR       = "#2C3E50"
SHADOW_COLOR     = "#d3d3d3"
FONT_FAMILY      = "Arial"

# ---------------- Logging Setup ----------------
logging.basicConfig(
    filename="app.log",
    level=logging.INFO,
    format="%(asctime)s:%(levelname)s:%(message)s"
)

# ---------------- Global Data Storage ----------------
# Schema: [Email, Entry Date, Month, Year, Unit, Emission Category,
#          Emission Name, Emission Type, Factor, Value, Total, Remarks, Document, RecordID]
emission_records = []
document_logs    = []
record_id_counter= 0

# ---------------- System Configuration ----------------
system_config = {
    "company_name": "RMX Joss",
    "document_base": r"C:\Users\Public\RMXJoss",
    "units": ["C-49", "B-37", "C-91", "2B-4"],
    "scope_factors": {
        "Fuel": {"Diesel": 2.54603, "Petrol": 2.296, "PNG": 2.02266, "LPG": 1.55537},
        "Refrigerants": {"R-22": 1810, "R-410A": 2088},
        "Electricity": {"Electricity": 0.6727}
    },
    "users": {
        "admin": {"email": "admin@gmail.com", "password": "admin123", "role": "Admin"},
        "manager": [],
        "employee": []
    },
    "database": {
        "type": "MSSQL",
        "mssql": {
            "server": "DESKTOP-GU12JEU",
            "database": "RMXCarbonDB",
            "user": "sa",
            "password": "sa@1234"
        }
    }
}

# ---------------- Config Persistence ----------------
CONFIG_FILE = os.path.join(system_config["document_base"], "config.json")

def save_config():
    try:
        os.makedirs(system_config["document_base"], exist_ok=True)
        with open(CONFIG_FILE, "w") as f:
            json.dump(system_config, f, indent=4)
        logging.info("Configuration saved.")
    except Exception as e:
        logging.error("Error saving config: " + str(e))

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                loaded = json.load(f)
            system_config.update(loaded)
            logging.info("Configuration loaded.")
        else:
            save_config()
    except Exception as e:
        logging.error("Error loading config: " + str(e))

# ---------------- Helpers ----------------
def update_total_value(factor, value_str):
    try:
        v = float(value_str)
        return f"{factor * v:.2f}"
    except:
        return "0.00"

def get_user_role(email):
    if email == system_config["users"]["admin"]["email"]:
        return "Admin"
    for u in system_config["users"]["manager"]:
        if u["email"] == email:
            return "Manager"
    return "Employee"

# ---------------- MSSQL Connection ----------------
def connect_mssql(server, database, user, password):
    try:
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={server};DATABASE={database};UID={user};PWD={password};"
            "Trusted_Connection=no;"
            "TrustServerCertificate=yes;"
            "Connection Timeout=30;"
        )
        conn = pyodbc.connect(conn_str)
        logging.info("Connected to MSSQL")
        return conn
    except Exception as e:
        logging.error(f"MSSQL connection error: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to connect to database: {str(e)}")
        return None

def init_db():
    cfg = system_config["database"]["mssql"]
    conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
    if not conn:
        return
    create_sql = """
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='EmissionRecords' AND xtype='U')
    BEGIN
      CREATE TABLE EmissionRecords (
        record_id INT IDENTITY(1,1) PRIMARY KEY,
        email NVARCHAR(255) NOT NULL,
        entry_date DATE NOT NULL,
        month NVARCHAR(50) NOT NULL,
        year NVARCHAR(10) NOT NULL,
        company_unit NVARCHAR(50) NOT NULL,
        emission_category NVARCHAR(50) NOT NULL,
        emission_name NVARCHAR(50) NOT NULL,
        emission_type NVARCHAR(50) NOT NULL,
        factor NUMERIC(18,4) NOT NULL,
        value NUMERIC(18,4) NOT NULL,
        total NUMERIC(18,4) NOT NULL,
        remarks NVARCHAR(500),
        document NVARCHAR(500),
        created_at DATETIME DEFAULT GETDATE(),
        updated_at DATETIME DEFAULT GETDATE()
      );
    END

    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='solar_energy_data' AND xtype='U')
    BEGIN
      CREATE TABLE solar_energy_data (
        id INT IDENTITY(1,1) PRIMARY KEY,
        gmail NVARCHAR(255) NOT NULL,
        entry_date DATE NOT NULL,
        month NVARCHAR(50) NOT NULL,
        year NVARCHAR(10) NOT NULL,
        company_unit NVARCHAR(50) NOT NULL,
        date DATE NOT NULL,
        inverter1 NUMERIC(18,4) NOT NULL DEFAULT 0,
        inverter2 NUMERIC(18,4) NOT NULL DEFAULT 0,
        inverter3 NUMERIC(18,4) NOT NULL DEFAULT 0,
        inverter4 NUMERIC(18,4) NOT NULL DEFAULT 0,
        old_total NUMERIC(18,4) NOT NULL DEFAULT 0,
        new_solar_inverter NUMERIC(18,4) NOT NULL DEFAULT 0,
        total_generated NUMERIC(18,4) NOT NULL DEFAULT 0,
        unit_type NVARCHAR(50) NOT NULL DEFAULT 'Kwh',
        remark NVARCHAR(500),
        document NVARCHAR(500),
        created_at DATETIME DEFAULT GETDATE(),
        updated_at DATETIME DEFAULT GETDATE()
      );
    END

    -- Add indexes for better performance
    IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_EmissionRecords_EntryDate' AND object_id = OBJECT_ID('EmissionRecords'))
    BEGIN
        CREATE INDEX IX_EmissionRecords_EntryDate ON EmissionRecords(entry_date);
    END

    IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SolarEnergyData_EntryDate' AND object_id = OBJECT_ID('solar_energy_data'))
    BEGIN
        CREATE INDEX IX_SolarEnergyData_EntryDate ON solar_energy_data(entry_date);
    END
    """
    try:
        cur = conn.cursor()
        cur.execute(create_sql)
        conn.commit()
        logging.info("DB initialized successfully")
    except Exception as e:
        logging.error(f"DB init error: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to initialize database: {str(e)}")
    finally:
        conn.close()

def save_emission_records():
    cfg = system_config["database"]["mssql"]
    conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
    if not conn:
        return
    try:
        cur = conn.cursor()
        insert_sql = """
        INSERT INTO EmissionRecords
          (email, entry_date, month, year, company_unit, emission_category,
           emission_name, emission_type, factor, value, total, remarks, document)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """
        for r in emission_records:
            try:
                # Convert date string to proper format
                entry_date = datetime.strptime(r[1], "%Y-%m-%d").strftime("%Y-%m-%d")
                params = (
                    r[0], entry_date, r[2], r[3], r[4], r[5],
                    r[6], r[7], float(r[8]), float(r[9]),
                    float(r[10]), r[11], r[12]
                )
                cur.execute(insert_sql, params)
            except Exception as e:
                logging.error(f"Error inserting record: {str(e)}")
                continue
        conn.commit()
        logging.info("Records saved successfully")
    except Exception as e:
        logging.error(f"Save records error: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to save records: {str(e)}")
    finally:
        conn.close()

def load_emission_records():
    global emission_records, record_id_counter
    cfg = system_config["database"]["mssql"]
    conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
    if not conn:
        return
    try:
        cur = conn.cursor()
        cur.execute("""
        SELECT email, entry_date, month, year, company_unit, emission_category,
               emission_name, emission_type, factor, value, total, remarks, document, record_id
        FROM EmissionRecords
        ORDER BY entry_date DESC, record_id DESC;
        """)
        rows = cur.fetchall()
        emission_records.clear()
        for row in rows:
            try:
                emission_records.append((
                    row[0],
                    row[1].strftime("%Y-%m-%d"),
                    row[2], row[3], row[4], row[5],
                    row[6], row[7],
                    str(row[8]), str(row[9]), str(row[10]),
                    row[11] or "", row[12], str(row[13])
                ))
            except Exception as e:
                logging.error(f"Error processing record: {str(e)}")
                continue
        if emission_records:
            record_id_counter = max(int(r[13]) for r in emission_records) + 1
        logging.info("Records loaded successfully")
    except Exception as e:
        logging.error(f"Load records error: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to load records: {str(e)}")
    finally:
        conn.close()

# ---------------- Document Management ----------------
class DocumentManagementSystem:
    BASE_DIR = system_config["document_base"]

    @staticmethod
    def generate_unique_code(unit_name, upload_date, emission_name, emission_type):
        dt = datetime.strptime(upload_date, "%Y-%m-%d")
        return f"{unit_name}_{dt.strftime('%d_%m_%Y')}_{emission_name}_{emission_type}"

    @staticmethod
    def get_storage_path(unit_name, upload_date):
        dt = datetime.strptime(upload_date, "%Y-%m-%d")
        folder_path = os.path.join(
            DocumentManagementSystem.BASE_DIR,
            unit_name,
            dt.strftime("%Y"),
            dt.strftime("%m_%B")
        )
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    @staticmethod
    def save_document(fp, unit_name, upload_date, emission_name, emission_type, uploader, role):
        code = DocumentManagementSystem.generate_unique_code(unit_name, upload_date, emission_name, emission_type)
        folder = DocumentManagementSystem.get_storage_path(unit_name, upload_date)
        ext = os.path.splitext(fp)[1]
        dest = os.path.join(folder, f"{code}{ext}")
        version = 1
        final = dest
        while os.path.exists(final):
            version += 1
            final = os.path.join(folder, f"{code}_v{version}{ext}")
        shutil.copy(fp, final)
        meta = {
            "unique_code": code,
            "file_path": final,
            "upload_date_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "uploader": uploader,
            "role": role,
            "unit_name": unit_name,
            "upload_date": upload_date,
            "emission_name": emission_name,
            "emission_type": emission_type,
            "file_status": "Stored Locally",
            "version": version
        }
        document_logs.append(meta)
        logging.info("Doc saved: " + json.dumps(meta, indent=4))
        return meta

def upload_document(var, unit, date_str, name, etype, uploader):
    fp = filedialog.askopenfilename(
        filetypes=[
            ("All files","*.*"),
            ("PDF","*.pdf"),
            ("Excel","*.xlsx;*.xls"),
            ("Images","*.png;*.jpg;*.jpeg")
        ],
        title="Select document"
    )
    if not fp:
        return
    role = "Admin" if uploader == system_config["users"]["admin"]["email"] else "User"
    meta = DocumentManagementSystem.save_document(fp, unit, date_str, name, etype, uploader, role)
    var.set(meta["file_path"])
    messagebox.showinfo("File Saved", f"Saved to:\n{meta['file_path']}")

# ---------------- Numeric Entry ----------------
class NumericEntry(tk.Entry):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        vcmd = (self.register(self._val), '%P')
        self.config(validate="key", validatecommand=vcmd)
    def _val(self, P):
        if P=="":
            self.config(bg="white")
            return True
        try:
            float(P)
            self.config(bg="white")
            return True
        except:
            self.config(bg="#ffcccc")
            return False

# ---------------- Hover & Focus ----------------
def add_hover(w, nbg, hbg):
    w.bind("<Enter>", lambda e: w.config(bg=hbg))
    w.bind("<Leave>", lambda e: w.config(bg=nbg))

def add_focus_effect(entry, nbg="white", fbg="#e0f7fa"):
    entry.bind("<FocusIn>", lambda e: entry.config(bg=fbg))
    entry.bind("<FocusOut>", lambda e: entry.config(bg=nbg))

# ---------------- Scrollable Frame ----------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *a, **k):
        super().__init__(container, *a, **k)
        canvas = tk.Canvas(self, bg=BACKGROUND_COLOR, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=BACKGROUND_COLOR)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

# ---------------- Card Utility ----------------
def create_card(parent, pady=15, padx=20, fill="x"):
    shadow = tk.Frame(parent, bg=SHADOW_COLOR)
    shadow.pack(pady=pady, padx=padx, fill=fill)
    card = tk.Frame(shadow, bg=CARD_COLOR)
    card.pack(padx=3, pady=3, fill=fill)
    return card

# ---------------- AdminPage ----------------
class AdminPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        tk.Label(self, text="Admin Panel", font=(FONT_FAMILY,24,"bold"), bg=BACKGROUND_COLOR, fg=TEXT_COLOR).pack(pady=10)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_company = tk.Frame(nb, bg=BACKGROUND_COLOR)
        nb.add(self.tab_company, text="Company Setup")
        self.build_company(self.tab_company)
        self.tab_role = tk.Frame(nb, bg=BACKGROUND_COLOR)
        nb.add(self.tab_role, text="Roles & Users")
        self.build_roles(self.tab_role)
        self.tab_db = tk.Frame(nb, bg=BACKGROUND_COLOR)
        nb.add(self.tab_db, text="DB Connection")
        self.build_db(self.tab_db)
        btn_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Save Settings", command=self.save_settings,
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), padx=15, pady=8).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Back to Home", command=lambda: controller.show_frame("HomePage"),
                  bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), padx=15, pady=8).pack(side="left", padx=10)

    def build_company(self, p):
        card = create_card(p)
        tk.Label(card, text="Company Name:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12)).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.company_name_var = tk.StringVar(value=system_config["company_name"])
        tk.Entry(card, textvariable=self.company_name_var, width=30).grid(row=0, column=1, padx=10, pady=10)

    def build_roles(self, p):
        top = tk.Frame(p, bg=BACKGROUND_COLOR)
        top.pack(pady=10)
        tk.Label(top, text="Add User Account", bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,16,"bold")).pack()
        card = create_card(p)
        frm = tk.Frame(card, bg=CARD_COLOR)
        frm.pack(pady=5)
        tk.Label(frm, text="Role:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12)).grid(row=0, column=0, padx=5)
        self.new_role_var = tk.StringVar(value="Manager")
        ttk.Combobox(frm, textvariable=self.new_role_var, values=["Manager","Employee"], state="readonly", width=10).grid(row=0, column=1, padx=5)
        tk.Label(frm, text="Email:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12)).grid(row=0, column=2, padx=5)
        self.new_email_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.new_email_var, width=25).grid(row=0, column=3, padx=5)
        tk.Label(frm, text="Password:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12)).grid(row=0, column=4, padx=5)
        self.new_pass_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.new_pass_var, width=20, show="*").grid(row=0, column=5, padx=5)
        tk.Button(frm, text="Add", command=self.add_user, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).grid(row=0, column=6, padx=5)
        table_card = create_card(p)
        tk.Label(table_card, text="Current Users", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,14,"bold")).pack(pady=5)
        self.users_tree = ttk.Treeview(table_card, columns=("Role","Email","Password"), show="headings", height=5)
        for c in ("Role", "Email", "Password"):
            self.users_tree.heading(c, text=c)
            self.users_tree.column(c, width=150, anchor="center")
        self.users_tree.pack(padx=10, pady=5, fill="x")
        af = tk.Frame(table_card, bg=CARD_COLOR)
        af.pack(pady=5)
        tk.Button(af, text="Edit", command=self.edit_user, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).pack(side="left", padx=10)
        tk.Button(af, text="Delete", command=self.delete_user, bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).pack(side="left", padx=10)
        self.refresh_users_table()

    def add_user(self):
        role, email, pwd = self.new_role_var.get(), self.new_email_var.get().strip(), self.new_pass_var.get().strip()
        if not email or not pwd:
            messagebox.showerror("Error", "Enter email & password.")
            return
        lst = system_config["users"]["manager" if role=="Manager" else "employee"]
        lst.append({"email": email, "password": pwd, "role": role})
        self.new_email_var.set("")
        self.new_pass_var.set("")
        self.refresh_users_table()

    def refresh_users_table(self):
        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        allu = []
        for u in system_config["users"]["manager"]:
            allu.append(("Manager", u["email"], u["password"]))
        for u in system_config["users"]["employee"]:
            allu.append(("Employee", u["email"], u["password"]))
        for i, data in enumerate(allu):
            self.users_tree.insert("", "end", iid=str(i), values=data)

    def delete_user(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showerror("Error", "No user selected.")
            return
        idx = int(sel[0])
        allu = []
        for u in system_config["users"]["manager"]:
            allu.append(("Manager", u))
        for u in system_config["users"]["employee"]:
            allu.append(("Employee", u))
        role, data = allu[idx]
        system_config["users"]["manager" if role=="Manager" else "employee"].remove(data)
        self.refresh_users_table()

    def edit_user(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showerror("Error", "No user selected.")
            return
        idx = int(sel[0])
        allu = []
        for u in system_config["users"]["manager"]:
            allu.append(("Manager", u))
        for u in system_config["users"]["employee"]:
            allu.append(("Employee", u))
        role, data = allu[idx]
        ew = tk.Toplevel(self)
        ew.title("Edit User")
        tk.Label(ew, text="Role:").grid(row=0, column=0, padx=5, pady=5)
        rv = tk.StringVar(value=role)
        ttk.Combobox(ew, textvariable=rv, values=["Manager", "Employee"], state="readonly", width=10).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(ew, text="Email:").grid(row=1, column=0, padx=5, pady=5)
        ev = tk.StringVar(value=data["email"])
        tk.Entry(ew, textvariable=ev, width=25).grid(row=1, column=1, padx=5, pady=5)
        tk.Label(ew, text="Password:").grid(row=2, column=0, padx=5, pady=5)
        pv = tk.StringVar(value=data["password"])
        tk.Entry(ew, textvariable=pv, width=20, show="*").grid(row=2, column=1, padx=5, pady=5)
        def save_edit():
            nr, ne, npw = rv.get(), ev.get().strip(), pv.get().strip()
            for lst in (system_config["users"]["manager"], system_config["users"]["employee"]):
                for u in lst:
                    if u["email"] == data["email"]:
                        lst.remove(u)
                        break
            system_config["users"]["manager" if nr=="Manager" else "employee"].append({"email": ne, "password": npw, "role": nr})
            self.refresh_users_table()
            ew.destroy()
        tk.Button(ew, text="Save", command=save_edit, bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).grid(row=3, column=0, columnspan=2, pady=10)

    def build_db(self, p):
        card = create_card(p)
        tk.Label(card, text="MSSQL Connection", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,14,"bold")).grid(row=0, column=0, columnspan=2, pady=10)
        labels = ["Server:", "Database:", "User:", "Password:"]
        vars   = ["server","database","user","password"]
        for i, lbl in enumerate(labels, start=1):
            tk.Label(card, text=f"MSSQL - {lbl}", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12)).grid(row=i, column=0, padx=10, pady=5, sticky="e")
            var = tk.StringVar(value=system_config["database"]["mssql"].get(vars[i-1], ""))
            setattr(self, f"mssql_{vars[i-1]}_var", var)
            show = "" if vars[i-1]!="password" else "*"
            tk.Entry(card, textvariable=var, width=30, show=show).grid(row=i, column=1, padx=10, pady=5, sticky="w")

    def save_settings(self):
        if self.controller.email != system_config["users"]["admin"]["email"]:
            messagebox.showerror("Denied", "Only admin may update.")
            return
        system_config["company_name"] = self.company_name_var.get().strip()
        system_config["database"]["mssql"] = {
            "server": self.mssql_server_var.get().strip(),
            "database": self.mssql_database_var.get().strip(),
            "user": self.mssql_user_var.get().strip(),
            "password": self.mssql_password_var.get().strip()
        }
        save_config()
        messagebox.showinfo("Saved", "Settings updated.")

    def tkraise(self, above=None):
        if self.controller.email != system_config["users"]["admin"]["email"]:
            messagebox.showerror("Denied", "No permission.")
            self.controller.show_frame("HomePage")
            return
        self.refresh_users_table()
        super().tkraise(above)

# ---------------- AnalysisPage ----------------
class AnalysisPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller

        # Make page scrollable
        scroll = ScrollableFrame(self)
        scroll.pack(fill="both", expand=True)
        container = scroll.scrollable_frame

        tk.Label(container, text="RMX Joss Carbon Tracking System", font=(FONT_FAMILY,24,"bold"),
                 bg=BACKGROUND_COLOR, fg=TEXT_COLOR).pack(pady=10)

        # KPI Cards
        self.kpi_frame = tk.Frame(container, bg=BACKGROUND_COLOR)
        self.kpi_frame.pack(fill="x", padx=10, pady=5)
        self.kpi_total    = self._make_kpi("Total Emissions", "0.00 kg CO2e")
        self.kpi_scope1   = self._make_kpi("Scope 1 Emissions", "0.00 kg CO2e")
        self.kpi_scope2   = self._make_kpi("Scope 2 Emissions", "0.00 kg CO2e")
        self.kpi_top_unit = self._make_kpi("Top Emitting Company Unit", "N/A")
        self.kpi_top_type = self._make_kpi("Top Emitting Type", "N/A")
        self.kpi_top_cat  = self._make_kpi("Top Emission Category", "N/A")
        for c in (self.kpi_total, self.kpi_scope1, self.kpi_scope2,
                  self.kpi_top_unit, self.kpi_top_type, self.kpi_top_cat):
            c.pack(side="left", padx=10, pady=5)

        # Filters
        self.filter_frame = tk.Frame(container, bg=BACKGROUND_COLOR)
        self.filter_frame.pack(pady=10)
        self.filter_vars = {}
        col = 0
        flts = [
            ("Company Unit:", "unit", ["All"] + system_config["units"]),
            ("Year:", "year", ["All"] + [str(y) for y in range(2020, 2032)]),
            ("Month:", "month", ["All", "January", "February", "March", "April", "May", "June",
                                  "July", "August", "September", "October", "November", "December"]),
            ("Category:", "emission_category", ["All", "Scope1", "Scope2", "Scope3"]),
            ("Name:", "emission_name", ["All", "Fuel", "Refrigerants", "Electricity"]),
            ("Type:", "emission_type", ["All", "Diesel", "Petrol", "PNG", "LPG", "R-22", "R-410A", "Electricity"])
        ]
        for lbl, key, opts in flts:
            tk.Label(self.filter_frame, text=lbl, bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=col, sticky="e")
            var = tk.StringVar(value="All")
            self.filter_vars[key] = var
            cb = ttk.Combobox(self.filter_frame, textvariable=var, values=opts, state="readonly", width=10)
            cb.grid(row=0, column=col+1, padx=3)
            cb.bind("<<ComboboxSelected>>", lambda e: self.update_analysis())
            col += 2

        # Y-Axis selector
        tk.Label(self.filter_frame, text="Y-Axis:", bg=BACKGROUND_COLOR, fg=TEXT_COLOR).grid(row=0, column=col, sticky="e")
        self.filter_vars["y_axis"] = tk.StringVar(value="Total")
        ycb = ttk.Combobox(self.filter_frame, textvariable=self.filter_vars["y_axis"],
                           values=["Total", "Value"], state="readonly", width=10)
        ycb.grid(row=0, column=col+1, padx=3)
        ycb.bind("<<ComboboxSelected>>", lambda e: self.update_analysis())

        # Reset Filters button
        reset_btn = tk.Button(self.filter_frame, text="Reset Filters", command=self.reset_filters,
                              bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"), padx=10, pady=5)
        reset_btn.grid(row=0, column=col+2, padx=10)
        add_hover(reset_btn, PRIMARY_COLOR, PRIMARY_HOVER)

        # Layout Panes: Charts, Summary, Table
        row_frame = tk.Frame(container, bg=BACKGROUND_COLOR)
        row_frame.pack(fill="both", expand=True, padx=10, pady=10)
        row_frame.grid_columnconfigure(0, weight=3)
        row_frame.grid_columnconfigure(1, weight=1)
        row_frame.grid_columnconfigure(2, weight=1)

        # Charts Notebook
        self.chart_nb = ttk.Notebook(row_frame)
        self.chart_nb.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        titles = ["Monthly by Type", "Distribution by Unit", "Yearly by Category", "Yearly by Type", "Yearly by Unit"]
        for idx, title in enumerate(titles, start=1):
            tab = tk.Frame(self.chart_nb, bg=BACKGROUND_COLOR)
            self.chart_nb.add(tab, text=title)
            cf = create_card(tab, pady=5, padx=5, fill="both")
            fig = Figure(figsize=(6,4), dpi=100)
            ax = fig.add_subplot(111)
            canvas = FigureCanvasTkAgg(fig, master=cf)
            canvas.get_tk_widget().pack(fill="both", expand=True)
            setattr(self, f"fig{idx}", fig)
            setattr(self, f"ax{idx}", ax)
            setattr(self, f"canvas{idx}", canvas)

        # Summary Notebook
        self.summary_nb = ttk.Notebook(row_frame)
        self.summary_nb.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        summ_tab = tk.Frame(self.summary_nb, bg=BACKGROUND_COLOR)
        self.summary_nb.add(summ_tab, text="Insights")
        sf = create_card(summ_tab, pady=5, padx=5, fill="both")
        self.summary_label = tk.Label(sf, text="", justify="left", bg=CARD_COLOR, fg=TEXT_COLOR, wraplength=300)
        self.summary_label.pack(anchor="nw", pady=5)

        # Table Notebook
        self.table_nb = ttk.Notebook(row_frame)
        self.table_nb.grid(row=0, column=2, sticky="nsew", padx=5, pady=5)
        table_tab = tk.Frame(self.table_nb, bg=BACKGROUND_COLOR)
        self.table_nb.add(table_tab, text="Records")
        tf = create_card(table_tab, pady=5, padx=5, fill="both")
        cols = ("Emission Name", "Emission Type", "Total", "Document")
        self.table = ttk.Treeview(tf, columns=cols, show="headings", height=12)
        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, anchor="center", width=120)
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=vsb.set)
        self.table.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.table.bind("<Double-1>", lambda e: self._open_document(e))

        # Navigation Buttons
        nav = tk.Frame(container, bg=BACKGROUND_COLOR)
        nav.pack(pady=10)
        for txt, cmd in [("Home", lambda: controller.show_frame("HomePage")),
                         ("Data Entry", lambda: controller.show_frame("DataEntryPage")),
                         ("Emissions", lambda: controller.show_frame("EmissionDataPage"))]:
            btn = tk.Button(nav, text=txt, command=cmd, bg=PRIMARY_COLOR, fg="white",
                            font=(FONT_FAMILY,12,"bold"), padx=15, pady=8)
            btn.pack(side="left", padx=5)
            add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)

        self.update_analysis()

    def _make_kpi(self, title, val):
        frm = tk.Frame(self.kpi_frame, bg="white", bd=2, relief="solid", padx=10, pady=5)
        tk.Label(frm, text=title, font=(FONT_FAMILY,10,"bold"), bg="white", fg=TEXT_COLOR).pack()
        lbl = tk.Label(frm, text=val, font=(FONT_FAMILY,12), bg="white", fg=TEXT_COLOR)
        lbl.pack()
        frm.lbl = lbl
        return frm

    def _open_document(self, event):
        item = self.table.identify_row(event.y)
        if not item:
            return
        doc = self.table.item(item, "values")[3]
        if doc and os.path.exists(doc):
            try:
                os.startfile(doc)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open:\n{e}")
        else:
            messagebox.showerror("No File", "Document not found.")

    def reset_filters(self):
        for key, var in self.filter_vars.items():
            if key == "y_axis":
                var.set("Total")
            else:
                var.set("All")
        self.update_analysis()

    def update_analysis(self):
        load_emission_records()
        df = pd.DataFrame(emission_records, columns=[
            "Email", "Entry Date", "Month", "Year", "Company Unit", "Emission Category",
            "Emission Name", "Emission Type", "Factor", "Value", "Total", "Remarks", "Document", "RecordID"
        ])
        df["Year"]  = pd.to_numeric(df["Year"], errors="coerce").fillna(0).astype(int)
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)

        mo_order = ["January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"]

        # filter
        fv, df_f = self.filter_vars, df.copy()
        if fv["unit"].get() != "All":
            df_f = df_f[df_f["Company Unit"] == fv["unit"].get()]
        if fv["year"].get() != "All":
            df_f = df_f[df_f["Year"] == int(fv["year"].get())]
        if fv["month"].get() != "All":
            df_f = df_f[df_f["Month"] == fv["month"].get()]
        if fv["emission_category"].get() != "All":
            df_f = df_f[df_f["Emission Category"] == fv["emission_category"].get()]
        if fv["emission_name"].get() != "All":
            df_f = df_f[df_f["Emission Name"] == fv["emission_name"].get()]
        if fv["emission_type"].get() != "All":
            df_f = df_f[df_f["Emission Type"] == fv["emission_type"].get()]

        y_field = fv["y_axis"].get()

        # determine unit suffix
        emission_name_units = {"Fuel": "L", "Refrigerants": "kg", "Electricity": "kWh"}
        if y_field == "Total":
            unit_suffix = "kg CO2e"
        else:
            en = fv["emission_name"].get()
            unit_suffix = emission_name_units.get(en, "") if en != "All" else ""

        # KPIs
        total = df_f[y_field].sum()
        s1    = df_f[df_f["Emission Category"]=="Scope1"][y_field].sum()
        s2    = df_f[df_f["Emission Category"]=="Scope2"][y_field].sum()
        unit_sum = df_f.groupby("Company Unit")[y_field].sum()
        top_unit= unit_sum.idxmax() if not unit_sum.empty else "N/A"
        type_sum= df_f.groupby("Emission Type")[y_field].sum()
        top_type= type_sum.idxmax() if not type_sum.empty else "N/A"
        cat_sum = df_f.groupby("Emission Category")[y_field].sum()
        top_cat = cat_sum.idxmax() if not cat_sum.empty else "N/A"

        self.kpi_total.lbl.config(text=f"{total:.2f} {unit_suffix}")
        self.kpi_scope1.lbl.config(text=f"{s1:.2f} {unit_suffix}")
        self.kpi_scope2.lbl.config(text=f"{s2:.2f} {unit_suffix}")
        self.kpi_top_unit.lbl.config(text=top_unit)
        self.kpi_top_type.lbl.config(text=top_type)
        self.kpi_top_cat.lbl.config(text=top_cat)

        # Summary
        if not df_f.empty and total>0:
            monthly_sum = df_f.groupby("Month")[y_field].sum().reindex(mo_order, fill_value=0)
            high_m, high_v = monthly_sum.idxmax(), monthly_sum.max()
            low_m, low_v   = monthly_sum.idxmin(), monthly_sum.min()
            low_unit = unit_sum.idxmin() if not unit_sum.empty else "N/A"
            low_u_val= unit_sum.min() if not unit_sum.empty else 0
            scope1 = cat_sum.get("Scope1",0)
            scope2 = cat_sum.get("Scope2",0)
            tot_scope = scope1+scope2 or 1
            p1 = scope1/tot_scope*100; p2=scope2/tot_scope*100
            emp_cnt = df_f.groupby("Email").size()
            top_emp = emp_cnt.idxmax() if not emp_cnt.empty else "N/A"
            top_emp_ct = emp_cnt.max() if not emp_cnt.empty else 0
            val_sum = df_f.groupby("Emission Type")["Value"].sum()
            intensity = {et: (type_sum[et]/val_sum.get(et,1)) if val_sum.get(et,0)>0 else 0
                         for et in type_sum.index}
            lines = [
                f"Monthly Trends: ↑ {high_m} ({high_v:.2f} {unit_suffix}), ↓ {low_m} ({low_v:.2f} {unit_suffix}).",
                f"Company Units: ↑ {top_unit} ({unit_sum.max():.2f} {unit_suffix}), ↓ {low_unit} ({low_u_val:.2f} {unit_suffix}).",
                f"Scopes: Scope1 {scope1:.2f} {unit_suffix} ({p1:.1f}%), Scope2 {scope2:.2f} {unit_suffix} ({p2:.1f}%).",
                f"Top User: {top_emp} ({top_emp_ct} entries).",
                f"Top Emission Type: {top_type} ({type_sum.max():.2f} {unit_suffix}).",
                "Intensity: " + "; ".join(f"{et}: {ratio:.2f}/{unit_suffix}" for et, ratio in intensity.items())
            ]
            summary_text = "\n\n".join("• "+ln for ln in lines)
        else:
            summary_text = "No data to analyze."
        self.summary_label.config(text=summary_text)

        # Chart 1: Monthly by Type
        pivot1 = df_f.pivot_table(index="Month", columns="Emission Type", values=y_field,
                                  aggfunc="sum", fill_value=0).reindex(mo_order, fill_value=0)
        ax1 = self.ax1; ax1.clear()
        idxs = np.arange(len(mo_order))
        cols = pivot1.columns.tolist()
        bw = 0.8/len(cols) if cols else 0.8
        for i, col in enumerate(cols):
            ax1.bar(idxs + i*bw, pivot1[col], bw, label=col)
        ax1.set_xticks(idxs + bw*(len(cols)-1)/2 if cols else idxs)
        ax1.set_xticklabels(mo_order, rotation=45, ha="right")
        ax1.set_title("Monthly Emissions by Type")
        ax1.set_ylabel(f"{y_field} ({unit_suffix})" if unit_suffix else y_field)
        ax1.grid(axis="y", linestyle="--")
        h1,l1 = ax1.get_legend_handles_labels()
        if h1: ax1.legend(bbox_to_anchor=(1.02,1), loc="upper left")
        ax1.ticklabel_format(style="plain", axis="y")
        self.fig1.tight_layout(); self.canvas1.draw()

        # Chart 2: Distribution by Unit
        ax2 = self.ax2; ax2.clear()
        grp2 = df_f.groupby("Company Unit")[y_field].sum()
        if not grp2.empty:
            vals, labs = grp2.values, grp2.index
            ax2.pie(vals, labels=labs, autopct=lambda p: f"{p:.1f}%", startangle=90, wedgeprops=dict(width=0.4))
            ax2.set_title("Emission Distribution by Company Unit")
        else:
            ax2.text(0.5,0.5,"No Data",ha="center")
        self.fig2.tight_layout(); self.canvas2.draw()

        # Chart 3: Yearly by Category
        pivot3 = df_f.pivot_table(index="Year", columns="Emission Category", values=y_field,
                                  aggfunc="sum", fill_value=0)
        ax3 = self.ax3; ax3.clear()
        yrs3 = sorted(pivot3.index)
        idxs3 = np.arange(len(yrs3))
        cols3 = pivot3.columns.tolist()
        bw3 = 0.8/len(cols3) if cols3 else 0.8
        for i,col in enumerate(cols3):
            ax3.bar(idxs3 + i*bw3, pivot3[col], bw3, label=col)
        ax3.set_xticks(idxs3 + bw3*(len(cols3)-1)/2 if cols3 else idxs3)
        ax3.set_xticklabels(yrs3, rotation=45)
        ax3.set_title("Yearly Emissions by Category")
        ax3.set_ylabel(f"{y_field} ({unit_suffix})" if unit_suffix else y_field)
        ax3.grid(axis="y", linestyle="--")
        h3,l3 = ax3.get_legend_handles_labels()
        if h3: ax3.legend(bbox_to_anchor=(1.02,1), loc="upper left")
        ax3.ticklabel_format(style="plain", axis="y")
        self.fig3.tight_layout(); self.canvas3.draw()

        # Chart 4: Yearly by Type
        pivot4 = df_f.pivot_table(index="Year", columns="Emission Type", values=y_field,
                                  aggfunc="sum", fill_value=0)
        ax4 = self.ax4; ax4.clear()
        yrs4 = sorted(pivot4.index)
        idxs4 = np.arange(len(yrs4))
        cols4 = pivot4.columns.tolist()
        bw4 = 0.8/len(cols4) if cols4 else 0.8
        for i,col in enumerate(cols4):
            ax4.bar(idxs4 + i*bw4, pivot4[col], bw4, label=col)
        ax4.set_xticks(idxs4 + bw4*(len(cols4)-1)/2 if cols4 else idxs4)
        ax4.set_xticklabels(yrs4, rotation=45)
        ax4.set_title("Yearly Emissions by Type")
        ax4.set_ylabel(f"{y_field} ({unit_suffix})" if unit_suffix else y_field)
        ax4.grid(axis="y", linestyle="--")
        h4,l4 = ax4.get_legend_handles_labels()
        if h4: ax4.legend(bbox_to_anchor=(1.02,1), loc="upper left")
        ax4.ticklabel_format(style="plain", axis="y")
        self.fig4.tight_layout(); self.canvas4.draw()

        # Chart 5: Yearly by Unit
        pivot5 = df_f.pivot_table(index="Year", columns="Company Unit", values=y_field,
                                  aggfunc="sum", fill_value=0)
        ax5 = self.ax5; ax5.clear()
        yrs5 = sorted(pivot5.index)
        idxs5 = np.arange(len(yrs5))
        cols5 = pivot5.columns.tolist()
        bw5 = 0.8/len(cols5) if cols5 else 0.8
        for i,col in enumerate(cols5):
            ax5.bar(idxs5 + i*bw5, pivot5[col], bw5, label=col)
        ax5.set_xticks(idxs5 + bw5*(len(cols5)-1)/2 if cols5 else idxs5)
        ax5.set_xticklabels(yrs5, rotation=45)
        ax5.set_title("Yearly Emissions by Company Unit")
        ax5.set_ylabel(f"{y_field} ({unit_suffix})" if unit_suffix else y_field)
        ax5.grid(axis="y", linestyle="--")
        h5,l5 = ax5.get_legend_handles_labels()
        if h5: ax5.legend(bbox_to_anchor=(1.02,1), loc="upper left")
        ax5.ticklabel_format(style="plain", axis="y")
        self.fig5.tight_layout(); self.canvas5.draw()

        # Update Table
        self.table.delete(*self.table.get_children())
        for _, row in df_f.iterrows():
            self.table.insert("", "end", values=(
                row["Emission Name"],
                row["Emission Type"],
                f"{row[y_field]:.2f} {unit_suffix}",
                row["Document"]
            ))

    def tkraise(self, above=None):
        self.update_analysis()
        super().tkraise(above)

# ---------------- EmissionDataPage ----------------
class EmissionDataPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        self.sort_ascending = True
        self.main_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        self.main_frame.pack(fill="both", expand=True)

        header_label = tk.Label(self.main_frame, text="RMX Joss Carbon Emission Tracking System",
                                bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,20,"bold"))
        header_label.pack(pady=10)

        header_card = create_card(self.main_frame)
        tk.Label(header_card, text="Emission Data Records", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,16,"bold")).pack(pady=10)
        self.user_label = tk.Label(header_card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12))
        self.user_label.pack(pady=5)

        filter_frame = tk.Frame(self.main_frame, bg=BACKGROUND_COLOR)
        filter_frame.pack(pady=10)
        filters = [
            ("Company Unit:", "unit", ["All"]+system_config["units"]),
            ("Month:", "month", ["All","January","February","March","April","May","June",
                                  "July","August","September","October","November","December"]),
            ("Year:", "year", ["All"]+[str(y) for y in range(2020,2031)]),
            ("Category:", "emission_category", ["All","Scope1","Scope2","Scope3"]),
            ("Name:", "emission_name", ["All","Fuel","Refrigerants","Electricity"]),
            ("Type:", "emission_type", ["All","Diesel","Petrol","PNG","LPG","R-22","R-410A","Electricity"])
        ]
        self.filter_vars = {}
        col = 0
        for lbl, key, opts in filters:
            tk.Label(filter_frame, text=lbl, bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=0, column=col, padx=3, pady=3, sticky="e")
            var = tk.StringVar(value="All")
            self.filter_vars[key] = var
            cb = ttk.Combobox(filter_frame, textvariable=var, values=opts, state="readonly", width=10)
            cb.grid(row=0, column=col+1, padx=3, pady=3)
            col += 2

        self.btn_edit = tk.Button(filter_frame, text="Edit Selected", command=self.edit_record,
                                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"), padx=10, pady=5)
        self.btn_edit.grid(row=0, column=col, padx=5)
        add_hover(self.btn_edit, PRIMARY_COLOR, PRIMARY_HOVER)
        self.btn_delete = tk.Button(filter_frame, text="Delete Selected", command=self.delete_record,
                                    bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"), padx=10, pady=5)
        self.btn_delete.grid(row=0, column=col+1, padx=5)
        add_hover(self.btn_delete, DANGER_COLOR, DANGER_HOVER)
        tk.Button(filter_frame, text="Apply Filters", command=self.apply_filters,
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"), padx=10, pady=5).grid(row=0, column=col+2, padx=10)

        table_card = create_card(self.main_frame)
        cols = ("Gmail", "Entry Date", "Month", "Year", "Company Unit", "Emission Category",
                "Emission Name", "Emission Type", "Factor", "Value", "Total", "Remarks", "Document", "RecordID")
        self.tree = ttk.Treeview(table_card, columns=cols, show="headings", height=20)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="center", width=100)
        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        table_card.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_treeview_double_click)

        btn_frame = tk.Frame(self.main_frame, bg=BACKGROUND_COLOR)
        btn_frame.pack(pady=20)
        for txt, cmd in [("Refresh", self.refresh_table),
                         ("Export to Excel", self.export_to_excel),
                         ("Go to Data Entry", lambda: controller.show_frame("DataEntryPage")),
                         ("Go to Analysis", lambda: controller.show_frame("AnalysisPage")),
                         ("Back to Home", lambda: controller.show_frame("HomePage"))]:
            bgc = PRIMARY_COLOR if txt != "Back to Home" else DANGER_COLOR
            btn = tk.Button(btn_frame, text=txt, command=cmd, bg=bgc, fg="white",
                            font=(FONT_FAMILY,12,"bold"), padx=15, pady=8)
            btn.pack(side="left", padx=5)
            add_hover(btn, bgc, PRIMARY_HOVER if bgc==PRIMARY_COLOR else DANGER_HOVER)

        self.refresh_table()

    def update_edit_delete_buttons(self):
        self.btn_edit.config(state="normal")
        role = get_user_role(self.controller.email)
        if role in ["Admin", "Manager"]:
            self.btn_delete.config(state="normal")
        else:
            self.btn_delete.config(state="disabled")

    def refresh_table(self, records=None):
        load_emission_records()
        if records is None:
            records = emission_records
        for item in self.tree.get_children():
            self.tree.delete(item)
        for rec in records:
            self.tree.insert("", "end", iid=str(rec[13]), values=list(rec[:14]))  # Include RecordID
        logging.info("Table refreshed.")

    def export_to_excel(self):
        fp = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                          filetypes=[("Excel files","*.xlsx"),("All files","*.*")],
                                          title="Save as")
        if not fp:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Emission Data"
            headers = ("Gmail","Entry Date","Month","Year","Company Unit","Emission Category",
                       "Emission Name","Emission Type","Factor","Value","Total","Remarks","Document")
            ws.append(headers)
            for rec in emission_records:
                ws.append(list(rec[:13]))
            wb.save(fp)
            messagebox.showinfo("Exported", f"Saved to:\n{fp}")
            logging.info("Exported to Excel.")
        except Exception as e:
            logging.error("Export error: " + str(e))
            messagebox.showerror("Error", str(e))

    def apply_filters(self):
        flts = self.filter_vars
        filtered = []
        for rec in emission_records:
            if ((flts["unit"].get()=="All" or rec[4]==flts["unit"].get()) and
                (flts["month"].get()=="All" or rec[2]==flts["month"].get()) and
                (flts["year"].get()=="All" or rec[3]==flts["year"].get()) and
                (flts["emission_category"].get()=="All" or rec[5]==flts["emission_category"].get()) and
                (flts["emission_name"].get()=="All" or rec[6]==flts["emission_name"].get()) and
                (flts["emission_type"].get()=="All" or rec[7]==flts["emission_type"].get())):
                filtered.append(rec)
        self.refresh_table(filtered)

    def on_treeview_double_click(self, event):
        col = int(self.tree.identify_column(event.x).replace("#",""))-1
        if col!=12:
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        doc = self.tree.item(item,"values")[12]
        if doc and os.path.exists(doc):
            try:
                os.startfile(doc)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open:\n{e}")
        else:
            messagebox.showerror("No File","Document not found.")

    def edit_record(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a record to edit.")
            return
            
        # Get the record ID from the selected item
        record_id = sel[0]
        
        # Find the record in emission_records
        record = None
        idx = None
        for i, r in enumerate(emission_records):
            if str(r[13]) == record_id:
                record = r
                idx = i
                break
                
        if not record:
            messagebox.showerror("Error", "Record not found.")
            return
            
        # Open edit dialog
        EditDialog(self, record, idx)

    def delete_record(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select a record to delete.")
            return
            
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?"):
            return
            
        # Get the record ID from the selected item
        record_id = sel[0]
        
        # Delete from database
        cfg = system_config["database"]["mssql"]
        conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
        if not conn:
            return
            
        try:
            cur = conn.cursor()
            sql = "DELETE FROM EmissionRecords WHERE record_id = ?"
            cur.execute(sql, (record_id,))
            conn.commit()
            
            # Remove from local list
            emission_records[:] = [r for r in emission_records if str(r[13]) != record_id]
            
            messagebox.showinfo("Success", "Record deleted successfully.")
            self.refresh_table()
            
        except Exception as e:
            logging.error(f"Error deleting record: {str(e)}")
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")
        finally:
            conn.close()

# ---------------- EditDialog ----------------
class EditDialog(tk.Toplevel):
    def __init__(self, parent_page, record, rec_index):
        super().__init__(parent_page)
        self.title("Edit Record")
        self.parent_page = parent_page
        self.rec_index = rec_index
        self.record = record
        
        # Define the fields and their corresponding indices in the record tuple
        fields = [
            ("Company Unit:", 4),
            ("Month:", 2),
            ("Year:", 3),
            ("Emission Category:", 5),
            ("Emission Name:", 6),
            ("Emission Type:", 7),
            ("Factor:", 8),
            ("Value:", 9),
            ("Remarks:", 11),
            ("Document:", 12)
        ]
        
        self.vars = []
        for i, (label, idx) in enumerate(fields):
            tk.Label(self, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            var = tk.StringVar(value=record[idx])
            ent = tk.Entry(self, textvariable=var, width=40 if i>=8 else 20)
            if i==6:  # Factor field
                ent.config(state="readonly")
            ent.grid(row=i, column=1, padx=5, pady=5)
            self.vars.append(var)
            
        tk.Button(self, text="Save Changes", command=self.save_changes,
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).grid(row=len(fields), column=0, columnspan=2, pady=10)

    def save_changes(self):
        try:
            # Get values from entry fields
            factor = float(self.vars[6].get())
            value = self.vars[7].get()
            total = update_total_value(factor, value)
            
            # Create updated record
            updated = (
                self.record[0],  # email
                self.record[1],  # entry_date
                self.vars[1].get(),  # month
                self.vars[2].get(),  # year
                self.vars[0].get(),  # company_unit
                self.vars[3].get(),  # emission_category
                self.vars[4].get(),  # emission_name
                self.vars[5].get(),  # emission_type
                str(factor),  # factor
                value,  # value
                total,  # total
                self.vars[8].get(),  # remarks
                self.vars[9].get(),  # document
                self.record[13]  # record_id
            )
            
            # Update database
            cfg = system_config["database"]["mssql"]
            conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
            if not conn:
                return
                
            try:
                cur = conn.cursor()
                update_sql = """
                UPDATE EmissionRecords
                SET month = ?, year = ?, company_unit = ?, emission_category = ?,
                    emission_name = ?, emission_type = ?, factor = ?, value = ?,
                    total = ?, remarks = ?, document = ?
                WHERE record_id = ?
                """
                params = (
                    updated[2], updated[3], updated[4], updated[5],
                    updated[6], updated[7], updated[8], updated[9],
                    updated[10], updated[11], updated[12],
                    updated[13]
                )
                cur.execute(update_sql, params)
                conn.commit()
                
                # Update local list
                emission_records[self.rec_index] = updated
                
                messagebox.showinfo("Success", "Record updated successfully.")
                self.parent_page.refresh_table()
                self.destroy()
                
            except Exception as e:
                logging.error(f"Error updating record: {str(e)}")
                messagebox.showerror("Error", f"Failed to update record: {str(e)}")
            finally:
                conn.close()
                
        except Exception as e:
            logging.error(f"Error saving changes: {str(e)}")
            messagebox.showerror("Error", f"Failed to save changes: {str(e)}")

# ---------------- DataEntryPage ----------------
class DataEntryPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        self.fuel_file_vars = {}
        self.refrig_file_vars= {}
        self.elec_file_var  = tk.StringVar()

        self.main_frame = ScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True)
        cf = self.main_frame.scrollable_frame

        tk.Label(cf, text="RMX Joss Carbon Emission Tracking System", bg=BACKGROUND_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,20,"bold")).pack(pady=10)

        top_card = create_card(cf)
        tk.Label(top_card, text="Choose Company Unit:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=0,padx=10,pady=10,sticky="w")
        self.unit_var = tk.StringVar()
        unit_dd = ttk.Combobox(top_card, textvariable=self.unit_var, state="readonly", width=12,
                               values=system_config["units"])
        unit_dd.grid(row=0,column=1,padx=10,pady=10)
        unit_dd.current(0)
        self.unit_var.trace('w', lambda *a: self.reset_input_fields())

        tk.Label(top_card, text="Month:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=2,padx=10,pady=10,sticky="w")
        self.month_var = tk.StringVar()
        mon_dd = ttk.Combobox(top_card, textvariable=self.month_var, state="readonly", width=12,
                              values=["January","February","March","April","May","June","July","August","September","October","November","December"])
        mon_dd.grid(row=0,column=3,padx=10,pady=10)
        mon_dd.current(0)

        tk.Label(top_card, text="Year:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=4,padx=10,pady=10,sticky="w")
        cur_year = datetime.now().year
        self.year_var = tk.StringVar(value=str(cur_year))
        yr_dd = ttk.Combobox(top_card, textvariable=self.year_var, state="readonly", width=10,
                             values=[str(y) for y in range(2020,2032)])
        yr_dd.grid(row=0,column=5,padx=10,pady=10)

        tk.Label(top_card, text="Current Date:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=6,padx=10,pady=10,sticky="w")
        self.current_date_label = tk.Label(top_card, text=datetime.now().strftime("%Y-%m-%d"),
                                           width=12, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10))
        self.current_date_label.grid(row=0,column=7,padx=10,pady=10)

        self.fuel_types   = [{"name":k,"unit":"Liters","factor":v} for k,v in system_config["scope_factors"]["Fuel"].items()]
        self.refrig_types = [{"name":k,"unit":"kg","factor":v} for k,v in system_config["scope_factors"]["Refrigerants"].items()]
        self.elec_factor  = system_config["scope_factors"]["Electricity"]["Electricity"]

        # Scope 1: Fuel & Refrigerants
        scope1_card = create_card(cf, fill="both")
        tk.Label(scope1_card, text="Scope 1: Fuel & Refrigerant Entries", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,14,"bold")).pack(pady=10)
        container = tk.Frame(scope1_card, bg=CARD_COLOR)
        container.pack(padx=10,pady=10,fill="both")

        # Fuel Frame
        fuel_frame = tk.LabelFrame(container, text="Fuel Data", bg=CARD_COLOR, fg=TEXT_COLOR,
                                   font=(FONT_FAMILY,12,"bold"), padx=10, pady=10)
        fuel_frame.pack(side="left", fill="both", expand=True, padx=10,pady=10)
        headers = ["Type","Unit","Factor","Value","Total","Remarks","Upload Document"]
        for c,h in enumerate(headers):
            tk.Label(fuel_frame, text=h, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10,"bold")).grid(row=0,column=c,padx=8,pady=8)
        self.fuel_amount_vars  = {}
        self.fuel_total_labels = {}
        self.fuel_remarks_vars = {}
        for i,ft in enumerate(self.fuel_types, start=1):
            tk.Label(fuel_frame, text=ft["name"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=i,column=0,padx=8,pady=8)
            tk.Label(fuel_frame, text=ft["unit"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=i,column=1,padx=8,pady=8)
            fe = tk.Entry(fuel_frame, width=10, font=(FONT_FAMILY,10))
            fe.insert(0,str(ft["factor"]))
            fe.config(state="readonly", readonlybackground=CARD_COLOR, fg=TEXT_COLOR)
            fe.grid(row=i,column=2,padx=8,pady=8)
            val_var = tk.StringVar()
            self.fuel_amount_vars[ft["name"]] = val_var
            ne = NumericEntry(fuel_frame, textvariable=val_var, width=10, font=(FONT_FAMILY,10))
            ne.grid(row=i,column=3,padx=8,pady=8)
            add_focus_effect(ne)
            tl = tk.Label(fuel_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10))
            tl.grid(row=i,column=4,padx=8,pady=8)
            self.fuel_total_labels[ft["name"]] = tl
            rem_var = tk.StringVar()
            self.fuel_remarks_vars[ft["name"]] = rem_var
            tk.Entry(fuel_frame, textvariable=rem_var, width=20).grid(row=i,column=5,padx=8,pady=8)
            fv = tk.StringVar()
            self.fuel_file_vars[ft["name"]] = fv
            btn = tk.Button(fuel_frame, text="Upload",
                            command=lambda var=fv, f=ft: upload_document(var,
                                                                         self.unit_var.get(),
                                                                         self.current_date_label.cget("text"),
                                                                         "Fuel", f["name"], self.controller.email),
                            bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10), relief="raised", bd=2, padx=10, pady=4)
            btn.grid(row=i,column=6,padx=8,pady=8)
            add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)
            def cb_fuel(*a, name=ft["name"], fac=ft["factor"]):
                self.fuel_total_labels[name].config(text=update_total_value(fac, self.fuel_amount_vars[name].get()))
            val_var.trace("w", cb_fuel)

        # Refrigerants Frame
        refrig_frame = tk.LabelFrame(container, text="Refrigerants", bg=CARD_COLOR, fg=TEXT_COLOR,
                                     font=(FONT_FAMILY,12,"bold"), padx=10, pady=10)
        refrig_frame.pack(side="right", fill="both", expand=True, padx=10,pady=10)
        for c,h in enumerate(headers):
            tk.Label(refrig_frame, text=h, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10,"bold")).grid(row=0,column=c,padx=8,pady=8)
        self.refrig_amount_vars  = {}
        self.refrig_total_labels = {}
        self.refrig_remarks_vars = {}
        for i,rt in enumerate(self.refrig_types, start=1):
            tk.Label(refrig_frame, text=rt["name"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=i,column=0,padx=8,pady=8)
            tk.Label(refrig_frame, text=rt["unit"], bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=i,column=1,padx=8,pady=8)
            rvf = tk.StringVar(value=str(rt["factor"]))
            fe = tk.Entry(refrig_frame, textvariable=rvf, width=10, font=(FONT_FAMILY,10))
            fe.config(state="readonly", readonlybackground=CARD_COLOR, fg=TEXT_COLOR)
            fe.grid(row=i,column=2,padx=8,pady=8)
            val_var = tk.StringVar()
            self.refrig_amount_vars[rt["name"]] = val_var
            ne = NumericEntry(refrig_frame, textvariable=val_var, width=10, font=(FONT_FAMILY,10))
            ne.grid(row=i,column=3,padx=8,pady=8)
            add_focus_effect(ne)
            tl = tk.Label(refrig_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10))
            tl.grid(row=i,column=4,padx=8,pady=8)
            self.refrig_total_labels[rt["name"]] = tl
            rem_var = tk.StringVar()
            self.refrig_remarks_vars[rt["name"]] = rem_var
            tk.Entry(refrig_frame, textvariable=rem_var, width=20).grid(row=i,column=5,padx=8,pady=8)
            fv = tk.StringVar()
            self.refrig_file_vars[rt["name"]] = fv
            btn = tk.Button(refrig_frame, text="Upload",
                            command=lambda var=fv, r=rt: upload_document(var,
                                                                         self.unit_var.get(),
                                                                         self.current_date_label.cget("text"),
                                                                         "Refrigerants", r["name"], self.controller.email),
                            bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10), relief="raised", bd=2, padx=10, pady=4)
            btn.grid(row=i,column=6,padx=8,pady=8)
            add_hover(btn, PRIMARY_COLOR, PRIMARY_HOVER)
            def cb_refrig(*a, name=rt["name"], fac_var=rvf):
                fac = float(fac_var.get()) if fac_var.get() else 0
                self.refrig_total_labels[name].config(text=update_total_value(fac, self.refrig_amount_vars[name].get()))
            val_var.trace("w", cb_refrig)

        # Scope 2: Electricity
        scope2_card = create_card(cf)
        tk.Label(scope2_card, text="Scope 2: Electricity Entries", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,14,"bold")).pack(pady=10)
        elec_frame = tk.LabelFrame(scope2_card, text="Electricity Data", bg=CARD_COLOR, fg=TEXT_COLOR,
                                   font=(FONT_FAMILY,12,"bold"), padx=10, pady=10)
        elec_frame.pack(fill="x", padx=10, pady=10)
        eheaders = ["Type","Unit","Factor","Value","Total","Remarks","Upload Document"]
        for c,h in enumerate(eheaders):
            tk.Label(elec_frame, text=h, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10,"bold")).grid(row=0,column=c,padx=8,pady=8)
        tk.Label(elec_frame, text="Electricity", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=1,column=0,padx=8,pady=8)
        tk.Label(elec_frame, text="kWh", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).grid(row=1,column=1,padx=8,pady=8)
        fe2 = tk.Entry(elec_frame, width=10, font=(FONT_FAMILY,10))
        fe2.insert(0,str(self.elec_factor))
        fe2.config(state="readonly", readonlybackground=CARD_COLOR, fg=TEXT_COLOR)
        fe2.grid(row=1,column=2,padx=8,pady=8)
        self.elec_amount_var = tk.StringVar()
        ne2 = NumericEntry(elec_frame, textvariable=self.elec_amount_var, width=10, font=(FONT_FAMILY,10))
        ne2.grid(row=1,column=3,padx=8,pady=8)
        add_focus_effect(ne2)
        elec_total_label = tk.Label(elec_frame, text="0.00", width=10, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10))
        elec_total_label.grid(row=1,column=4,padx=8,pady=8)
        def cb_elec(*a):
            elec_total_label.config(text=update_total_value(self.elec_factor, self.elec_amount_var.get()))
        self.elec_amount_var.trace("w", cb_elec)
        self.elec_remarks_var = tk.StringVar()
        tk.Entry(elec_frame, textvariable=self.elec_remarks_var, width=20).grid(row=1,column=5,padx=8,pady=8)
        btn_elec = tk.Button(elec_frame, text="Upload",
                             command=lambda var=self.elec_file_var: upload_document(var,
                                                                                   self.unit_var.get(),
                                                                                   self.current_date_label.cget("text"),
                                                                                   "Electricity","Electricity",self.controller.email),
                             bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10),
                             relief="raised", bd=2, padx=10, pady=4)
        btn_elec.grid(row=1,column=6,padx=8,pady=8)
        add_hover(btn_elec, PRIMARY_COLOR, PRIMARY_HOVER)

        # Scope 3 placeholder
        scope3_card = create_card(cf)
        tk.Label(scope3_card, text="Scope 3: Reserved for Future Edits", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,14,"bold")).pack(pady=10)
        tk.Label(scope3_card, text="Reserved for future enhancements", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,12)).pack(pady=10)

        btn_frame = tk.Frame(cf, bg=BACKGROUND_COLOR)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="Submit", command=self.submit_data_handler, bg=PRIMARY_COLOR, fg="white",
                  font=(FONT_FAMILY,12,"bold"), relief="raised", bd=2, padx=20, pady=10).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Go to Emission Data", command=lambda: self.controller.show_frame("EmissionDataPage"),
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), relief="raised", bd=2, padx=20, pady=10).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Back to Home", command=lambda: self.controller.show_frame("HomePage"),
                  bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), relief="raised", bd=2, padx=20, pady=10).pack(side="left", padx=10)

    def reset_input_fields(self):
        for k in self.fuel_file_vars:   self.fuel_file_vars[k].set("")
        for k in self.refrig_file_vars: self.refrig_file_vars[k].set("")
        self.elec_file_var.set("")
        for k in self.fuel_amount_vars: self.fuel_amount_vars[k].set("")
        for k in self.fuel_total_labels: self.fuel_total_labels[k].config(text="0.00")
        for k in self.fuel_remarks_vars: self.fuel_remarks_vars[k].set("")
        for k in self.refrig_amount_vars: self.refrig_amount_vars[k].set("")
        for k in self.refrig_total_labels: self.refrig_total_labels[k].config(text="0.00")
        for k in self.refrig_remarks_vars: self.refrig_remarks_vars[k].set("")
        self.elec_amount_var.set("")
        self.elec_remarks_var.set("")

    def submit_data_handler(self):
        try:
            unit = self.unit_var.get().strip()
            month= self.month_var.get().strip()
            year = self.year_var.get().strip()
            entry_date = self.current_date_label.cget("text")
            user_email = self.controller.email
            if not (unit and month and year and entry_date):
                messagebox.showerror("Mandatory Fields Missing","Please fill out all common fields.")
                return
            for name,var in self.fuel_amount_vars.items():
                if var.get().strip() and not self.fuel_file_vars[name].get():
                    messagebox.showerror("Document Missing",f"Please upload document for fuel '{name}'.")
                    return
            for name,var in self.refrig_amount_vars.items():
                if var.get().strip() and not self.refrig_file_vars[name].get():
                    messagebox.showerror("Document Missing",f"Please upload document for refrigerant '{name}'.")
                    return
            if self.elec_amount_var.get().strip() and not self.elec_file_var.get():
                messagebox.showerror("Document Missing","Please upload document for Electricity.")
                return

            global record_id_counter
            new_records = []

            for ft in self.fuel_types:
                val = self.fuel_amount_vars[ft["name"]].get().strip()
                if val:
                    tot = self.fuel_total_labels[ft["name"]].cget("text")
                    rem = self.fuel_remarks_vars[ft["name"]].get().strip()
                    fp  = self.fuel_file_vars[ft["name"]].get()
                    rec = (
                        user_email, entry_date, month, year, unit,
                        "Scope1","Fuel",ft["name"],
                        str(ft["factor"]), val, tot,
                        rem, fp, record_id_counter
                    )
                    record_id_counter += 1
                    new_records.append(rec)

            for rt in self.refrig_types:
                val = self.refrig_amount_vars[rt["name"]].get().strip()
                if val:
                    tot = self.refrig_total_labels[rt["name"]].cget("text")
                    rem = self.refrig_remarks_vars[rt["name"]].get().strip()
                    fp  = self.refrig_file_vars[rt["name"]].get()
                    rec = (
                        user_email, entry_date, month, year, unit,
                        "Scope1","Refrigerants",rt["name"],
                        str(rt["factor"]), val, tot,
                        rem, fp, record_id_counter
                    )
                    record_id_counter += 1
                    new_records.append(rec)

            val = self.elec_amount_var.get().strip()
            if val:
                tot = update_total_value(self.elec_factor, val)
                rem = self.elec_remarks_var.get().strip()
                fp  = self.elec_file_var.get()
                rec = (
                    user_email, entry_date, month, year, unit,
                    "Scope2","Electricity","Electricity",
                    str(self.elec_factor), val, tot,
                    rem, fp, record_id_counter
                )
                record_id_counter += 1
                new_records.append(rec)

            if not new_records:
                messagebox.showwarning("No Data","Enter some values before submitting.")
                return

            emission_records.extend(new_records)
            save_emission_records()
            load_emission_records()
            messagebox.showinfo("Data Submitted","Data submitted successfully!")
            self.reset_input_fields()
            if "EmissionDataPage" in self.controller.frames:
                self.controller.frames["EmissionDataPage"].refresh_table()

        except Exception as e:
            logging.error("Submission Error: " + str(e))
            messagebox.showerror("Submission Error", f"An error occurred: {e}")

# ---------------- LoginPage ----------------
class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        frame = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=300, height=250)
        tk.Label(frame, text="Login to RMX Joss System", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,14,"bold")).pack(pady=10)
        tk.Label(frame, text="Email:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).pack(pady=5)
        self.email_entry = tk.Entry(frame, width=30)
        self.email_entry.pack(pady=5)
        tk.Label(frame, text="Password:", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10)).pack(pady=5)
        self.password_entry = tk.Entry(frame, show="*", width=30)
        self.password_entry.pack(pady=5)
        btn_login = tk.Button(frame, text="Login", command=self.login, bg=PRIMARY_COLOR, fg="white",
                              font=(FONT_FAMILY,10,"bold"))
        btn_login.pack(pady=10)
        add_hover(btn_login, PRIMARY_COLOR, PRIMARY_HOVER)

    def login(self):
        email = self.email_entry.get().strip()
        pwd   = self.password_entry.get().strip()
        if email == system_config["users"]["admin"]["email"] and pwd == system_config["users"]["admin"]["password"]:
            logging.info(f"Admin {email} logged in.")
            self.controller.email = email
            self.controller.show_frame("HomePage")
            return
        for role in ("manager","employee"):
            for u in system_config["users"].get(role, []):
                if u["email"]==email and u["password"]==pwd:
                    logging.info(f"User {email} logged in as {role}.")
                    self.controller.email = email
                    self.controller.show_frame("HomePage")
                    return
        messagebox.showerror("Login Failed","Invalid credentials.")

    def reset(self):
        self.email_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)

# ---------------- HomePage ----------------
class HomePage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        card = tk.Frame(self, bg=CARD_COLOR, bd=1, relief="groove")
        card.place(relx=0.5, rely=0.5, anchor="center", width=500, height=400)
        tk.Label(card, text="Welcome to RMX Joss Carbon Tracking System", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,16,"bold")).pack(pady=20)
        self.user_label = tk.Label(card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12))
        self.user_label.pack(pady=10)
        btn_data = tk.Button(card, text="Data Entry", command=lambda: controller.show_frame("DataEntryPage"),
                             bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_data.pack(pady=10)
        add_hover(btn_data, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_emission = tk.Button(card, text="Emission Data", command=lambda: controller.show_frame("EmissionDataPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_emission.pack(pady=10)
        add_hover(btn_emission, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_analysis = tk.Button(card, text="Analysis", command=lambda: controller.show_frame("AnalysisPage"),
                                 bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_analysis.pack(pady=10)
        add_hover(btn_analysis, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_solar = tk.Button(card, text="Solar Entry", command=lambda: controller.show_frame("SolarEntryPage"),
                              bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_solar.pack(pady=10)
        add_hover(btn_solar, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_admin = tk.Button(card, text="Admin Panel", command=lambda: controller.show_frame("AdminPage"),
                              bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_admin.pack(pady=10)
        add_hover(btn_admin, PRIMARY_COLOR, PRIMARY_HOVER)
        btn_logout = tk.Button(card, text="Logout", command=lambda: controller.logout(),
                               bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,12,"bold"), width=20)
        btn_logout.pack(pady=10)
        add_hover(btn_logout, DANGER_COLOR, DANGER_HOVER)

    def tkraise(self, above=None):
        self.user_label.config(text=f"Logged in as: {self.controller.email}")
        super().tkraise(above)

# ---------------- MainApp ----------------
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RMX Joss Carbon Tracking System")
        self.geometry("1400x900")  # Increased window width
        self.email = None

        load_config()

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (LoginPage, HomePage, AdminPage, DataEntryPage, EmissionDataPage, AnalysisPage, SolarEntryPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        init_db()
        load_emission_records()
        self.show_frame("LoginPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        if hasattr(frame, "update_edit_delete_buttons"):
            frame.update_edit_delete_buttons()
        if page_name == "EmissionDataPage":
            frame.refresh_table()
            frame.update_edit_delete_buttons()
        if hasattr(frame, "user_label"):
            frame.user_label.config(text=f"User: {self.email}")
        frame.tkraise()

    def logout(self):
        self.email = None
        self.frames["LoginPage"].reset()
        self.show_frame("LoginPage")

# ---------------- SolarEntryPage ----------------
class SolarEntryPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=BACKGROUND_COLOR)
        self.controller = controller
        self.sort_ascending = True
        self.main_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        self.main_frame.pack(fill="both", expand=True)

        # Import tkcalendar at the top of the file
        from tkcalendar import DateEntry

        header_label = tk.Label(self.main_frame, text="RMX Joss Carbon Emission Tracking System",
                                bg=BACKGROUND_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,20,"bold"))
        header_label.pack(pady=10)

        header_card = create_card(self.main_frame)
        tk.Label(header_card, text="Solar Energy Data Records", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,16,"bold")).pack(pady=10)
        self.user_label = tk.Label(header_card, text="", bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,12))
        self.user_label.pack(pady=5)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Data Entry
        self.entry_tab = tk.Frame(self.notebook, bg=BACKGROUND_COLOR)
        self.notebook.add(self.entry_tab, text="Solar Data Entry")

        # Tab 2: View Records
        self.view_tab = tk.Frame(self.notebook, bg=BACKGROUND_COLOR)
        self.notebook.add(self.view_tab, text="View Solar Records")

        # Build tabs
        self.build_entry_tab()
        self.build_view_tab()

    def build_entry_tab(self):
        # Top card with unit selection and date
        top_card = create_card(self.entry_tab)
        top_card.pack(fill="x", padx=10, pady=5)

        # Company Unit selection
        tk.Label(top_card, text="Choose Company Unit:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=0,padx=10,pady=10,sticky="w")
        self.unit_var = tk.StringVar()
        unit_dd = ttk.Combobox(top_card, textvariable=self.unit_var, state="readonly", width=12,
                               values=system_config["units"])
        unit_dd.grid(row=0,column=1,padx=10,pady=10)
        unit_dd.current(0)

        # Month selection
        tk.Label(top_card, text="Month:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=2,padx=10,pady=10,sticky="w")
        self.month_var = tk.StringVar()
        mon_dd = ttk.Combobox(top_card, textvariable=self.month_var, state="readonly", width=12,
                              values=["January","February","March","April","May","June",
                                     "July","August","September","October","November","December"])
        mon_dd.grid(row=0,column=3,padx=10,pady=10)
        mon_dd.current(0)

        # Year selection
        tk.Label(top_card, text="Year:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=4,padx=10,pady=10,sticky="w")
        cur_year = datetime.now().year
        self.year_var = tk.StringVar(value=str(cur_year))
        yr_dd = ttk.Combobox(top_card, textvariable=self.year_var, state="readonly", width=10,
                             values=[str(y) for y in range(2020,2032)])
        yr_dd.grid(row=0,column=5,padx=10,pady=10)

        # Current date
        tk.Label(top_card, text="Current Date:", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,10,"bold")).grid(row=0,column=6,padx=10,pady=10,sticky="w")
        self.current_date_label = tk.Label(top_card, text=datetime.now().strftime("%Y-%m-%d"),
                                           width=12, bg=CARD_COLOR, fg=TEXT_COLOR, font=(FONT_FAMILY,10))
        self.current_date_label.grid(row=0,column=7,padx=10,pady=10)

        # Data entry card
        entry_card = create_card(self.entry_tab)
        entry_card.pack(fill="both", expand=True, padx=10, pady=5)
        
        tk.Label(entry_card, text="Solar Energy Data Entry", bg=CARD_COLOR, fg=TEXT_COLOR,
                 font=(FONT_FAMILY,14,"bold")).pack(pady=10)

        # Form frame
        form_frame = tk.Frame(entry_card, bg=CARD_COLOR)
        form_frame.pack(padx=10, pady=10, fill="x")

        # Headers
        headers = ["Date", "Inverter 1", "Inverter 2", "Inverter 3", "Inverter 4", 
                  "Old total", "New Solar Inverter", "Total Generated", "Unit", "Remark", "Document"]
        for i, header in enumerate(headers):
            tk.Label(form_frame, text=header, bg=CARD_COLOR, fg=TEXT_COLOR, 
                     font=(FONT_FAMILY,10,"bold")).grid(row=0, column=i, padx=5, pady=5)

        # Entry fields
        self.date_var = tk.StringVar()
        self.inverter_vars = [tk.StringVar() for _ in range(4)]
        self.old_total_var = tk.StringVar()
        self.new_solar_var = tk.StringVar()
        self.total_generated_var = tk.StringVar()
        self.unit_var_solar = tk.StringVar(value="Kwh")
        self.remark_var = tk.StringVar()
        self.document_var = tk.StringVar()
        
        # Calendar widget for date entry
        from tkcalendar import DateEntry
        date_picker = DateEntry(form_frame, width=12, background=PRIMARY_COLOR,
                              foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                              textvariable=self.date_var)
        date_picker.grid(row=1, column=0, padx=5, pady=5)

        # Inverter entries
        for i in range(4):
            entry = NumericEntry(form_frame, textvariable=self.inverter_vars[i], width=12)
            entry.grid(row=1, column=i+1, padx=5, pady=5)
            entry.bind('<KeyRelease>', self.update_totals)
            add_focus_effect(entry)

        # Old total (read-only)
        old_total_entry = ttk.Entry(form_frame, textvariable=self.old_total_var, 
                                   width=12, state='readonly')
        old_total_entry.grid(row=1, column=5, padx=5, pady=5)

        # New solar inverter entry
        new_solar_entry = NumericEntry(form_frame, textvariable=self.new_solar_var, width=12)
        new_solar_entry.grid(row=1, column=6, padx=5, pady=5)
        new_solar_entry.bind('<KeyRelease>', self.update_totals)
        add_focus_effect(new_solar_entry)

        # Total generated (read-only)
        total_generated_entry = ttk.Entry(form_frame, textvariable=self.total_generated_var, 
                                        width=12, state='readonly')
        total_generated_entry.grid(row=1, column=7, padx=5, pady=5)

        # Unit dropdown (Kwh)
        unit_dd = ttk.Combobox(form_frame, textvariable=self.unit_var_solar, values=["Kwh"], 
                              state="readonly", width=10)
        unit_dd.grid(row=1, column=8, padx=5, pady=5)
        unit_dd.current(0)

        # Remark entry
        ttk.Entry(form_frame, textvariable=self.remark_var, width=15).grid(row=1, column=9, padx=5, pady=5)

        # Document upload button
        upload_btn = ttk.Button(form_frame, text="Upload", command=self.upload_document)
        upload_btn.grid(row=1, column=10, padx=5, pady=5)

        # Action buttons frame
        btn_frame = tk.Frame(entry_card, bg=CARD_COLOR)
        btn_frame.pack(pady=20)

        # Submit button
        submit_btn = ttk.Button(btn_frame, text="Submit", command=self.submit_data)
        submit_btn.pack(side="left", padx=10)

        # Clear button
        clear_btn = ttk.Button(btn_frame, text="Clear", command=self.clear_form)
        clear_btn.pack(side="left", padx=10)

        # Home button
        home_btn = ttk.Button(btn_frame, text="Home", command=lambda: self.controller.show_frame("HomePage"))
        home_btn.pack(side="left", padx=10)

    def build_view_tab(self):
        # Filters frame
        filter_frame = tk.Frame(self.view_tab, bg=BACKGROUND_COLOR)
        filter_frame.pack(fill="x", padx=10, pady=5)

        # Filter dropdowns
        filters = [
            ("Company Unit:", "unit", ["All"] + system_config["units"]),
            ("Month:", "month", ["All", "January", "February", "March", "April", "May", "June",
                               "July", "August", "September", "October", "November", "December"]),
            ("Year:", "year", ["All"] + [str(y) for y in range(2020, 2031)])
        ]
        
        self.filter_vars = {}
        col = 0
        for lbl, key, opts in filters:
            tk.Label(filter_frame, text=lbl, bg=BACKGROUND_COLOR, fg=TEXT_COLOR,
                     font=(FONT_FAMILY,10)).grid(row=0, column=col, padx=3, pady=3, sticky="e")
            var = tk.StringVar(value="All")
            self.filter_vars[key] = var
            cb = ttk.Combobox(filter_frame, textvariable=var, values=opts, state="readonly", width=10)
            cb.grid(row=0, column=col+1, padx=3, pady=3)
            col += 2

        # Action buttons
        btn_frame = tk.Frame(filter_frame, bg=BACKGROUND_COLOR)
        btn_frame.grid(row=0, column=col, columnspan=3, padx=5)

        self.btn_edit = tk.Button(btn_frame, text="Edit Selected", command=self.edit_record,
                                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"),
                                  padx=10, pady=5)
        self.btn_edit.pack(side="left", padx=5)
        add_hover(self.btn_edit, PRIMARY_COLOR, PRIMARY_HOVER)

        self.btn_delete = tk.Button(btn_frame, text="Delete Selected", command=self.delete_record,
                                    bg=DANGER_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"),
                                    padx=10, pady=5)
        self.btn_delete.pack(side="left", padx=5)
        add_hover(self.btn_delete, DANGER_COLOR, DANGER_HOVER)

        apply_btn = tk.Button(btn_frame, text="Apply Filters", command=self.apply_filters,
                              bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10,"bold"),
                              padx=10, pady=5)
        apply_btn.pack(side="left", padx=5)
        add_hover(apply_btn, PRIMARY_COLOR, PRIMARY_HOVER)

        # Table frame with proper weights
        table_frame = tk.Frame(self.view_tab, bg=BACKGROUND_COLOR)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Table card
        table_card = create_card(table_frame)
        table_card.pack(fill="both", expand=True)

        cols = ("Gmail", "Entry Date", "Month", "Year", "Company Unit", "Date", "Inverter 1", "Inverter 2",
                "Inverter 3", "Inverter 4", "Old total", "New Solar Inverter", "Total Generated",
                "Unit", "Remark", "Document")

        self.tree = ttk.Treeview(table_card, columns=cols, show="headings", height=20)

        # Configure optimized column widths
        column_widths = {
            "Gmail": 150,
            "Entry Date": 90,
            "Month": 70,
            "Year": 50,
            "Company Unit": 90,
            "Date": 90,
            "Inverter 1": 70,
            "Inverter 2": 70,
            "Inverter 3": 70,
            "Inverter 4": 70,
            "Old total": 70,
            "New Solar Inverter": 110,
            "Total Generated": 100,
            "Unit": 50,
            "Remark": 120,
            "Document": 180
        }

        # Apply column configurations
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="center", width=column_widths[c], minwidth=column_widths[c])

        # Scrollbars
        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout with proper scrollbar positioning
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # Configure grid weights for proper expansion
        table_card.grid_columnconfigure(0, weight=1)
        table_card.grid_rowconfigure(0, weight=1)

        # Bind events
        self.tree.bind("<Double-1>", self.on_treeview_double_click)

        # Bottom button frame with fixed height
        bottom_frame = tk.Frame(self.view_tab, bg=BACKGROUND_COLOR, height=50)
        bottom_frame.pack(fill="x", padx=10, pady=10)
        bottom_frame.pack_propagate(False)  # Prevent frame from shrinking

        # Navigation buttons
        nav_frame = tk.Frame(bottom_frame, bg=BACKGROUND_COLOR)
        nav_frame.pack(expand=True)

        buttons = [
            ("Refresh", self.refresh_table, PRIMARY_COLOR),
            ("Export to Excel", self.export_to_excel, PRIMARY_COLOR),
            ("Back to Home", lambda: self.controller.show_frame("HomePage"), DANGER_COLOR)
        ]

        for text, command, color in buttons:
            btn = tk.Button(nav_frame, text=text, command=command,
                           bg=color, fg="white", font=(FONT_FAMILY,12,"bold"),
                           relief="raised", bd=2, padx=20, pady=10)
            btn.pack(side="left", padx=5)
            add_hover(btn, color, PRIMARY_HOVER if color == PRIMARY_COLOR else DANGER_HOVER)

        # Initial refresh
        self.refresh_table()

    def update_totals(self, *args):
        """Update the Old total and Total Generated fields based on inverter values"""
        try:
            # Calculate old total (sum of inverters)
            inverter_values = []
            for var in self.inverter_vars:
                try:
                    value = float(var.get() or 0)
                    inverter_values.append(value)
                except ValueError:
                    inverter_values.append(0)
            
            old_total = sum(inverter_values)
            self.old_total_var.set(f"{old_total:.2f}")

            # Calculate total generated (old total + new solar inverter)
            try:
                new_solar = float(self.new_solar_var.get() or 0)
            except ValueError:
                new_solar = 0

            total_generated = old_total + new_solar
            self.total_generated_var.set(f"{total_generated:.2f}")

        except Exception as e:
            logging.error(f"Error updating totals: {str(e)}")

    def upload_document(self):
        fp = filedialog.askopenfilename(
            filetypes=[
                ("All files","*.*"),
                ("PDF","*.pdf"),
                ("Excel","*.xlsx;*.xls"),
                ("Images","*.png;*.jpg;*.jpeg")
            ],
            title="Select document"
        )
        if not fp:
            return
        role = "Admin" if self.controller.email == system_config["users"]["admin"]["email"] else "User"
        meta = DocumentManagementSystem.save_document(
            fp, self.unit_var.get(), self.current_date_label.cget("text"),
            "Solar", "Solar", self.controller.email, role
        )
        self.document_var.set(meta["file_path"])
        messagebox.showinfo("File Saved", f"Saved to:\n{meta['file_path']}")

    def edit_record(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a record to edit")
            return
        values = self.tree.item(selected_items[0])["values"]
        dialog = EditSolarDialog(self, values)
        self.wait_window(dialog)
        if dialog.result:
            self.refresh_table()

    def delete_record(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a record to delete")
            return
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?"):
            return
        
        values = self.tree.item(selected_items[0])["values"]
        cfg = system_config["database"]["mssql"]
        conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
        if not conn:
            return
        try:
            cur = conn.cursor()
            sql = """
            DELETE FROM solar_energy_data 
            WHERE gmail = ? AND entry_date = ? AND company_unit = ? AND date = ?
            """
            # Use the utility to handle both string and date
            entry_date = to_date_str(values[1])
            date = to_date_str(values[5])
            cur.execute(sql, (values[0], entry_date, values[4], date))
            conn.commit()
            messagebox.showinfo("Success", "Record deleted successfully")
            self.refresh_table()
        except Exception as e:
            logging.error("Error deleting solar data: " + str(e))
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")
        finally:
            conn.close()

    def on_treeview_double_click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item)["values"]
        document_path = values[-1]
        if document_path and os.path.exists(document_path):
            os.startfile(document_path)
        else:
            messagebox.showwarning("Warning", "Document not found")

    def export_to_excel(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Export to Excel"
        )
        if not filename:
            return
            
        try:
            data = []
            columns = [
                "Gmail", "Entry Date", "Month", "Year", "Company Unit", "Date",
                "Inverter 1", "Inverter 2", "Inverter 3", "Inverter 4",
                "Old Total", "New Solar Inverter", "Total Generated",
                "Unit Type", "Remark", "Document"
            ]
            
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                data.append(values)
                
            df = pd.DataFrame(data, columns=columns)
            df.to_excel(filename, index=False)
            messagebox.showinfo("Success", "Data exported successfully")
            
        except Exception as e:
            logging.error("Error exporting to Excel: " + str(e))
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")

    def submit_data(self):
        try:
            # Validate required fields
            if not all([self.date_var.get().strip(), self.unit_var.get().strip(),
                       self.month_var.get().strip(), self.year_var.get().strip()]):
                messagebox.showerror("Error", "Please fill all required fields.")
                return

            if not self.document_var.get():
                messagebox.showerror("Error", "Please upload a document.")
                return

            # Get values
            values = [
                self.controller.email,  # Gmail
                self.current_date_label.cget("text"),  # Entry Date
                self.month_var.get(),  # Month
                self.year_var.get(),  # Year
                self.unit_var.get(),  # Company Unit
                self.date_var.get(),  # Date
                self.inverter_vars[0].get() or "0",  # Inverter 1
                self.inverter_vars[1].get() or "0",  # Inverter 2
                self.inverter_vars[2].get() or "0",  # Inverter 3
                self.inverter_vars[3].get() or "0",  # Inverter 4
                self.old_total_var.get() or "0",  # Old total
                self.new_solar_var.get() or "0",  # New Solar Inverter
                self.total_generated_var.get() or "0",  # Total Generated
                self.unit_var_solar.get(),  # Unit
                self.remark_var.get(),  # Remark
                self.document_var.get()  # Document
            ]

            # Save to database
            cfg = system_config["database"]["mssql"]
            conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
            if not conn:
                return

            try:
                cur = conn.cursor()
                insert_sql = """
                INSERT INTO solar_energy_data (
                    gmail, entry_date, month, year, company_unit, date,
                    inverter1, inverter2, inverter3, inverter4,
                    old_total, new_solar_inverter, total_generated,
                    unit_type, remark, document
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                cur.execute(insert_sql, values)
                conn.commit()
                messagebox.showinfo("Success", "Solar energy data saved successfully!")
                self.clear_form()
            except Exception as e:
                logging.error("Error saving solar data: " + str(e))
                messagebox.showerror("Error", f"Failed to save data: {str(e)}")
            finally:
                conn.close()

        except Exception as e:
            logging.error("Solar data submission error: " + str(e))
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def clear_form(self):
        self.date_var.set("")
        for var in self.inverter_vars:
            var.set("")
        self.old_total_var.set("")
        self.new_solar_var.set("")
        self.total_generated_var.set("")
        self.remark_var.set("")
        self.document_var.set("")

    def apply_filters(self):
        cfg = system_config["database"]["mssql"]
        conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
        if not conn:
            return
        try:
            cur = conn.cursor()
            where_clauses = []
            params = []
            if self.filter_vars["unit"].get() != "All":
                where_clauses.append("company_unit = ?")
                params.append(self.filter_vars["unit"].get())
            if self.filter_vars["month"].get() != "All":
                where_clauses.append("month = ?")
                params.append(self.filter_vars["month"].get())
            if self.filter_vars["year"].get() != "All":
                where_clauses.append("year = ?")
                params.append(self.filter_vars["year"].get())

            where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"
            sql = f"""
            SELECT gmail, entry_date, month, year, company_unit, date,
                   inverter1, inverter2, inverter3, inverter4,
                   old_total, new_solar_inverter, total_generated,
                   unit_type, remark, document
            FROM solar_energy_data
            WHERE {where_sql}
            ORDER BY entry_date DESC
            """
            cur.execute(sql, params)
            rows = cur.fetchall()
            for item in self.tree.get_children():
                self.tree.delete(item)
            for row in rows:
                # Convert all date/datetime fields to string
                row_list = list(row)
                # entry_date (index 1) and date (index 5) may be date/datetime
                for idx in [1, 5]:
                    if hasattr(row_list[idx], 'strftime'):
                        row_list[idx] = row_list[idx].strftime("%Y-%m-%d")
                self.tree.insert("", "end", values=row_list)
        except Exception as e:
            logging.error("Error filtering solar data: " + str(e))
            messagebox.showerror("Error", f"Failed to filter data: {str(e)}")
        finally:
            conn.close()

    def refresh_table(self):
        cfg = system_config["database"]["mssql"]
        conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
        if not conn:
            return
        try:
            cur = conn.cursor()
            sql = """
            SELECT gmail, entry_date, month, year, company_unit, date,
                   inverter1, inverter2, inverter3, inverter4,
                   old_total, new_solar_inverter, total_generated,
                   unit_type, remark, document
            FROM solar_energy_data
            ORDER BY entry_date DESC
            """
            cur.execute(sql)
            rows = cur.fetchall()
            for item in self.tree.get_children():
                self.tree.delete(item)
            for row in rows:
                # Convert all date/datetime fields to string
                row_list = list(row)
                # entry_date (index 1) and date (index 5) may be date/datetime
                for idx in [1, 5]:
                    if hasattr(row_list[idx], 'strftime'):
                        row_list[idx] = row_list[idx].strftime("%Y-%m-%d")
                self.tree.insert("", "end", values=row_list)
        except Exception as e:
            logging.error("Error refreshing solar data: " + str(e))
            messagebox.showerror("Error", f"Failed to refresh data: {str(e)}")
        finally:
            conn.close()

class EditSolarDialog(tk.Toplevel):
    def __init__(self, parent_page, values):
        super().__init__(parent_page)
        self.title("Edit Solar Record")
        self.parent_page = parent_page
        self.values = values

        # Create form
        form_frame = tk.Frame(self)
        form_frame.pack(padx=10, pady=10)

        # Labels and entries
        labels = ["Date:", "Inverter 1:", "Inverter 2:", "Inverter 3:", "Inverter 4:",
                 "Old Total:", "New Solar Inverter:", "Total Generated:", "Remark:", "Document:"]
        
        self.date_var = tk.StringVar(value=values[5])
        self.inverter_vars = [tk.StringVar(value=values[6+i]) for i in range(4)]
        self.old_total_var = tk.StringVar(value=values[10])
        self.new_solar_var = tk.StringVar(value=values[11])
        self.total_generated_var = tk.StringVar(value=values[12])
        self.remark_var = tk.StringVar(value=values[14])
        self.document_var = tk.StringVar(value=values[15])
        
        # Date entry
        tk.Label(form_frame, text=labels[0]).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(form_frame, textvariable=self.date_var, width=20).grid(row=0, column=1, padx=5, pady=5)
        
        # Inverter entries
        for i in range(4):
            tk.Label(form_frame, text=labels[i+1]).grid(row=i+1, column=0, padx=5, pady=5, sticky="e")
            entry = NumericEntry(form_frame, textvariable=self.inverter_vars[i], width=20)
            entry.grid(row=i+1, column=1, padx=5, pady=5)
            entry.bind('<KeyRelease>', self.update_totals)
            
        # Old total (read-only)
        tk.Label(form_frame, text=labels[5]).grid(row=5, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(form_frame, textvariable=self.old_total_var, width=20, state='readonly').grid(row=5, column=1, padx=5, pady=5)
        
        # New solar inverter
        tk.Label(form_frame, text=labels[6]).grid(row=6, column=0, padx=5, pady=5, sticky="e")
        entry = NumericEntry(form_frame, textvariable=self.new_solar_var, width=20)
        entry.grid(row=6, column=1, padx=5, pady=5)
        entry.bind('<KeyRelease>', self.update_totals)
        
        # Total generated (read-only)
        tk.Label(form_frame, text=labels[7]).grid(row=7, column=0, padx=5, pady=5, sticky="e")
        ttk.Entry(form_frame, textvariable=self.total_generated_var, width=20, state='readonly').grid(row=7, column=1, padx=5, pady=5)
        
        # Remark
        tk.Label(form_frame, text=labels[8]).grid(row=8, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(form_frame, textvariable=self.remark_var, width=20).grid(row=8, column=1, padx=5, pady=5)
        
        # Document
        tk.Label(form_frame, text=labels[9]).grid(row=9, column=0, padx=5, pady=5, sticky="e")
        doc_frame = tk.Frame(form_frame)
        doc_frame.grid(row=9, column=1, padx=5, pady=5, sticky="w")
        tk.Entry(doc_frame, textvariable=self.document_var, width=20).pack(side="left")
        tk.Button(doc_frame, text="Upload", command=self.upload_document,
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,10)).pack(side="left", padx=5)

        # Calculate initial totals
        self.update_totals()
        
        # Save button
        tk.Button(form_frame, text="Save Changes", command=self.save_changes,
                  bg=PRIMARY_COLOR, fg="white", font=(FONT_FAMILY,12), padx=10, pady=5).grid(row=10, column=0, columnspan=2, pady=10)

    def update_totals(self, *args):
        """Update the Old total and Total Generated fields based on inverter values"""
        try:
            # Calculate old total (sum of inverters)
            inverter_values = []
            for var in self.inverter_vars:
                try:
                    value = float(var.get() or 0)
                    inverter_values.append(value)
                except ValueError:
                    inverter_values.append(0)
            
            old_total = sum(inverter_values)
            self.old_total_var.set(f"{old_total:.2f}")

            # Calculate total generated (old total + new solar inverter)
            try:
                new_solar = float(self.new_solar_var.get() or 0)
            except ValueError:
                new_solar = 0

            total_generated = old_total + new_solar
            self.total_generated_var.set(f"{total_generated:.2f}")

        except Exception as e:
            logging.error(f"Error updating totals in edit dialog: {str(e)}")

    def upload_document(self):
        fp = filedialog.askopenfilename(
            filetypes=[
                ("All files","*.*"),
                ("PDF","*.pdf"),
                ("Excel","*.xlsx;*.xls"),
                ("Images","*.png;*.jpg;*.jpeg")
            ],
            title="Select document"
        )
        if not fp:
            return
        role = "Admin" if self.parent_page.controller.email == system_config["users"]["admin"]["email"] else "User"
        meta = DocumentManagementSystem.save_document(
            fp, self.values[4], self.values[1],  # unit and entry_date
            "Solar", "Solar", self.parent_page.controller.email, role
        )
        self.document_var.set(meta["file_path"])
        messagebox.showinfo("File Saved", f"Saved to:\n{meta['file_path']}")

    def save_changes(self):
        try:
            # Calculate old total and total generated
            self.update_totals()
            
            cfg = system_config["database"]["mssql"]
            conn = connect_mssql(cfg["server"], cfg["database"], cfg["user"], cfg["password"])
            if not conn:
                return

            try:
                cur = conn.cursor()
                update_sql = """
                UPDATE solar_energy_data
                SET date = ?, inverter1 = ?, inverter2 = ?, inverter3 = ?, inverter4 = ?,
                    old_total = ?, new_solar_inverter = ?, total_generated = ?,
                    remark = ?, document = ?
                WHERE gmail = ? AND entry_date = ? AND company_unit = ? AND date = ?
                """
                params = (
                    self.date_var.get(),  # date
                    self.inverter_vars[0].get() or "0",  # inverter1
                    self.inverter_vars[1].get() or "0",  # inverter2
                    self.inverter_vars[2].get() or "0",  # inverter3
                    self.inverter_vars[3].get() or "0",  # inverter4
                    self.old_total_var.get(),  # old_total
                    self.new_solar_var.get() or "0",  # new_solar_inverter
                    self.total_generated_var.get(),  # total_generated
                    self.remark_var.get(),  # remark
                    self.document_var.get(),  # document
                    self.values[0],  # gmail
                    self.values[1],  # entry_date
                    self.values[4],  # company_unit
                    self.values[5]   # original date
                )
                cur.execute(update_sql, params)
                conn.commit()
                self.parent_page.refresh_table()
                messagebox.showinfo("Success", "Record updated successfully.")
                self.destroy()
            except Exception as e:
                logging.error("Error updating solar data: " + str(e))
                messagebox.showerror("Error", f"Failed to update record: {str(e)}")
            finally:
                conn.close()
        except Exception as e:
            logging.error("Solar data update error: " + str(e))
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Utility for date string conversion

def to_date_str(val):
    if isinstance(val, str):
        return val  # assume already in correct format
    elif hasattr(val, 'strftime'):
        return val.strftime("%Y-%m-%d")
    else:
        return str(val)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
