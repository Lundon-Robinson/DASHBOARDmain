import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import traceback
import datetime

# Defensive import so if pywin32 is missing we log a clear error and raise.
try:
    import win32com.client as win32
except Exception as ex:
    msg = "pywin32 (win32com) is required by Bulk Mail.py but is not available in this Python environment."
    try:
        with open(r"C:\Users\NADLUROB\Desktop\Dash\bulk_mail_import_error.txt", "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n{traceback.format_exc()}\n")
    except Exception:
        pass
    raise ImportError(msg) from ex

# --- Configuration ---
EXCEL_PATH = r"\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\FAS DSC\Purchase Cards info\Card Holder List\Purchase cardholder list DSC.xls"
FALLBACK_PATHS = [
    r"/mnt/data/Purchase cardholder list DHSC.xlsx",
    os.path.join(os.getcwd(), "Purchase cardholder list DHSC.xlsx"),
]
SHEET_NAME = "OUTSTANDING LOGS"
LOG_FILE = r"C:\Users\NADLUROB\Desktop\Dash\log.txt"

# ---------- Logging helpers ----------
def log_uncaught_exceptions(exctype, value, tb):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("\n[UNCAUGHT EXCEPTION] {}\n".format(datetime.datetime.now()))
            f.write(''.join(traceback.format_exception(exctype, value, tb)))
            f.write("\n" + "-" * 60 + "\n")
    except Exception:
        pass
    sys.__excepthook__(exctype, value, tb)

sys.excepthook = log_uncaught_exceptions

def log_error(exc: Exception):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] ERROR:\n")
            f.write(''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
            f.write("\n" + "-" * 60 + "\n")
    except Exception:
        pass

# ---------- Utility helpers ----------
def normalize_string(s: str) -> str:
    """Normalize strings for matching: lowercase, no spaces or hyphens."""
    return re.sub(r'[\s\-]', '', (s or "").lower())

def parse_emails(email_cell: str):
    """
    Parse a semicolon-separated string of emails into (to_email, cc_email_str).
    Returns empty strings if input is falsy.
    """
    if not email_cell:
        return "", ""
    if not isinstance(email_cell, str):
        email_cell = str(email_cell)
    parts = [e.strip() for e in email_cell.split(";") if e.strip()]
    to_email = parts[0] if parts else ""
    cc_emails = parts[1:] if len(parts) > 1 else []
    cc_email_str = "; ".join(cc_emails)
    return to_email, cc_email_str

# Replacement helper (auto-replace placeholders at send time)
def perform_replacements(template: str, row_map: dict):
    """
    Replace placeholders of form <placeholder> in template using row_map (keys lowercased).
    Handles generic date/time placeholders and a few derived ones (e.g., <card_last4>).
    Missing values replaced with empty string.
    """
    if not template:
        return ""
    # Build lookup dict with lower keys
    lookup = {k.lower(): ("" if v is None else str(v)) for k, v in (row_map or {}).items()}

    # derived fields
    # full_name, first_name, last_name
    full = lookup.get("full_name") or (lookup.get("first_name", "") + (" " + lookup.get("last_name", "") if lookup.get("last_name") else ""))
    lookup.setdefault("full_name", full)
    if "first_name" not in lookup and "name" in lookup:
        # try to split name
        parts = lookup["name"].split()
        if parts:
            lookup.setdefault("first_name", parts[0])
            lookup.setdefault("last_name", " ".join(parts[1:]) if len(parts) > 1 else "")

    # card_last4
    cardnum = lookup.get("card_number") or lookup.get("cardno") or lookup.get("card")
    if cardnum:
        digits = re.sub(r'\D', '', cardnum)
        lookup.setdefault("card_last4", digits[-4:] if len(digits) >= 4 else digits)

    # email fields normalization
    # cardholder & manager might be together in one field (we store them explicitly in mapping if available)
    # date/time generics
    now = datetime.datetime.now()
    lookup.setdefault("today", now.strftime("%d/%m/%Y"))
    lookup.setdefault("full_date", now.strftime("%d %B %Y"))
    lookup.setdefault("time", now.strftime("%H:%M:%S"))
    lookup.setdefault("timestamp", now.isoformat())
    lookup.setdefault("year", now.strftime("%Y"))
    lookup.setdefault("month", now.strftime("%m"))
    lookup.setdefault("weekday", now.strftime("%A"))

    # perform replacement: find all <...> tokens
    def repl(match):
        key = match.group(1).strip().lower()
        return lookup.get(key, "")

    return re.sub(r'<([^<>]+)>', repl, template)

# ---------- Main application ----------
class BulkMailerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Purchase Card Bulk Mailer")
        # initial geometry chosen to resemble Outlook New Email proportionally
        self.root.geometry("1000x760")

        # Font styling state
        self.font_family_var = tk.StringVar(value="Segoe UI")
        self.font_size_var = tk.StringVar(value="10")
        self.font_bold = False
        self.font_italic = False
        self.font_underline = False

        self.urgent_var = tk.BooleanVar()
        self.common_attachments = []  # Attach to all emails
        self.folder_attachments = []  # Files from selected folder
        self.recipients = []          # list of dicts with row_map and meta
        self._detected_headers = []   # headers detected from workbook
        self.variables = []           # list of placeholders
        self._last_focus_widget = None

        try:
            self._build_ui()
            self._load_data_from_excel()
            self.variables = self.__class__._generate_variable_list(self._detected_headers)
            # create the variables window lazily (button opens it)
            self._apply_dark_theme()
            self._apply_body_font()
        except Exception as e:
            log_error(e)
            if getattr(self.root, "winfo_exists", lambda: False)():
                messagebox.showerror("Startup Error", f"An error occurred:\n{e}")
                try:
                    self.root.destroy()
                except Exception:
                    pass

    # ---------- UI build (Outlook-like layout: To, Cc, Subject, Body) ----------
    def _build_ui(self):
        # Ribbon-like top bar
        ribbon = ttk.Frame(self.root, relief="raised", padding=4)
        ribbon.pack(fill="x", side="top")

        # Basic font controls similar to your previous controls
        ttk.Label(ribbon, text="Font:", padding=(5, 0)).pack(side="left")
        families = list(font.families())
        families.sort()
        font_cb = ttk.Combobox(ribbon, textvariable=self.font_family_var,
                               values=families, width=22, state="readonly")
        font_cb.pack(side="left", padx=3)
        ttk.Label(ribbon, text="Size:", padding=(5, 0)).pack(side="left")
        size_values = [str(i) for i in range(8, 25)]
        size_cb = ttk.Combobox(ribbon, textvariable=self.font_size_var,
                               values=size_values, width=3, state="readonly")
        size_cb.pack(side="left", padx=3)
        font_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_body_font())
        size_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_body_font())
        ttk.Button(ribbon, text="B", width=3, command=self._toggle_bold).pack(side="left", padx=2)
        ttk.Button(ribbon, text="I", width=3, command=self._toggle_italic).pack(side="left", padx=2)
        ttk.Button(ribbon, text="U", width=3, command=self._toggle_underline).pack(side="left", padx=2)

        # Variables button (opens searchable, scrollable window)
        ttk.Button(ribbon, text="Variables", command=self._open_variables_window).pack(side="left", padx=10)

        # Urgent and attachments on right
        ttk.Checkbutton(ribbon, text="Mark as Urgent", variable=self.urgent_var).pack(side="right", padx=12)
        ttk.Button(ribbon, text="Attach Files", command=self._attach_files).pack(side="right", padx=6)
        ttk.Button(ribbon, text="Select Folder", command=self._select_folder).pack(side="right", padx=6)

        # Main area: left recipients list, right compose area arranged like Outlook
        mainfrm = ttk.Frame(self.root, padding=10)
        mainfrm.pack(fill="both", expand=True)

        # Left column: recipients list
        left_col = ttk.Frame(mainfrm)
        left_col.grid(row=0, column=0, sticky="nswe", padx=(0,10))
        ttk.Label(left_col, text="Recipients (select to send)").pack(anchor="w")
        self.lst = tk.Listbox(left_col, selectmode="extended", height=34, width=36, font=("Segoe UI", 10))
        self.lst.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(left_col, orient="vertical", command=self.lst.yview)
        self.lst.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # Right column: compose area (Outlook-like)
        right_col = ttk.Frame(mainfrm)
        right_col.grid(row=0, column=1, sticky="nsew")
        mainfrm.columnconfigure(1, weight=1)
        mainfrm.rowconfigure(0, weight=1)

        # To, Cc, Subject fields (Outlook order)
        row_frame = ttk.Frame(right_col)
        row_frame.pack(fill="x", pady=(0,6))
        ttk.Label(row_frame, text="To:").grid(row=0, column=0, sticky="w")
        self.to_entry = ttk.Entry(row_frame, width=90)
        self.to_entry.grid(row=0, column=1, sticky="we", padx=(6,0))
        ttk.Label(row_frame, text="Cc:").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.cc_entry = ttk.Entry(row_frame, width=90)
        self.cc_entry.grid(row=1, column=1, sticky="we", padx=(6,0), pady=(6,0))
        ttk.Label(row_frame, text="Subject:").grid(row=2, column=0, sticky="w", pady=(6,0))
        self.subj = ttk.Entry(row_frame, width=90)
        self.subj.grid(row=2, column=1, sticky="we", padx=(6,0), pady=(6,0))

        # body area
        body_frame = ttk.Frame(right_col)
        body_frame.pack(fill="both", expand=True)
        self.body = tk.Text(body_frame, width=80, height=25, wrap="word")
        self.body.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(body_frame, orient="vertical", command=self.body.yview)
        scrollbar.pack(side="right", fill="y")
        self.body.configure(yscrollcommand=scrollbar.set)

        # footer: signature note and send button
        footer = ttk.Frame(right_col)
        footer.pack(fill="x", pady=(6,0))
        ttk.Label(footer, text="Signature will be appended automatically to the emails.").pack(side="left")
        ttk.Button(footer, text="Send Emails", command=self._send_emails).pack(side="right")

        # recipients count
        self.recipients_count_label = ttk.Label(left_col, text="Recipients: 0")
        self.recipients_count_label.pack(anchor="ne", pady=(6,0))

        # Track focus for variable insertion
        self._last_focus_widget = self.body
        self.subj.bind("<FocusIn>", lambda e: setattr(self, "_last_focus_widget", self.subj))
        self.body.bind("<FocusIn>", lambda e: setattr(self, "_last_focus_widget", self.body))
        self.to_entry.bind("<FocusIn>", lambda e: setattr(self, "_last_focus_widget", self.to_entry))
        self.cc_entry.bind("<FocusIn>", lambda e: setattr(self, "_last_focus_widget", self.cc_entry))

    # ---------- Theme & font ----------
    def _apply_dark_theme(self):
        bg_color = "#121212"
        entry_bg = "#1e1e1e"
        fg_color = "#e0e0e0"
        border_color = "#00bfff"
        highlight_thickness = 1
        style_font = ("Segoe UI", 10)

        self.root.configure(background=bg_color)
        style = ttk.Style()
        try:
            style.theme_use("default")
        except Exception:
            pass
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=fg_color, font=style_font)
        style.configure("TButton", background=bg_color, foreground=fg_color, font=style_font, borderwidth=0)
        style.map("TButton", background=[("active", "#007acc")])
        style.configure("TCheckbutton", background=bg_color, foreground=fg_color)

        try:
            self.body.configure(bg=entry_bg, fg=fg_color, insertbackground=fg_color, relief="flat", borderwidth=0,
                                highlightthickness=highlight_thickness, highlightbackground=border_color,
                                highlightcolor=border_color)
            self.lst.configure(bg=entry_bg, fg=fg_color, selectbackground="#007acc", selectforeground="white",
                               relief="flat", borderwidth=0, highlightthickness=highlight_thickness,
                               highlightbackground=border_color, highlightcolor=border_color)
        except Exception:
            pass

    def _apply_body_font(self):
        underline = 1 if self.font_underline else 0
        try:
            size = 10
            try:
                size = int(self.font_size_var.get())
            except Exception:
                size = 10
            fnt = font.Font(family=self.font_family_var.get(),
                            size=size,
                            weight="bold" if self.font_bold else "normal",
                            slant="italic" if self.font_italic else "roman",
                            underline=underline)
            self.body.configure(font=fnt)
            # Entry widgets font sync
            fnt_entry = font.Font(family=self.font_family_var.get(), size=max(10, size-1))
            self.subj.configure(font=fnt_entry)
            self.to_entry.configure(font=fnt_entry)
            self.cc_entry.configure(font=fnt_entry)
        except Exception as e:
            log_error(e)

    def _toggle_bold(self):
        self.font_bold = not self.font_bold
        self._apply_body_font()

    def _toggle_italic(self):
        self.font_italic = not self.font_italic
        self._apply_body_font()

    def _toggle_underline(self):
        self.font_underline = not self.font_underline
        self._apply_body_font()

    # ---------- Excel loading & header detection (and store per-row maps) ----------
    def _load_data_from_excel(self):
        """
        Load recipients and full row data from workbook. Column 8 (H) contains cardholder & manager emails separated by ';'.
        The function detects the header row (1..10) then maps header->cell for each data row.
        Recipients list contains dicts:
            { 'row_map': {headername: value, ...}, 'to_email': ..., 'cc_email': ..., 'display': 'Name (email)' }
        """
        excel = None
        wb = None
        found_path = None
        try:
            # locate workbook
            candidates = [EXCEL_PATH] + FALLBACK_PATHS
            for p in candidates:
                try:
                    if p and os.path.exists(p):
                        found_path = p
                        break
                except Exception:
                    continue

            if not found_path:
                found_path = filedialog.askopenfilename(title="Select Purchase cardholder workbook",
                                                        filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm"), ("All files","*.*")])
                if not found_path:
                    self._detected_headers = []
                    self.recipients = []
                    self.lst.delete(0, tk.END)
                    self.recipients_count_label.config(text=f"Recipients: {0}")
                    return

            excel = win32.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(found_path, ReadOnly=True)
            ws = None
            for sname in (SHEET_NAME, "Sheet1"):
                try:
                    ws = wb.Worksheets(sname)
                    break
                except Exception:
                    ws = None
            if ws is None:
                ws = wb.Worksheets(1)

            used_range = ws.UsedRange
            last_row = int(used_range.Rows.Count)
            last_col = int(used_range.Columns.Count)

            # detect header row
            header_row = 1
            max_nonempty = 0
            for r in range(1, min(10, last_row) + 1):
                nonempty = 0
                for c in range(1, last_col + 1):
                    try:
                        val = ws.Cells(r, c).Value
                        if val not in (None, ""):
                            nonempty += 1
                    except Exception:
                        continue
                if nonempty > max_nonempty:
                    max_nonempty = nonempty
                    header_row = r

            # collect headers aligned to columns 1..last_col
            headers = []
            for c in range(1, last_col + 1):
                try:
                    v = ws.Cells(header_row, c).Value
                    if v not in (None, ""):
                        headers.append(str(v).strip())
                    else:
                        headers.append(f"col_{c:02d}")
                except Exception:
                    headers.append(f"col_{c:02d}")
            self._detected_headers = headers

            # iterate data rows and build recipient entries
            recipients = []
            for row in range(header_row + 1, last_row + 1):
                try:
                    send_flag = ws.Cells(row, 1).Value
                except Exception:
                    send_flag = None
                if not send_flag or str(send_flag).strip().upper() != 'X':
                    continue

                # build row_map header->value
                row_map = {}
                for c in range(1, last_col + 1):
                    key = headers[c-1] if c-1 < len(headers) else f"col_{c:02d}"
                    try:
                        val = ws.Cells(row, c).Value
                    except Exception:
                        val = None
                    row_map_key = str(key).strip()
                    row_map[row_map_key] = val

                # specifically parse column 8 (H) for cardholder and manager emails (semicolons)
                col8_val = None
                try:
                    col8_val = ws.Cells(row, 8).Value
                except Exception:
                    col8_val = row_map.get(headers[7]) if len(headers) >= 8 else None
                cardholder_email, manager_email_part = parse_emails(col8_val)

                # combine manager_email_part into cc string (could include manager name or multiple)
                cc_email = manager_email_part or ""
                # if there are any existing cc from header names like 'CC' include them
                # attempt to find any column that looks like cc or manager; include if present
                potential_ccs = []
                for hname, hval in row_map.items():
                    lname = str(hname).lower()
                    if 'manager' in lname and isinstance(hval, str):
                        potential_ccs.append(hval)
                    if 'cc' == lname or 'cc_email' in lname or 'ccemail' in lname:
                        if isinstance(hval, str) and hval.strip():
                            potential_ccs.append(hval)
                if potential_ccs:
                    # append to cc_email
                    if cc_email:
                        cc_email = "; ".join([cc_email] + potential_ccs)
                    else:
                        cc_email = "; ".join(potential_ccs)

                # prepare display name: prefer first/surname columns, else fallback to 'name' headers
                display_name = ""
                fname = None
                lname = None

                # 1. Try first/last name columns by header text
                for k in row_map.keys():
                    kl = k.lower()
                    val = str(row_map[k] or "").strip()
                    # skip if value is all digits or fewer than 2 characters (avoids card numbers)
                    if val.isdigit() or len(val) < 2:
                        continue
                    if 'first' in kl and not fname:
                        fname = val
                    elif ('last' in kl or 'surname' in kl) and not lname:
                        lname = val

                # 2. If no first/last, try any header containing 'name'
                if not (fname or lname):
                    for k in row_map.keys():
                        if 'name' in k.lower():
                            val = str(row_map[k] or "").strip()
                            if not val.isdigit() and len(val) > 1:
                                display_name = val
                                break
                else:
                    display_name = (str(fname or "").strip() + " " + str(lname or "").strip()).strip()

                # 3. Fallback to specific Excel columns 6 & 7
                if not display_name:
                    try:
                        first_name = str(ws.Cells(row, 6).Value or "").strip()
                        surname = str(ws.Cells(row, 7).Value or "").strip()
                        if not first_name.isdigit() and not surname.isdigit():
                            display_name = f"{first_name} {surname}".replace("-", " ").strip()
                    except Exception:
                        val = str(row_map.get(headers[0]) or "").strip()
                        if not val.isdigit():
                            display_name = val

                # if still empty, last resort = unknown
                if not display_name:
                    display_name = "Unknown"

                # if no to_email skip
                if not cardholder_email:
                    continue

                # augment row_map with normalized keys for convenience
                norm_map = {}
                for k, v in row_map.items():
                    norm_map[k.lower()] = v
                    sk = re.sub(r'[^0-9a-z_]', '_', k.strip().lower())
                    norm_map[sk] = v

                # add derived email keys
                norm_map['to_email'] = cardholder_email
                norm_map['cc_email'] = cc_email
                norm_map['cardholder_email'] = cardholder_email
                norm_map['manager_email'] = cc_email

                # add name fields
                if display_name:
                    norm_map['name'] = display_name

                # first/last if not present, try to split display_name
                if 'first_name' not in norm_map or not norm_map.get('first_name'):
                    parts = display_name.split()
                    if parts:
                        norm_map.setdefault('first_name', parts[0])
                        norm_map.setdefault('last_name', " ".join(parts[1:]) if len(parts) > 1 else "")

                recipients.append({
                    'row_map': norm_map,
                    'to_email': cardholder_email,
                    'cc_email': cc_email,
                    'display': f"{display_name} ({cardholder_email})"
                })

            self.recipients = recipients
            self.lst.delete(0, tk.END)
            for rec in recipients:
                self.lst.insert(tk.END, rec['display'])
            self.recipients_count_label.config(text=f"Recipients: {len(recipients)}")
        except Exception as e:
            log_error(e)
            raise
        finally:
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            try:
                del wb
            except Exception:
                pass
            try:
                del excel
            except Exception:
                pass

    # ---------- Variable generation ----------
    @staticmethod
    def _generate_variable_list(headers):
        """
        From a list of header strings, generate up to 100 placeholder variables,
        sorted by likely usefulness.
        """
        try:
            placeholders = []
            seen = set()

            def add(ph):
                ph = ph.strip()
                if not ph:
                    return
                if ph not in seen:
                    seen.add(ph)
                    placeholders.append(ph)

            priority_generics = [
                "<name>", "<full_name>", "<first_name>", "<last_name>", "<title>",
                "<job_title>", "<email>", "<to_email>", "<cc_email>", "<manager_email>",
                "<card_number>", "<card_last4>", "<monthly_limit>", "<single_limit>",
                "<today>", "<full_date>", "<time>", "<timestamp>", "<year>", "<month>", "<weekday>",
                "<department>", "<cost_center>", "<costcode>", "<reference>", "<bank_reference>",
                "<amount>", "<currency>", "<sender_name>", "<sender_email>", "<signature>"
            ]
            for g in priority_generics:
                add(g)

            for h in (headers or []):
                s = str(h).strip()
                if not s:
                    continue
                base = re.sub(r'[^0-9A-Za-z]+', '_', s).strip('_').lower()
                if not base:
                    continue
                add(f"<{base}>")
                # add likely aliases
                if re.search(r'name|forename|given', base):
                    add("<full_name>")
                    add("<first_name>")
                    add("<last_name>")
                if 'email' in base:
                    add("<email>")
                if 'manager' in base or 'mgr' in base:
                    add("<manager_email>")
                if 'card' in base:
                    add("<card_number>")
                    add("<card_last4>")
                if 'limit' in base:
                    add("<monthly_limit>")
                    add("<single_limit>")
                if 'cost' in base or 'centre' in base or 'center' in base:
                    add("<cost_center>")
                    add("<costcode>")

            extras = [
                "<due_date>", "<start_date>", "<end_date>", "<period>", "<month_start>", "<month_end>",
                "<week_start>", "<week_end>", "<reporting_period>", "<approver>", "<approver_email>",
                "<organisation>", "<project_code>", "<project_name>", "<invoice_number>", "<invoice_date>",
                "<payment_date>", "<payment_ref>", "<payee>", "<phone>", "<mobile>", "<location>", "<notes>",
                "<status>", "<priority>", "<category>"
            ]
            for ex in extras:
                add(ex)
                if len(placeholders) >= 100:
                    break

            counter = 1
            while len(placeholders) < 100:
                add(f"<var_{counter:03d}>")
                counter += 1

            return placeholders[:100]
        except Exception as e:
            log_error(e)
            fallback = ["<name>", "<email>", "<card_number>", "<today>", "<full_date>", "<time>", "<var_001>"]
            i = 1
            while len(fallback) < 100:
                fallback.append(f"<var_fallback_{i:03d}>")
                i += 1
            return fallback[:100]

    # ---------- Variables window (searchable, scrollable) ----------
    def _open_variables_window(self):
        # create top-level where user can search and insert variables
        try:
            if hasattr(self, "_vars_win") and self._vars_win.winfo_exists():
                self._vars_win.lift()
                return

            self._vars_win = tk.Toplevel(self.root)
            self._vars_win.title("Insert Variable")
            self._vars_win.geometry("420x520")
            # don't block main app
            self._vars_win.transient(self.root)

            ttk.Label(self._vars_win, text="Search:").pack(anchor="nw", padx=8, pady=(8,0))
            search_var = tk.StringVar()
            search_entry = ttk.Entry(self._vars_win, textvariable=search_var)
            search_entry.pack(fill="x", padx=8, pady=(0,6))
            search_entry.focus_set()

            # listbox with scrollbar
            list_frame = ttk.Frame(self._vars_win)
            list_frame.pack(fill="both", expand=True, padx=8, pady=(0,8))

            vars_listbox = tk.Listbox(list_frame, height=20)
            vars_listbox.pack(side="left", fill="both", expand=True)
            vscroll = ttk.Scrollbar(list_frame, orient="vertical", command=vars_listbox.yview)
            vscroll.pack(side="right", fill="y")
            vars_listbox.configure(yscrollcommand=vscroll.set)

            # populate
            self.variables = self.__class__._generate_variable_list(self._detected_headers)
            for var in self.variables:
                vars_listbox.insert(tk.END, var)

            # double-click or button to insert
            def insert_selected():
                sel = vars_listbox.curselection()
                if not sel:
                    return
                var = vars_listbox.get(sel[0])
                self._insert_variable_at_focus(var)
                # do not close; allow multiple inserts

            insert_btn = ttk.Button(self._vars_win, text="Insert", command=insert_selected)
            insert_btn.pack(side="left", padx=8, pady=(0,8))

            close_btn = ttk.Button(self._vars_win, text="Close", command=self._vars_win.destroy)
            close_btn.pack(side="right", padx=8, pady=(0,8))

            # search behaviour
            def on_search_change(*_):
                q = search_var.get().strip().lower()
                vars_listbox.delete(0, tk.END)
                if not q:
                    for v in self.variables:
                        vars_listbox.insert(tk.END, v)
                else:
                    for v in self.variables:
                        if q in v.lower():
                            vars_listbox.insert(tk.END, v)

            search_var.trace_add("write", on_search_change)

            def on_dblclick(event):
                insert_selected()

            vars_listbox.bind("<Double-Button-1>", on_dblclick)
            # allow keyboard Enter to insert
            vars_listbox.bind("<Return>", lambda e: insert_selected())

        except Exception as e:
            log_error(e)
            messagebox.showerror("Variables error", f"Failed to open variables window: {e}")

    def _insert_variable_at_focus(self, var_text: str):
        """
        Insert var_text into the current focus widget (subject / body / to / cc).
        """
        try:
            target = getattr(self, "_last_focus_widget", None)
            current_focus = self.root.focus_get()
            if current_focus in (self.subj, self.body, self.to_entry, self.cc_entry):
                target = current_focus
            if target is None:
                target = self.body

            # Entries (ttk.Entry) behave like this
            if target in (self.subj, self.to_entry, self.cc_entry) or isinstance(target, ttk.Entry):
                try:
                    # try 'insert' index
                    target.insert("insert", var_text)
                    target.focus_set()
                except Exception:
                    # fallback manual manipulation
                    val = target.get()
                    try:
                        idx = target.index("insert")
                    except Exception:
                        idx = len(val)
                    new = val[:idx] + var_text + val[idx:]
                    target.delete(0, tk.END)
                    target.insert(0, new)
                    target.icursor(idx + len(var_text))
                    target.focus_set()
            else:
                # Text widget
                try:
                    target.insert("insert", var_text)
                    target.focus_set()
                except Exception:
                    self.body.insert(tk.END, var_text)
                    self.body.focus_set()
        except Exception as e:
            log_error(e)
            messagebox.showerror("Insert failed", f"Failed to insert variable {var_text}: {e}")

    # ---------- Attach / folder ----------
    def _attach_files(self):
        files = filedialog.askopenfilenames(title="Select files to attach to all emails")
        if files:
            self.common_attachments = list(files)
        else:
            self.common_attachments = []

    def _select_folder(self):
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return
        self.lbl_attach = getattr(self, "lbl_attach", None)
        if not self.lbl_attach:
            # lazy create a small label if not present
            self.lbl_attach = ttk.Label(self.root, text=folder_path)
            self.lbl_attach.place(x=10, y=10)
        self.lbl_attach.config(text=folder_path)
        # list files
        try:
            files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
            self.folder_attachments = [os.path.join(folder_path, f) for f in files]
        except Exception as e:
            log_error(e)
            messagebox.showerror("Error", f"Failed to list files in folder:\n{e}")
            self.folder_attachments = []
            return

    # ---------- Outlook helpers & send (with auto-replacement) ----------
    @staticmethod
    def _get_outlook_signature(outlook):
        try:
            tmp_mail = outlook.CreateItem(0)
            tmp_mail.Display()
            signature = tmp_mail.HTMLBody or ""
            try:
                tmp_mail.Close(0)
            except Exception:
                pass
            return signature
        except Exception:
            return ""

    def _send_emails(self):
        def clean_emails(email_str: str):
            if not email_str:
                return []
            emails = [e.strip() for e in email_str.split(";") if e.strip()]
            valid_emails = [e for e in emails if "@" in e and "." in e]
            return valid_emails

        selected_indices = self.lst.curselection()
        if not selected_indices:
            # Ask if user wants to send test email to yourself
            send_to_self = messagebox.askyesno(
                "No Recipients Selected",
                "No recipients selected. Do you want to send the email only to 'lundon.robinson@gov.im'?"
            )
            if not send_to_self:
                return  # Cancel sending
            # If yes, create a fake recipient list with only your email
            self.recipients = [{
                'name': 'Lundon Robinson',
                'to_email': 'lundon.robinson@gov.im',
                'cc_email': ''
            }]
            selected_indices = [0]  # Ensure loop will process this recipient

        try:
            outlook = win32.Dispatch("Outlook.Application")
        except Exception as e:
            log_error(e)
            messagebox.showerror("Outlook Error", "Outlook not installed or cannot start.")
            return

        errors = []
        sent_count = 0

        # Grab templates from UI once
        subj_template = self.subj.get() or ""
        body_template = self.body.get("1.0", "end-1c") or ""

        for idx in selected_indices:
            try:
                rec = self.recipients[idx]
            except Exception as e:
                log_error(e)
                errors.append(f"Index error for selected index {idx}: {e}")
                continue

            try:
                mail = outlook.CreateItem(0)  # MailItem

                # prepare recipient-specific row_map
                row_map = rec.get('row_map', {})
                to_emails = clean_emails(rec.get('to_email', "") or row_map.get('to_email', "") or row_map.get('cardholder_email', ""))
                if not to_emails:
                    errors.append(f"No valid To email for recipient {rec.get('display', '<unknown>')}")
                    continue
                mail.To = to_emails[0]

                cc_emails = clean_emails(rec.get('cc_email', "") or row_map.get('cc_email', ""))
                cc_emails = [e for e in cc_emails if e.lower() != mail.To.lower()]
                mail.CC = "; ".join(cc_emails)

                # replacements
                subject_final = perform_replacements(subj_template, row_map)
                body_text_final = perform_replacements(body_template, row_map)

                if self.urgent_var.get():
                    try:
                        mail.Importance = 2  # High importance
                    except Exception:
                        pass

                signature = self.__class__._get_outlook_signature(outlook)
                body_html = body_text_final.replace('\n', '<br>') + "<br><br>" + signature
                mail.Subject = subject_final
                mail.HTMLBody = body_html

                # attachments
                for f in self.common_attachments:
                    try:
                        if os.path.isfile(f):
                            mail.Attachments.Add(f)
                    except Exception:
                        log_error(Exception(f"Attachment failed: {f}"))

                # folder attachments: attach only if filename matches recipient name heuristics
                recipient_norm = normalize_string(row_map.get('name', '') or row_map.get('full_name', ''))
                for fpath in self.folder_attachments:
                    try:
                        fname_norm = normalize_string(os.path.basename(fpath))
                        if recipient_norm and (recipient_norm in fname_norm or fname_norm in recipient_norm):
                            if os.path.isfile(fpath):
                                mail.Attachments.Add(fpath)
                    except Exception:
                        log_error(Exception(f"Folder attachment check failed: {fpath}"))

                mail.Send()
                sent_count += 1

            except Exception as e:
                log_error(e)
                errors.append(f"Failed to send to {rec.get('display', '')}: {e}")

        summary_msg = f"Emails sent: {sent_count}\n"
        if errors:
            summary_msg += f"\nErrors:\n" + "\n".join(errors)
            messagebox.showwarning("Send Completed with Errors", summary_msg)
        else:
            messagebox.showinfo("Send Completed", summary_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = BulkMailerApp(root)
    root.mainloop()
