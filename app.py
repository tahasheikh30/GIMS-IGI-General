"""
Outstanding Balance Reporting & Client Email System
Desktop GUI (Tkinter) - Finance department automation tool.

Run with:  python app.py
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pandas as pd

import data_processor as dp
import report_generator as rg
import email_generator as eg
import outlook_sender as osend
import pivot_native_generator as pivot_native
import run_log


class BalanceApp(tk.Tk):
    # Palette
    COLOR_BG = "#EAF0F6"       # soft blue-grey app background (was flat light grey)
    COLOR_PANEL = "#FFFFFF"
    COLOR_PANEL_ALT = "#F7FAFC"  # subtle tint for nested/alternating panels
    COLOR_PRIMARY = "#1F4E78"
    COLOR_PRIMARY_DARK = "#163A5C"
    COLOR_ACCENT = "#2E86AB"
    COLOR_SUCCESS = "#2E7D32"
    COLOR_WARN = "#B45309"
    COLOR_TEXT = "#1C2733"
    COLOR_MUTED = "#5B6B79"
    COLOR_BORDER = "#D3DEE8"
    COLOR_SIDEBAR = "#F7FAFC"       # sidebar background - close to panel-alt, distinct from main canvas
    COLOR_SIDEBAR_ACTIVE = "#E4EDF5"  # active nav item background tint

    def __init__(self):
        super().__init__()
        self.title("GIMS - General Insurance Management System")
        self.geometry("1300x760")
        self.minsize(1050, 620)
        self.configure(bg=self.COLOR_BG)

        self.raw_df = None
        self.clean_df = None
        self.issues_df = None
        self.filtered_df = None
        self.grand_totals = None
        self.qualifying = None
        self.email_lookup = {}
        self.poc_lookup = {}
        self.drafts = []
        self.output_dir = os.path.join(os.getcwd(), "output")

        self._setup_style()
        self._build_ui()

    def _setup_style(self):
        style = ttk.Style(self)
        available = style.theme_names()
        for preferred in ("vista", "xpnative", "clam"):
            if preferred in available:
                try:
                    style.theme_use(preferred)
                    break
                except tk.TclError:
                    continue

        base_font = ("Segoe UI", 10)
        heading_font = ("Segoe UI Semibold", 11)
        title_font = ("Segoe UI Semibold", 16)

        style.configure(".", font=base_font, background=self.COLOR_BG)
        style.configure("TFrame", background=self.COLOR_BG)
        style.configure("Panel.TFrame", background=self.COLOR_PANEL)
        style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=base_font)
        style.configure("Panel.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT, font=base_font)
        style.configure("Muted.TLabel", background=self.COLOR_BG, foreground=self.COLOR_MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=self.COLOR_PRIMARY, foreground="white", font=title_font)
        style.configure("SectionHeader.TLabel", background=self.COLOR_BG, foreground=self.COLOR_PRIMARY, font=heading_font)

        style.configure("TLabelframe", background=self.COLOR_PANEL, bordercolor="#D6DCE1", relief="solid")
        style.configure("TLabelframe.Label", background=self.COLOR_PANEL, foreground=self.COLOR_PRIMARY, font=heading_font)

        current_theme = style.theme_use()
        using_native = current_theme in ("vista", "xpnative")

        style.configure("TNotebook", background=self.COLOR_BG, borderwidth=0)
        if using_native:
            # Native Windows themes already render clean, readable tabs -
            # forcing custom colors here tends to look worse, not better.
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 8))
        else:
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 10))
            style.map("TNotebook.Tab",
                      background=[("selected", self.COLOR_PANEL), ("!selected", "#E3E8EC")],
                      foreground=[("selected", self.COLOR_PRIMARY), ("!selected", self.COLOR_MUTED)])

        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 6))
        if using_native:
            # Native theme buttons don't reliably support background-color
            # overrides, so keep them native-styled but bold/larger for emphasis.
            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8))
            style.configure("Send.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8))
        else:
            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8),
                            background=self.COLOR_ACCENT, foreground="white")
            style.map("Accent.TButton", background=[("active", self.COLOR_PRIMARY_DARK), ("disabled", "#A9B7C2")])
            style.configure("Send.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8),
                            background=self.COLOR_SUCCESS, foreground="white")
            style.map("Send.TButton", background=[("active", "#1B5E20")])

        style.configure("Treeview", font=("Segoe UI", 10), rowheight=26,
                        background="white", fieldbackground="white", bordercolor="#D6DCE1")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background=self.COLOR_PRIMARY, foreground="white", relief="flat")
        style.map("Treeview.Heading", background=[("active", self.COLOR_PRIMARY_DARK)])
        style.map("Treeview", background=[("selected", self.COLOR_ACCENT)], foreground=[("selected", "white")])

        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("TCheckbutton", background=self.COLOR_BG, font=base_font)
        style.configure("Horizontal.TSeparator", background="#D6DCE1")

    # ---------- UI construction ----------
    def _build_ui(self):
        self._build_header()

        body = tk.Frame(self, bg=self.COLOR_BG)
        body.pack(fill="both", expand=True, side="top")

        self._build_sidebar(body)
        self._build_content_area(body)

        status_frame = tk.Frame(self, bg="#E3E8EC", height=30)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(status_frame, textvariable=self.status_var, bg="#E3E8EC", fg=self.COLOR_MUTED,
                 font=("Segoe UI", 9), anchor="w").pack(fill="both", padx=12, pady=4)

    def _build_header(self):
        """Fixed top header with the GIMS wordmark - stays visible across
        every sidebar section."""
        header = tk.Frame(self, bg=self.COLOR_PRIMARY, height=64)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        brand = tk.Frame(header, bg=self.COLOR_PRIMARY)
        brand.pack(side="left", padx=20, pady=10)

        # Simple drawn logo mark (no external image asset needed) - a
        # rounded accent-colored square with the "G" initial.
        mark = tk.Canvas(brand, width=40, height=40, bg=self.COLOR_PRIMARY, highlightthickness=0)
        mark.pack(side="left", padx=(0, 12))
        mark.create_oval(0, 0, 40, 40, fill=self.COLOR_ACCENT, outline="")
        mark.create_text(20, 21, text="G", fill="white", font=("Segoe UI", 17, "bold"))

        text_col = tk.Frame(brand, bg=self.COLOR_PRIMARY)
        text_col.pack(side="left")
        tk.Label(text_col, text="GIMS", font=("Segoe UI Semibold", 20, "bold"),
                 bg=self.COLOR_PRIMARY, fg="white").pack(anchor="w")
        tk.Label(text_col, text="General Insurance Management System", font=("Segoe UI", 9),
                 bg=self.COLOR_PRIMARY, fg="#BFD4E6").pack(anchor="w")

        tk.Label(header, text="General Department", font=("Segoe UI", 10),
                 bg=self.COLOR_PRIMARY, fg="#CFE0EE").pack(side="right", padx=20)

    def _build_sidebar(self, parent):
        """Persistent left sidebar with the three module sections. The
        active section is shown with a left accent bar, bold text, and a
        tinted background."""
        sidebar = tk.Frame(parent, bg=self.COLOR_SIDEBAR, width=232)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=18).pack(fill="x")
        tk.Label(sidebar, text="MODULES", font=("Segoe UI", 8, "bold"), bg=self.COLOR_SIDEBAR,
                 fg=self.COLOR_MUTED).pack(anchor="w", padx=20, pady=(0, 6))

        nav_labels = [
            ("outstanding_premium", "Outstanding Premium"),
            ("loss_ratio", "Loss Ratio"),
            ("renewal", "Renewal"),
        ]
        self.nav_items = {}
        for key, label in nav_labels:
            self.nav_items[key] = self._make_nav_item(sidebar, key, label)

        # Thin divider between the sidebar and content area
        tk.Frame(parent, bg=self.COLOR_BORDER, width=1).pack(side="left", fill="y")

    def _make_nav_item(self, parent, key, text):
        wrapper = tk.Frame(parent, bg=self.COLOR_SIDEBAR, cursor="hand2")
        wrapper.pack(fill="x")

        accent = tk.Frame(wrapper, bg=self.COLOR_SIDEBAR, width=4)
        accent.pack(side="left", fill="y")

        label = tk.Label(wrapper, text=text, font=("Segoe UI", 11), bg=self.COLOR_SIDEBAR,
                          fg=self.COLOR_MUTED, anchor="w")
        label.pack(side="left", fill="both", expand=True, padx=(14, 10), pady=14)

        def on_click(event=None, k=key):
            self.select_nav(k)

        for widget in (wrapper, accent, label):
            widget.bind("<Button-1>", on_click)

        wrapper.accent = accent
        wrapper.label = label
        return wrapper

    def select_nav(self, key):
        """Switch the visible content page and restyle the sidebar so only
        the active section shows the accent bar / bold text / tinted bg."""
        self.active_nav = key
        for k, item in self.nav_items.items():
            selected = (k == key)
            bg = self.COLOR_SIDEBAR_ACTIVE if selected else self.COLOR_SIDEBAR
            item.configure(bg=bg)
            item.accent.configure(bg=self.COLOR_ACCENT if selected else self.COLOR_SIDEBAR)
            item.label.configure(
                bg=bg,
                fg=self.COLOR_PRIMARY if selected else self.COLOR_MUTED,
                font=("Segoe UI", 11, "bold" if selected else "normal"),
            )
        self.nav_pages[key].tkraise()

    def _build_content_area(self, parent):
        """Stacked pages sharing one grid cell; select_nav() raises the
        active one. Outstanding Premium hosts the existing full workflow;
        the other two are placeholders reserved for future modules."""
        container = tk.Frame(parent, bg=self.COLOR_BG)
        container.pack(side="left", fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.nav_pages = {}

        outstanding_page = ttk.Frame(container)
        outstanding_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["outstanding_premium"] = outstanding_page
        self._build_outstanding_premium_page(outstanding_page)

        for key, title in [("loss_ratio", "Loss Ratio"), ("renewal", "Renewal")]:
            page = self._build_placeholder_page(container, title)
            page.grid(row=0, column=0, sticky="nsew")
            self.nav_pages[key] = page

        self.select_nav("outstanding_premium")

    def _build_placeholder_page(self, parent, title):
        """Clean, on-brand empty state for modules not built yet."""
        page = tk.Frame(parent, bg=self.COLOR_BG)
        center = tk.Frame(page, bg=self.COLOR_BG)
        center.place(relx=0.5, rely=0.42, anchor="center")

        icon = tk.Canvas(center, width=64, height=64, bg=self.COLOR_BG, highlightthickness=0)
        icon.pack(pady=(0, 16))
        icon.create_oval(2, 2, 62, 62, outline=self.COLOR_BORDER, width=2, fill=self.COLOR_PANEL)
        icon.create_text(32, 34, text="+", fill=self.COLOR_ACCENT, font=("Segoe UI", 26, "bold"))

        tk.Label(center, text=title, font=("Segoe UI Semibold", 18, "bold"),
                 bg=self.COLOR_BG, fg=self.COLOR_PRIMARY).pack(pady=(0, 6))
        tk.Label(center, text="This module is coming soon.", font=("Segoe UI", 11),
                 bg=self.COLOR_BG, fg=self.COLOR_MUTED).pack()
        return page

    def _build_outstanding_premium_page(self, parent):
        """The existing Outstanding Premium workflow (Upload & Filter,
        Qualifying Clients, Review & Send Emails) - unchanged logic, now
        hosted as one sidebar section instead of the whole app."""
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self.tab_setup = ttk.Frame(notebook)
        self.tab_results = ttk.Frame(notebook)
        self.tab_emails = ttk.Frame(notebook)
        notebook.add(self.tab_setup, text="1. Upload & Filter")
        notebook.add(self.tab_results, text="2. Qualifying Clients")
        notebook.add(self.tab_emails, text="3. Review & Send Emails")

        self._build_setup_tab()
        self._build_results_tab()
        self._build_emails_tab()

    def _build_setup_tab(self):
        f = self.tab_setup
        f.configure(style="TFrame")

        # Scrollable container: on smaller screens/windows the stacked sections
        # below can exceed the visible height, so wrap them in a canvas with a
        # scrollbar rather than risk any control being pushed off-screen.
        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, background=self.COLOR_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        canvas_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=canvas_frame, anchor="nw")

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas_frame.bind("<Configure>", _on_frame_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        pad = {"padx": 10, "pady": 8}

        # ---- Section: Data Source ----
        source_box = ttk.Labelframe(canvas_frame, text="  Data Source  ", padding=14)
        source_box.pack(fill="x", padx=12, pady=(12, 8))
        source_box.columnconfigure(1, weight=1)

        ttk.Label(source_box, text="Balance Data File (.xlsx)", style="Panel.TLabel").grid(row=0, column=0, sticky="w", **pad)
        self.file_path_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.file_path_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_data_file).grid(row=0, column=2, **pad)

        ttk.Label(source_box, text="Sheet name (raw data)", style="Panel.TLabel").grid(row=1, column=0, sticky="w", **pad)
        self.sheet_name_var = tk.StringVar(value="Sheet1")
        ttk.Entry(source_box, textvariable=self.sheet_name_var, width=20).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(source_box, text="Use which date for 'Year' breakdown", style="Panel.TLabel").grid(row=2, column=0, sticky="w", **pad)
        self.date_col_var = tk.StringVar(value="Issue Date")
        date_combo = ttk.Combobox(source_box, textvariable=self.date_col_var,
                                   values=["Issue Date", "Comm.  Date"], state="readonly", width=20)
        date_combo.grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(source_box, text="Email Lookup File (.xlsx/.csv, optional)", style="Panel.TLabel").grid(row=3, column=0, sticky="w", **pad)
        self.email_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.email_file_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_email_file).grid(row=3, column=2, **pad)

        ttk.Button(source_box, text="Load & Validate Data", style="Accent.TButton", width=24,
                   command=self.load_data).grid(row=4, column=1, sticky="w", pady=(10, 4))

        # ---- Section: Threshold & Filters ----
        filter_box = ttk.Labelframe(canvas_frame, text="  Threshold & Filters  ", padding=14)
        filter_box.pack(fill="x", padx=12, pady=8)
        filter_box.columnconfigure(0, weight=1)
        filter_box.columnconfigure(1, weight=1)

        threshold_row = ttk.Frame(filter_box, style="Panel.TFrame")
        threshold_row.grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 10))
        ttk.Label(threshold_row, text="Outstanding Balance Threshold (PKR)", style="Panel.TLabel").pack(side="left", padx=(0, 10))
        self.threshold_var = tk.StringVar(value="2000000")
        ttk.Entry(threshold_row, textvariable=self.threshold_var, width=20, font=("Segoe UI", 11, "bold")).pack(side="left")

        # side-by-side wide filter lists
        dept_col = ttk.Frame(filter_box, style="Panel.TFrame")
        dept_col.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        ttk.Label(dept_col, text="Department filter", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(dept_col, text="Ctrl+click to select multiple · none = all", style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
        dept_frame = tk.Frame(dept_col, highlightbackground="#D6DCE1", highlightthickness=1)
        dept_frame.pack(fill="both", expand=True)
        dept_scroll = ttk.Scrollbar(dept_frame, orient="vertical")
        self.dept_listbox = tk.Listbox(dept_frame, selectmode="extended", height=11, width=65,
                                        exportselection=False, font=("Segoe UI", 10),
                                        activestyle="none", relief="flat", borderwidth=0,
                                        selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                        highlightthickness=0,
                                        yscrollcommand=dept_scroll.set)
        dept_scroll.config(command=self.dept_listbox.yview)
        self.dept_listbox.pack(side="left", fill="both", expand=True)
        dept_scroll.pack(side="right", fill="y")

        agent_col = ttk.Frame(filter_box, style="Panel.TFrame")
        agent_col.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)
        ttk.Label(agent_col, text="Agency/Agent filter", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(agent_col, text="Ctrl+click to select multiple · none = all", style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
        agent_frame = tk.Frame(agent_col, highlightbackground="#D6DCE1", highlightthickness=1)
        agent_frame.pack(fill="both", expand=True)
        agent_scroll = ttk.Scrollbar(agent_frame, orient="vertical")
        self.agent_listbox = tk.Listbox(agent_frame, selectmode="extended", height=11, width=65,
                                         exportselection=False, font=("Segoe UI", 10),
                                         activestyle="none", relief="flat", borderwidth=0,
                                         selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                         highlightthickness=0,
                                         yscrollcommand=agent_scroll.set)
        agent_scroll.config(command=self.agent_listbox.yview)
        self.agent_listbox.pack(side="left", fill="both", expand=True)
        agent_scroll.pack(side="right", fill="y")

        company_row = ttk.Frame(filter_box, style="Panel.TFrame")
        company_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 4))
        ttk.Label(company_row, text="Company filter (optional, comma-separated exact names)",
                  style="Panel.TLabel").pack(anchor="w")
        self.company_filter_var = tk.StringVar()
        ttk.Entry(company_row, textvariable=self.company_filter_var, width=90).pack(fill="x", pady=(4, 0))

        ttk.Button(filter_box, text="Analyze / Apply Filters", style="Accent.TButton", width=24,
                   command=self.analyze).grid(row=3, column=0, sticky="w", padx=6, pady=(12, 4))

        # ---- Section: Log ----
        log_box = ttk.Labelframe(canvas_frame, text="  Activity Log  ", padding=10)
        log_box.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.setup_log = scrolledtext.ScrolledText(log_box, height=10, font=("Consolas", 9),
                                                    bg="#FBFCFD", relief="flat", borderwidth=0)
        self.setup_log.pack(fill="both", expand=True)

    def _build_results_tab(self):
        f = self.tab_results

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Generate Excel Reports", style="Accent.TButton", width=24,
                   command=self.generate_reports).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Open Output Folder", command=self.open_output_folder, width=20).pack(side="left")
        self.results_summary_var = tk.StringVar(value="Run Analyze on tab 1 to see qualifying clients here.")
        ttk.Label(top, textvariable=self.results_summary_var, style="Muted.TLabel").pack(side="right")

        tree_frame = ttk.Frame(f)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("client", "total", "agent")
        self.results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.results_tree.heading("client", text="Client / Participant")
        self.results_tree.heading("total", text="Grand Total (O/s)")
        self.results_tree.heading("agent", text="Agent(s) / Agency")
        self.results_tree.column("client", width=420)
        self.results_tree.column("total", width=160, anchor="e")
        self.results_tree.column("agent", width=300)
        self.results_tree.tag_configure("oddrow", background="#F4F7FA")
        self.results_tree.tag_configure("evenrow", background="white")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb.set)
        self.results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_emails_tab(self):
        f = self.tab_emails

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Draft Emails", style="Accent.TButton", command=self.draft_emails, width=18).pack(side="left", padx=(0, 12))
        self.dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Dry run (preview only, do not actually send)", variable=self.dry_run_var).pack(side="left", padx=(0, 12))
        ttk.Button(top, text="Send All", style="Send.TButton", command=self.send_all_emails, width=14).pack(side="left")

        sig_frame = ttk.Labelframe(f, text="  Your Signature (appears after 'Regards,' in every email)  ", padding=10)
        sig_frame.pack(fill="x", padx=12, pady=(0, 8))
        sig_frame.columnconfigure(0, weight=1)
        self.signature_text = tk.Text(sig_frame, height=3, font=("Segoe UI", 10), relief="solid", borderwidth=1, wrap="word")
        self.signature_text.grid(row=0, column=0, sticky="ew")
        self.signature_text.insert("1.0", eg.load_signature())
        sig_btn = tk.Button(
            sig_frame, text="Save Signature", command=self.save_signature,
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=10, pady=4, cursor="hand2", borderwidth=0,
        )
        sig_btn.grid(row=0, column=1, sticky="n", padx=(10, 0))

        body = ttk.Frame(f)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        # Give the list a fixed proportion and the preview panel the rest,
        # using grid so the split stays consistent as the window resizes.
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(left, text="Qualifying Clients", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        columns = ("client", "email", "status")
        self.email_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=26)
        self.email_tree.heading("client", text="Client")
        self.email_tree.heading("email", text="Email")
        self.email_tree.heading("status", text="Status")
        self.email_tree.column("client", width=250)
        self.email_tree.column("email", width=190)
        self.email_tree.column("status", width=170)
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.email_tree.yview)
        self.email_tree.configure(yscrollcommand=tree_scroll.set)
        self.email_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.email_tree.bind("<<TreeviewSelect>>", self.on_select_email)

        right = ttk.Labelframe(body, text="  Email Preview  ", padding=14)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)  # body text area gets the extra vertical space

        ttk.Label(right, text="To (recipient email)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.subject_var, font=("Segoe UI", 10))\
            .grid(row=2, column=0, sticky="ew")

        # Body label + text area share the remaining vertical space
        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=3, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1)
        self.body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        # Plain tk.Button (not ttk) for this action: on some Windows ttk
        # themes, ttk.Button silently ignores the 'width' option for bold
        # fonts and clips the label. A classic tk.Button always sizes
        # itself exactly to its own text, so the full label is guaranteed
        # to render.
        btn_row = ttk.Frame(right)
        btn_row.grid(row=4, column=0, sticky="w")

        save_btn = tk.Button(
            btn_row, text="Save Edits to Selected Email", command=self.save_email_edit,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=14, pady=8, cursor="hand2", borderwidth=0,
        )
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = tk.Button(
            btn_row, text="Send This Email", command=self.send_individual_email,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_SUCCESS, fg="white",
            activebackground="#1B5E20", activeforeground="white",
            relief="flat", padx=14, pady=8, cursor="hand2", borderwidth=0,
        )
        send_one_btn.pack(side="left")

    # ---------- Actions ----------
    def browse_data_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path_var.set(path)

    def browse_email_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.csv")])
        if path:
            self.email_file_var.set(path)

    def log(self, widget, msg):
        widget.insert("end", msg + "\n")
        widget.see("end")

    def load_data(self):
        path = self.file_path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid balance data file.")
            return
        try:
            self.raw_df = dp.load_raw_data(path, sheet_name=self.sheet_name_var.get())
            self.clean_df, self.issues_df, dropped_columns = dp.validate_data(
                self.raw_df, date_col_for_year=self.date_col_var.get()
            )
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))
            return

        self.dept_listbox.delete(0, "end")
        for i, dept in enumerate(sorted(self.clean_df["Department"].dropna().unique())):
            self.dept_listbox.insert("end", dept)
            self.dept_listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

        self.agent_listbox.delete(0, "end")
        if "Agent Name" in self.clean_df.columns:
            for i, agent in enumerate(sorted(self.clean_df["Agent Name"].dropna().unique())):
                self.agent_listbox.insert("end", agent)
                self.agent_listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

        self.log(self.setup_log, f"Loaded {len(self.raw_df)} rows, {len(self.clean_df)} valid after cleaning.")
        if dropped_columns:
            self.log(self.setup_log, f"Removed {len(dropped_columns)} completely empty column(s) from the data: {', '.join(dropped_columns)}")
        if len(self.issues_df) > 0:
            self.log(self.setup_log, f"{len(self.issues_df)} rows flagged with issues (missing client/department, non-numeric balance):")
            for _, r in self.issues_df.head(20).iterrows():
                self.log(self.setup_log, f"   Row {r['row']}: {r['issue']} (value: {r['value']})")
            if len(self.issues_df) > 20:
                self.log(self.setup_log, f"   ...and {len(self.issues_df) - 20} more.")
        else:
            self.log(self.setup_log, "No data issues found.")

        email_path = self.email_file_var.get().strip()
        self.email_lookup = {}
        self.poc_lookup = {}
        if email_path and os.path.exists(email_path):
            try:
                if email_path.lower().endswith(".csv"):
                    edf = pd.read_csv(email_path)
                else:
                    edf = pd.read_excel(email_path)
                name_col = next((c for c in edf.columns if "client" in c.lower() or "participant" in c.lower() or "company" in c.lower()), edf.columns[0])
                email_col = next((c for c in edf.columns if "email" in c.lower()), None)
                if email_col is None:
                    raise ValueError("Could not find an 'Email' column in the lookup file.")
                self.email_lookup = dict(zip(edf[name_col].astype(str).str.strip(), edf[email_col].astype(str).str.strip()))
                self.log(self.setup_log, f"Loaded {len(self.email_lookup)} client-email mappings from lookup file.")

                # Optional POC / contact-person column: if present, the email
                # greeting addresses this person by name instead of the company.
                poc_col = next((c for c in edf.columns
                                 if c != email_col and c != name_col
                                 and any(k in c.lower() for k in ["poc", "contact", "attn", "attention", "recipient name", "contact person"])),
                                None)
                if poc_col:
                    poc_lookup = {}
                    for client_raw, poc_raw in zip(edf[name_col], edf[poc_col]):
                        if pd.isna(client_raw) or pd.isna(poc_raw):
                            continue
                        client_name = str(client_raw).strip()
                        poc_name = str(poc_raw).strip()
                        if client_name and poc_name:
                            poc_lookup[client_name] = poc_name
                    self.poc_lookup = poc_lookup
                    self.log(self.setup_log, f"Found POC column '{poc_col}' - {len(self.poc_lookup)} client(s) will be greeted by contact name.")
            except Exception as e:
                self.log(self.setup_log, f"Warning: could not load email lookup file: {e}")

        self.status_var.set("Data loaded. Set threshold/filters and click Analyze.")

    def analyze(self):
        if self.clean_df is None:
            messagebox.showerror("Error", "Load data first.")
            return
        try:
            threshold = float(self.threshold_var.get())
        except ValueError:
            messagebox.showerror("Error", "Threshold must be a number.")
            return

        selected_depts = [self.dept_listbox.get(i) for i in self.dept_listbox.curselection()]
        selected_agents = [self.agent_listbox.get(i) for i in self.agent_listbox.curselection()]
        company_text = self.company_filter_var.get().strip()
        company_filter = [c.strip() for c in company_text.split(",") if c.strip()] if company_text else None

        self.filtered_df, self.grand_totals = dp.build_pivot(
            self.clean_df, company_filter=company_filter,
            department_filter=selected_depts or None,
            agent_filter=selected_agents or None,
        )
        self.qualifying = dp.qualifying_clients(self.filtered_df, self.grand_totals, threshold)

        self.results_tree.delete(*self.results_tree.get_children())
        for i, (client, total) in enumerate(self.qualifying.items()):
            agents = ", ".join(dp.client_agents(self.filtered_df, client))
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.results_tree.insert("", "end", values=(client, f"{total:,.0f}", agents), tags=(tag,))
        self.results_summary_var.set(f"{len(self.qualifying)} client(s) meet/exceed threshold {threshold:,.0f}")

        self.log(self.setup_log, f"\nAnalysis complete: {len(self.qualifying)} clients meet/exceed threshold {threshold:,.0f}.")
        self.status_var.set(f"{len(self.qualifying)} qualifying clients. See tab 2.")

    def generate_reports(self):
        if self.qualifying is None or len(self.qualifying) == 0:
            messagebox.showwarning("Nothing to generate", "Run Analyze first (tab 1).")
            return
        os.makedirs(self.output_dir, exist_ok=True)

        use_native = pivot_native.EXCEL_AVAILABLE
        excel_app = None
        generated = []
        native_failures = []

        self.status_var.set("Generating reports..." + (" (building native PivotTables in Excel)" if use_native else ""))
        self.update_idletasks()

        try:
            if use_native:
                try:
                    excel_app = pivot_native.start_excel()
                except Exception as e:
                    use_native = False
                    self.log(self.setup_log, f"Could not start Excel for native PivotTables ({e}) - using static formatted tables instead.")

            for client in self.qualifying.index:
                raw_rows = dp.get_client_raw_rows(self.filtered_df, client)
                fp = None
                if use_native:
                    try:
                        fp = rg.write_raw_only_workbook(client, raw_rows, self.output_dir)
                        pivot_native.add_native_pivot_table(fp, client_name=client, excel_app=excel_app)
                    except Exception as e:
                        native_failures.append(client)
                        pivot_df = dp.client_pivot_table(self.filtered_df, client)
                        fp = rg.write_client_report(client, pivot_df, self.output_dir, raw_df=raw_rows)
                else:
                    pivot_df = dp.client_pivot_table(self.filtered_df, client)
                    fp = rg.write_client_report(client, pivot_df, self.output_dir, raw_df=raw_rows)
                generated.append(fp)
        finally:
            pivot_native.stop_excel(excel_app)

        msg = f"Generated {len(generated)} report(s) in:\n{self.output_dir}"
        if not pivot_native.EXCEL_AVAILABLE:
            msg += ("\n\nMicrosoft Excel/pywin32 wasn't detected on this machine, so reports use a "
                    "static formatted table instead of an interactive native PivotTable. Install "
                    "Excel (and pywin32) for real, drillable PivotTables.")
        elif native_failures:
            msg += (f"\n\n{len(native_failures)} report(s) fell back to a static table after an Excel "
                    f"error: {', '.join(native_failures[:5])}" + (", ..." if len(native_failures) > 5 else ""))
        else:
            msg += "\n\nEach report contains a genuine, interactive Excel PivotTable (with the field list, filters, and drill-down) built from that client's Raw Data sheet."
        messagebox.showinfo("Done", msg)
        self.status_var.set(f"Generated {len(generated)} report(s).")

    def open_output_folder(self):
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir) if hasattr(os, "startfile") else os.system(f'open "{self.output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate reports first.")

    def draft_emails(self):
        if self.qualifying is None or len(self.qualifying) == 0:
            messagebox.showwarning("Nothing to draft", "Run Analyze first (tab 1).")
            return
        template = eg.load_templates().get("Standard Reminder", eg.DEFAULT_TEMPLATE)
        signature = self.signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)  # persist automatically so it's remembered next time
        from datetime import date
        self.drafts = eg.build_email_drafts(
            self.qualifying.index, self.grand_totals, self.email_lookup,
            poc_lookup=self.poc_lookup, template=template, report_date=date.today().isoformat(),
            signature=signature,
        )
        self.email_tree.delete(*self.email_tree.get_children())
        for d in self.drafts:
            report_path = os.path.join(self.output_dir, f"{rg.safe_filename(d['client_name'])}_OutstandingBalance_Report.xlsx")
            has_report = os.path.exists(report_path)
            status = d["status"] if d["status"] != "Ready" else ("Ready" if has_report else "Ready (no report yet)")
            self.email_tree.insert("", "end", values=(d["client_name"], d["email"] or "(missing)", status))
        self.status_var.set(f"Drafted {len(self.drafts)} emails.")

    def save_signature(self):
        signature = self.signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)
        self.status_var.set("Signature saved. It will be used the next time you draft emails (or click Draft Emails again to refresh existing drafts).")

    def on_select_email(self, event):
        sel = self.email_tree.selection()
        if not sel:
            return
        idx = self.email_tree.index(sel[0])
        d = self.drafts[idx]
        self.email_to_var.set(d.get("email") or "")
        self.subject_var.set(d["subject"])
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", d["body"])

    def save_email_edit(self):
        sel = self.email_tree.selection()
        if not sel:
            messagebox.showinfo("No client selected", "Select a client from the list on the left first.")
            return
        idx = self.email_tree.index(sel[0])
        new_email = self.email_to_var.get().strip()
        self.drafts[idx]["email"] = new_email or None
        self.drafts[idx]["subject"] = self.subject_var.get()
        self.drafts[idx]["body"] = self.body_text.get("1.0", "end").strip()

        # Reflect the change immediately in the client list (email + status columns)
        report_path = os.path.join(self.output_dir, f"{rg.safe_filename(self.drafts[idx]['client_name'])}_OutstandingBalance_Report.xlsx")
        has_report = os.path.exists(report_path)
        if new_email:
            status = "Ready" if has_report else "Ready (no report yet)"
        else:
            status = "Missing email - needs manual entry"
        row_id = self.email_tree.get_children()[idx]
        self.email_tree.set(row_id, "email", new_email or "(missing)")
        self.email_tree.set(row_id, "status", status)

        self.status_var.set(f"Updated draft for {self.drafts[idx]['client_name']}.")

    def send_all_emails(self):
        if not self.drafts:
            messagebox.showwarning("Nothing to send", "Draft emails first.")
            return

        missing = [d["client_name"] for d in self.drafts
                   if not os.path.exists(os.path.join(
                       self.output_dir, f"{rg.safe_filename(d['client_name'])}_OutstandingBalance_Report.xlsx"))]
        if missing:
            proceed = messagebox.askyesno(
                "Reports not generated",
                f"{len(missing)} of {len(self.drafts)} client(s) have no generated Excel report yet, "
                f"so those emails would go out with no attachment:\n\n"
                + "\n".join(missing[:10]) + (f"\n...and {len(missing) - 10} more" if len(missing) > 10 else "")
                + "\n\nGo to tab 2 and click 'Generate Excel Reports' first, "
                  "or continue anyway without attachments?\n\nContinue anyway?"
            )
            if not proceed:
                return

        dry_run = self.dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send {len(self.drafts)} emails via Outlook now? This cannot be undone."):
                return

        def worker():
            results = osend.send_bulk(self.drafts, attachments_dir=self.output_dir, dry_run=dry_run)
            sent = sum(1 for r in results if r["sent"])
            failed = len(results) - sent
            run_log.log_run(
                threshold=float(self.threshold_var.get()),
                company_filter=self.company_filter_var.get(),
                department_filter=[self.dept_listbox.get(i) for i in self.dept_listbox.curselection()],
                num_qualifying=len(self.qualifying),
                num_sent=sent, num_failed=failed, dry_run=dry_run,
            )
            for i, r in enumerate(results):
                self.email_tree.set(self.email_tree.get_children()[i], "status", "Sent" if r["sent"] else f"Failed: {r['send_message']}")
            messagebox.showinfo("Done", f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")
            self.status_var.set(f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")

        threading.Thread(target=worker, daemon=True).start()

    def send_individual_email(self):
        sel = self.email_tree.selection()
        if not sel:
            messagebox.showinfo("No client selected", "Click a client in the list on the left first.")
            return
        idx = self.email_tree.index(sel[0])

        # Save whatever is currently in the To/Subject/Body fields first,
        # so the send always reflects what's on screen.
        self.save_email_edit()
        d = self.drafts[idx]

        if not d.get("email"):
            messagebox.showerror("Missing email", "Enter a recipient email address in the 'To' field before sending.")
            return

        report_path = os.path.join(self.output_dir, f"{rg.safe_filename(d['client_name'])}_OutstandingBalance_Report.xlsx")
        attachment_path = report_path if os.path.exists(report_path) else None
        if attachment_path is None:
            if not messagebox.askyesno(
                "No report generated",
                f"No Excel report has been generated yet for {d['client_name']}. "
                f"Send this email without an attachment?\n\n"
                f"(Go to tab 2 and click 'Generate Excel Reports' first if you'd rather attach one.)"
            ):
                return

        dry_run = self.dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send this email to {d['email']} ({d['client_name']}) now?"):
                return

        row_id = self.email_tree.get_children()[idx]

        def worker():
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run)
            self.email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
            prefix = "[DRY RUN] " if dry_run else ""
            messagebox.showinfo("Done", f"{prefix}{message}")
            self.status_var.set(f"{prefix}{message}")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = BalanceApp()
    app.mainloop()
