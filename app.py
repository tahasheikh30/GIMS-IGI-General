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
import loss_ratio_processor as lrp
import loss_ratio_generator as lrg
import receivable_automation as ra
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

        self.loss_group_files = []  # list of filepaths selected in Group mode

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
        """Fixed top header showing only the GIMS logo + wordmark - stays
        visible across every sidebar section and never breaks on resize
        (fixed height, left-anchored content, no elements that depend on
        window width)."""
        header = tk.Frame(self, bg=self.COLOR_PRIMARY, height=64)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        brand = tk.Frame(header, bg=self.COLOR_PRIMARY)
        brand.place(relx=0, rely=0.5, x=20, anchor="w")  # vertically centered, fixed left inset

        # Simple drawn logo mark (no external image asset needed) - a
        # circular accent-colored badge with the "G" initial.
        mark = tk.Canvas(brand, width=40, height=40, bg=self.COLOR_PRIMARY, highlightthickness=0)
        mark.pack(side="left", padx=(0, 12))
        mark.create_oval(0, 0, 40, 40, fill=self.COLOR_ACCENT, outline="")
        mark.create_text(20, 21, text="G", fill="white", font=("Segoe UI", 17, "bold"))

        tk.Label(brand, text="GIMS", font=("Segoe UI Semibold", 20, "bold"),
                 bg=self.COLOR_PRIMARY, fg="white").pack(side="left")

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
            ("receivable", "Premium & Claims Receivable"),
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

        loss_ratio_page = ttk.Frame(container)
        loss_ratio_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["loss_ratio"] = loss_ratio_page
        self._build_loss_ratio_page(loss_ratio_page)

        receivable_page = ttk.Frame(container)
        receivable_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["receivable"] = receivable_page
        self._build_receivable_page(receivable_page)

        renewal_page = self._build_placeholder_page(container, "Renewal")
        renewal_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["renewal"] = renewal_page

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

    # ---------- Loss Ratio module ----------
    def _build_loss_ratio_page(self, parent):
        # Scrollable container - same safety pattern as the Outstanding
        # Premium setup tab, so nothing can ever be pushed off-screen on a
        # smaller window.
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, background=self.COLOR_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        body = ttk.Frame(canvas)
        body_window = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        body.bind("<Configure>", _on_body_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(body_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        PAD_X = 20

        # --- Page header ---
        header = ttk.Frame(body)
        header.pack(fill="x", padx=PAD_X, pady=(20, 6))
        ttk.Label(header, text="Loss Ratio", font=("Segoe UI Semibold", 17, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        ttk.Label(header, text="Calculate and summarize Loss Ratio for one company or a group, from Excel or PDF source files.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

        formula_badge = tk.Frame(header, bg=self.COLOR_SIDEBAR_ACTIVE)
        formula_badge.pack(anchor="w")
        tk.Label(formula_badge, text="Loss Ratio %  =  (Incurred Loss + Commission Accrual)  \u00f7  Gross Premium Written  \u00d7  100",
                 bg=self.COLOR_SIDEBAR_ACTIVE, fg=self.COLOR_PRIMARY, font=("Consolas", 10, "bold"),
                 padx=14, pady=8).pack()

        # --- Section 1: Mode ---
        s1 = self._loss_section_card(body, "1", "Mode & Report Name", PAD_X)
        self.loss_mode_var = tk.StringVar(value="individual")
        mode_row = ttk.Frame(s1)
        mode_row.pack(fill="x", pady=(0, 12))
        ttk.Radiobutton(mode_row, text="Individual Company  (1 file)", variable=self.loss_mode_var,
                         value="individual", command=self._toggle_loss_mode).pack(side="left", padx=(0, 30))
        ttk.Radiobutton(mode_row, text="Group of Companies  (multiple files)", variable=self.loss_mode_var,
                         value="group", command=self._toggle_loss_mode).pack(side="left")

        name_row = ttk.Frame(s1)
        name_row.pack(fill="x")
        ttk.Label(name_row, text="Group / Report Name", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Label(name_row, text="Shown in the green banner at the top of the output (e.g. \"POPULAR GROUP\"). Optional.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 4))
        self.loss_group_name_var = tk.StringVar(value="")
        ttk.Entry(name_row, textvariable=self.loss_group_name_var, width=40, font=("Segoe UI", 10)).pack(anchor="w")

        # --- Section 2: File selection ---
        s2 = self._loss_section_card(body, "2", "Input Files", PAD_X)
        ttk.Label(s2, text="Accepts .xlsx and .pdf. A single multi-client PDF (covering a client range) can populate an entire group automatically.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 10))

        file_container = ttk.Frame(s2)
        file_container.pack(fill="x")
        file_container.columnconfigure(0, weight=1)

        self.loss_individual_frame = ttk.Frame(file_container)
        self.loss_individual_frame.grid(row=0, column=0, sticky="ew")
        self.loss_individual_file_var = tk.StringVar()
        ttk.Entry(self.loss_individual_frame, textvariable=self.loss_individual_file_var, font=("Segoe UI", 10))\
            .pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(self.loss_individual_frame, text="Browse...", command=self._browse_loss_individual_file, width=14)\
            .pack(side="left")

        self.loss_group_frame = ttk.Frame(file_container)
        self.loss_group_frame.grid(row=0, column=0, sticky="ew")
        group_top = ttk.Frame(self.loss_group_frame)
        group_top.pack(fill="x", pady=(0, 8))
        ttk.Button(group_top, text="Add File(s)...", command=self._add_loss_group_files, width=16).pack(side="left", padx=(0, 8))
        ttk.Button(group_top, text="Remove Selected", command=self._remove_loss_group_file, width=16).pack(side="left")

        list_frame = tk.Frame(self.loss_group_frame, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_frame.pack(fill="x")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.loss_group_listbox = tk.Listbox(list_frame, height=6, font=("Segoe UI", 10), selectmode="extended",
                                              activestyle="none", relief="flat", borderwidth=0,
                                              selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                              highlightthickness=0, yscrollcommand=list_scroll.set)
        list_scroll.configure(command=self.loss_group_listbox.yview)
        self.loss_group_listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        # --- Section 3: Threshold + Calculate ---
        s3 = self._loss_section_card(body, "3", "Threshold & Generate", PAD_X)
        threshold_row = ttk.Frame(s3)
        threshold_row.pack(fill="x", pady=(0, 14))
        ttk.Label(threshold_row, text="Red highlight threshold:", font=("Segoe UI", 10)).pack(side="left", padx=(0, 8))
        self.loss_threshold_var = tk.StringVar(value="100")
        ttk.Entry(threshold_row, textvariable=self.loss_threshold_var, width=6, font=("Segoe UI", 10)).pack(side="left")
        ttk.Label(threshold_row, text="%", font=("Segoe UI", 10)).pack(side="left", padx=(4, 16))
        ttk.Label(threshold_row, text="Rows \u2265 70% are always flagged amber, regardless of this value.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(side="left")

        calc_btn = tk.Button(
            s3, text="Calculate Loss Ratio", command=self._calculate_loss_ratio,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=18, pady=9, cursor="hand2", borderwidth=0,
        )
        calc_btn.pack(anchor="w")

        # --- Section 4: Log ---
        s4 = self._loss_section_card(body, "4", "Processing Log", PAD_X, expand_bottom=True)
        self.loss_log = scrolledtext.ScrolledText(s4, height=10, font=("Consolas", 9),
                                                    relief="solid", borderwidth=1)
        self.loss_log.pack(fill="both", expand=True)

        self._toggle_loss_mode()  # show the correct file-selection frame initially

    def _loss_section_card(self, parent, number, title, pad_x, expand_bottom=False):
        """A numbered section card matching the app's '1. Upload & Filter'
        style numbering convention, for visual consistency across pages."""
        wrapper = tk.Frame(parent, bg=self.COLOR_PANEL, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        wrapper.pack(fill="both" if expand_bottom else "x", expand=expand_bottom,
                     padx=pad_x, pady=8)

        title_row = tk.Frame(wrapper, bg=self.COLOR_PANEL)
        title_row.pack(fill="x", padx=16, pady=(14, 10))

        badge = tk.Canvas(title_row, width=22, height=22, bg=self.COLOR_PANEL, highlightthickness=0)
        badge.pack(side="left", padx=(0, 8))
        badge.create_oval(1, 1, 21, 21, fill=self.COLOR_ACCENT, outline="")
        badge.create_text(11, 11, text=number, fill="white", font=("Segoe UI", 9, "bold"))

        tk.Label(title_row, text=title, bg=self.COLOR_PANEL, fg=self.COLOR_PRIMARY,
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        content = tk.Frame(wrapper, bg=self.COLOR_PANEL)
        content.pack(fill="both" if expand_bottom else "x", expand=expand_bottom, padx=16, pady=(0, 16))
        return content

    def _toggle_loss_mode(self):
        if self.loss_mode_var.get() == "individual":
            self.loss_individual_frame.tkraise()
        else:
            self.loss_group_frame.tkraise()

    def _browse_loss_individual_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel or PDF files", "*.xlsx *.xls *.pdf"),
            ("Excel files", "*.xlsx *.xls"),
            ("PDF files", "*.pdf"),
        ])
        if path:
            self.loss_individual_file_var.set(path)

    def _add_loss_group_files(self):
        paths = filedialog.askopenfilenames(filetypes=[
            ("Excel or PDF files", "*.xlsx *.xls *.pdf"),
            ("Excel files", "*.xlsx *.xls"),
            ("PDF files", "*.pdf"),
        ])
        added = 0
        for p in paths:
            if p not in self.loss_group_files:
                self.loss_group_files.append(p)
                self.loss_group_listbox.insert("end", os.path.basename(p))
                added += 1
        if added:
            self.status_var.set(f"Added {added} file(s) to the Loss Ratio group ({len(self.loss_group_files)} total).")

    def _remove_loss_group_file(self):
        selected = list(self.loss_group_listbox.curselection())
        if not selected:
            messagebox.showinfo("Nothing selected", "Select one or more files in the list to remove.")
            return
        for idx in reversed(selected):
            self.loss_group_listbox.delete(idx)
            del self.loss_group_files[idx]
        self.status_var.set(f"{len(self.loss_group_files)} file(s) remaining in the group.")

    def _log_loss(self, msg):
        self.loss_log.insert("end", msg + "\n")
        self.loss_log.see("end")

    def _calculate_loss_ratio(self):
        mode = self.loss_mode_var.get()

        if mode == "individual":
            path = self.loss_individual_file_var.get().strip()
            if not path:
                messagebox.showerror("No file selected", "Browse to an Excel file first.")
                return
            if not os.path.exists(path):
                messagebox.showerror("File not found", f"Could not find:\n{path}")
                return
            filepaths = [path]
        else:
            if not self.loss_group_files:
                messagebox.showerror("No files selected", "Add at least one Excel file to the group first.")
                return
            filepaths = list(self.loss_group_files)

        try:
            threshold = float(self.loss_threshold_var.get())
        except ValueError:
            messagebox.showerror("Invalid threshold", "The red highlight threshold must be a number, e.g. 100.")
            return

        self.loss_log.delete("1.0", "end")
        self._log_loss(f"Processing {len(filepaths)} file(s) in {'Individual' if mode == 'individual' else 'Group'} mode...\n")

        rows, errors = lrp.process_files(filepaths)

        from collections import Counter
        rows_per_source = Counter(r.get("source_file", "") for r in rows)

        for r in rows:
            lr = f"{r['loss_ratio']*100:.2f}%" if r["loss_ratio"] is not None else "N/A"
            note = f"  [!] {r['warning']}" if r["warning"] else ""
            multi_note = "  (multi-client PDF)" if rows_per_source.get(r.get("source_file", ""), 1) > 1 else ""
            self._log_loss(f"  OK   [{r.get('source_file', '?'):<28}] {r['Company Name']:<35} GP={r['Gross Premium Written']:>14,.0f}  Claim={r['Incurred Loss']:>12,.0f}  Comm={r['Commission Accrual']:>12,.0f}  LR={lr}{note}{multi_note}")
        for e in errors:
            self._log_loss(f"  ERROR  {e.filename}: {e.message}")

        if not rows:
            self._log_loss("\nNo files processed successfully - nothing to write.")
            messagebox.showerror(
                "Processing failed",
                "None of the selected file(s) could be processed. See the Processing Log for details."
            )
            return

        if errors:
            proceed = messagebox.askyesno(
                "Some files had errors",
                f"{len(rows)} file(s) processed successfully, but {len(errors)} file(s) had errors "
                f"(see the Processing Log for details).\n\nContinue and generate the report for the "
                f"{len(rows)} successful file(s)?"
            )
            if not proceed:
                return

        group_name = self.loss_group_name_var.get().strip() or ("LOSS RATIO SUMMARY" if mode == "individual" else "GROUP SUMMARY")
        default_name = "LossRatio_Group_Summary.xlsx" if mode == "group" else f"LossRatio_{rg.safe_filename(rows[0]['Company Name'])}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Save Loss Ratio Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
        )
        if not output_path:
            self._log_loss("\nSave cancelled - report was not written.")
            return

        try:
            lrg.build_loss_ratio_workbook(rows, mode=mode, output_path=output_path,
                                           group_name=group_name, red_threshold_pct=threshold)
        except Exception as e:
            self._log_loss(f"\nERROR writing output file: {e}")
            messagebox.showerror("Could not save report", f"An error occurred while writing the Excel file:\n\n{e}")
            return

        self._log_loss(f"\nSaved: {output_path}")
        self.status_var.set(f"Loss Ratio report saved: {os.path.basename(output_path)}")
        messagebox.showinfo("Done", f"Loss Ratio report generated successfully:\n{output_path}")

    # ---------- Premium & Claims Receivable module ----------
    def _build_receivable_page(self, parent):
        """Premium & Claims Receivable automation - reads a master file
        (transaction sheet + address-book sheet), generates one two-sheet
        workbook per Client+Company group, and drafts/sends the statement
        emails, mirroring the Outstanding Premium / Loss Ratio pages'
        numbered-card layout."""
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, background=self.COLOR_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        body = ttk.Frame(canvas)
        body_window = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        body.bind("<Configure>", _on_body_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(body_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        PAD_X = 20

        header = ttk.Frame(body)
        header.pack(fill="x", padx=PAD_X, pady=(20, 6))
        ttk.Label(header, text="Premium & Claims Receivable", font=("Segoe UI Semibold", 17, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        ttk.Label(header, text="Generate a Premium Receivable / Claims Receivable statement per Client + "
                                "Company group, and draft/send the notification emails.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

        # --- Section 1: Source file ---
        s1 = self._loss_section_card(body, "1", "Source File", PAD_X)
        ttk.Label(s1, text="Master workbook (.xlsx) - both the transaction sheet and the address-book "
                            "sheet are read from this same file.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))

        file_row = ttk.Frame(s1)
        file_row.pack(fill="x", pady=(0, 10))
        self.receivable_file_var = tk.StringVar()
        ttk.Entry(file_row, textvariable=self.receivable_file_var, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_row, text="Browse...", command=self._browse_receivable_file, width=14).pack(side="left")

        sheet_row = ttk.Frame(s1)
        sheet_row.pack(fill="x")
        ttk.Label(sheet_row, text="Transaction sheet name", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(sheet_row, text="Address-book sheet name", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.receivable_sheet_var = tk.StringVar(value="Detail")
        self.receivable_addresses_sheet_var = tk.StringVar(value="Addresses Co Insurance")
        ttk.Entry(sheet_row, textvariable=self.receivable_sheet_var, width=28, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(4, 0))
        ttk.Entry(sheet_row, textvariable=self.receivable_addresses_sheet_var, width=28, font=("Segoe UI", 10)).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(4, 0))

        # --- Section 2: Options ---
        s2 = self._loss_section_card(body, "2", "Output & Options", PAD_X)

        out_row = ttk.Frame(s2)
        out_row.pack(fill="x", pady=(0, 10))
        ttk.Label(out_row, text="Output folder", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        out_entry_row = ttk.Frame(out_row)
        out_entry_row.pack(fill="x", pady=(4, 0))
        self.receivable_output_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "receivable_output"))
        ttk.Entry(out_entry_row, textvariable=self.receivable_output_dir_var, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_entry_row, text="Browse...", command=self._browse_receivable_output_dir, width=14).pack(side="left")

        self.receivable_dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(s2, text="Dry run (generate files and preview emails, do not actually send)",
                         variable=self.receivable_dry_run_var).pack(anchor="w", pady=(4, 2))

        self.receivable_require_ref_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(s2, text="Require Leader Ref Number on claims-only rows (unchecking still flags "
                                  "these rows, but keeps them in the Claims Receivable sheet)",
                         variable=self.receivable_require_ref_var).pack(anchor="w", pady=(2, 2))

        self.receivable_load_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(s2, text="Look up recipient emails from the address-book sheet",
                         variable=self.receivable_load_email_var).pack(anchor="w", pady=(2, 8))

        run_btn = tk.Button(
            s2, text="Run Receivable Automation", command=self._run_receivable_automation,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=18, pady=9, cursor="hand2", borderwidth=0,
        )
        run_btn.pack(anchor="w")

        # --- Section 3: Log ---
        s3 = self._loss_section_card(body, "3", "Processing Log", PAD_X, expand_bottom=True)
        self.receivable_log = scrolledtext.ScrolledText(s3, height=16, font=("Consolas", 9),
                                                          relief="solid", borderwidth=1)
        self.receivable_log.pack(fill="both", expand=True)

    def _browse_receivable_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.receivable_file_var.set(path)

    def _browse_receivable_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.receivable_output_dir_var.set(path)

    def _log_receivable(self, msg):
        self.receivable_log.insert("end", str(msg) + "\n")
        self.receivable_log.see("end")
        self.receivable_log.update_idletasks()

    def _run_receivable_automation(self):
        filepath = self.receivable_file_var.get().strip()
        if not filepath or not os.path.exists(filepath):
            messagebox.showerror("No file selected", "Browse to a master workbook first.")
            return

        output_dir = self.receivable_output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("No output folder", "Choose an output folder first.")
            return

        dry_run = self.receivable_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", "This will send real emails via Outlook for every "
                                                    "Client + Company group with a recipient email. "
                                                    "This cannot be undone. Continue?"):
                return

        self.receivable_log.delete("1.0", "end")
        self.status_var.set("Running Premium & Claims Receivable automation...")

        sheet_name = self.receivable_sheet_var.get().strip() or "Detail"
        addresses_sheet_name = self.receivable_addresses_sheet_var.get().strip() or "Addresses Co Insurance"
        require_ref = self.receivable_require_ref_var.get()
        load_email = self.receivable_load_email_var.get()

        def worker():
            try:
                summary = ra.run_pipeline(
                    master_filepath=filepath,
                    output_dir=output_dir,
                    sheet_name=sheet_name,
                    addresses_sheet_name=addresses_sheet_name,
                    load_email_lookup=load_email,
                    require_ref_for_claims_only_rows=require_ref,
                    dry_run=dry_run,
                    logger=self._log_receivable,
                )
            except Exception as e:
                self._log_receivable(f"\nUNEXPECTED ERROR: {e}")
                messagebox.showerror("Error", f"An unexpected error occurred:\n\n{e}")
                self.status_var.set("Receivable automation failed.")
                return

            num_sent = sum(1 for r in summary["send_results"] if r["sent"])
            num_failed = len(summary["send_results"]) - num_sent
            msg = (f"Groups processed: {summary['num_groups']}\n"
                   f"Workbooks generated: {len(summary['generated_files'])}\n"
                   f"Emails {'previewed (dry run)' if dry_run else 'sent'}: {num_sent}\n"
                   f"Emails failed: {num_failed}\n"
                   f"Missing email: {len(summary['missing_email'])}\n"
                   f"Other errors: {len(summary['errors'])}\n\n"
                   f"Output folder:\n{output_dir}")
            messagebox.showinfo("Done", msg)
            self.status_var.set(f"Receivable automation complete - {len(summary['generated_files'])} file(s), "
                                 f"{num_sent} email(s) {'previewed' if dry_run else 'sent'}.")

        threading.Thread(target=worker, daemon=True).start()

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

        # Wrap the button in its own cell-filling frame and pack(expand=True)
        # inside it, so the button is centered vertically against the
        # multi-line text box rather than floating at the top of the row.
        sig_btn_holder = tk.Frame(sig_frame, bg=self.COLOR_PANEL)
        sig_btn_holder.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        sig_btn = tk.Button(
            sig_btn_holder, text="Save Signature", command=self.save_signature,
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=10, pady=4, cursor="hand2", borderwidth=0,
        )
        sig_btn.pack(expand=True)

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
