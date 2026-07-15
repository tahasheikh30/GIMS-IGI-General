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

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

import data_processor as dp
import report_generator as rg
import email_generator as eg
import outlook_sender as osend
import pivot_native_generator as pivot_native
import loss_ratio_processor as lrp
import loss_ratio_generator as lrg
import receivable_processor as rvp
import receivable_generator as rvg
import receivable_email as rve
import receivable_automation as ra
import receivable_run_log as rvlog
import run_log


ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
GIMS_LOGO_PATH = os.path.join(ASSETS_DIR, "gims_logo_cropped.png")
COMPANY_LOGO_PATH = os.path.join(ASSETS_DIR, "company_logo_cropped.png")

# Simple leading glyphs for each sidebar nav item - kept as plain
# characters (no emoji) so they render consistently across Windows font
# sets instead of relying on a specific icon font being installed.
NAV_ICONS = {
    "outstanding_premium": "\u25A3",   # filled-grid square - balances/reporting
    "loss_ratio": "\u25B3",            # triangle - ratio/analysis
    "coinsurance_pc_email": "\u2709",  # envelope - email automation
    "coinsurance_premium_register": "\u2263",  # stacked lines - register/list
    "coinsurance_claims_register": "\u2263",
    "renewal": "\u27F3",               # circular arrow - renewal
}


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
    COLOR_SIDEBAR = "#F9FBFD"       # sidebar background - close to panel-alt, distinct from main canvas
    COLOR_SIDEBAR_ACTIVE = "#E4EDF5"  # active nav item background tint
    COLOR_SIDEBAR_HOVER = "#EEF3F8"   # hover tint, one step lighter than active
    COLOR_SIDEBAR_SUBBG = "#F2F6FA"   # subtle tint behind the Coinsurance submenu block
    COLOR_HEADER_BORDER = "#153252"   # hairline shadow strip under the header, elevates it off the page

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

        self.active_nav = None
        self.loss_group_files = []  # list of filepaths selected in Group mode

        self.receivable_raw_df = None
        self.receivable_clean_df = None
        self.receivable_issues_df = None
        self.receivable_colmap = None
        self.receivable_groups = None
        self.receivable_totals = None
        self.receivable_filtered_df = None
        self.receivable_filename_lookup = {}
        self.receivable_email_lookup = {}
        self.receivable_poc_lookup = {}
        self.receivable_drafts = []

        self._load_logos()
        self._setup_style()
        self._build_ui()

    def _load_logos(self):
        """Load the GIMS wordmark logo and the company logo for the header,
        scaled to a fixed pixel height so both stay crisp and evenly
        matched regardless of their native resolution. Falls back to the
        drawn 'G' badge (see _build_header) if Pillow or the asset files
        aren't available, so the app still runs without either."""
        self.logo_gims_img = None
        self.logo_company_img = None
        if not PIL_AVAILABLE:
            return

        def _load(path, target_height):
            if not os.path.exists(path):
                return None
            try:
                im = Image.open(path).convert("RGBA")
                ratio = target_height / im.height
                w = max(1, int(im.width * ratio))
                im = im.resize((w, target_height), Image.LANCZOS)
                return ImageTk.PhotoImage(im)
            except Exception:
                return None

        self.logo_gims_img = _load(GIMS_LOGO_PATH, 40)
        self.logo_company_img = _load(COMPANY_LOGO_PATH, 34)

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
        """Fixed top header - GIMS mark + wordmark on the left, the
        company logo on the right. Stays visible across every sidebar
        section and never breaks on resize (fixed height, edge-anchored
        content via place(), no elements that depend on window width)."""
        header_wrap = tk.Frame(self, bg=self.COLOR_PRIMARY)
        header_wrap.pack(fill="x", side="top")

        header = tk.Frame(header_wrap, bg=self.COLOR_PRIMARY, height=68)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # Thin darker strip under the header reads as a subtle drop
        # shadow, lifting the header off the page instead of it sitting
        # perfectly flush with the content below.
        tk.Frame(header_wrap, bg=self.COLOR_HEADER_BORDER, height=2).pack(fill="x", side="top")

        # ---- Left: GIMS mark + wordmark ----
        brand = tk.Frame(header, bg=self.COLOR_PRIMARY)
        brand.place(relx=0, rely=0.5, x=22, anchor="w")  # vertically centered, fixed left inset

        if self.logo_gims_img is not None:
            tk.Label(brand, image=self.logo_gims_img, bg=self.COLOR_PRIMARY).pack(side="left", padx=(0, 12))
        else:
            # Fallback drawn mark if Pillow / the asset file isn't available
            mark = tk.Canvas(brand, width=40, height=40, bg=self.COLOR_PRIMARY, highlightthickness=0)
            mark.pack(side="left", padx=(0, 12))
            mark.create_oval(0, 0, 40, 40, fill=self.COLOR_ACCENT, outline="")
            mark.create_text(20, 21, text="G", fill="white", font=("Segoe UI", 17, "bold"))

        wordmark_col = tk.Frame(brand, bg=self.COLOR_PRIMARY)
        wordmark_col.pack(side="left")
        tk.Label(wordmark_col, text="GIMS", font=("Segoe UI Semibold", 20, "bold"),
                 bg=self.COLOR_PRIMARY, fg="white").pack(anchor="w")
        tk.Label(wordmark_col, text="General Insurance Management System",
                 font=("Segoe UI", 8), bg=self.COLOR_PRIMARY, fg="#B9CBDD").pack(anchor="w")

        # ---- Right: company logo ----
        company_slot = tk.Frame(header, bg=self.COLOR_PRIMARY)
        company_slot.place(relx=1.0, rely=0.5, x=-22, anchor="e")  # vertically centered, fixed right inset

        if self.logo_company_img is not None:
            # White pill behind the logo so it reads cleanly against the
            # dark header regardless of the logo's own background.
            pill = tk.Frame(company_slot, bg="white")
            pill.pack()
            tk.Label(pill, image=self.logo_company_img, bg="white").pack(padx=12, pady=6)

    def _build_sidebar(self, parent):
        """Persistent left sidebar. Every section now lives under one of
        two collapsible parent menus: 'General Operations' (Outstanding
        Premium, Loss Ratio, Renewal) and 'Coinsurance' (its existing
        three sub-items). The active section is shown with a left accent
        bar, bold text, and a tinted background, same as before."""
        sidebar = tk.Frame(parent, bg=self.COLOR_SIDEBAR, width=240)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=20).pack(fill="x")
        tk.Label(sidebar, text="MODULES", font=("Segoe UI", 8, "bold"), bg=self.COLOR_SIDEBAR,
                 fg=self.COLOR_MUTED).pack(anchor="w", padx=22, pady=(0, 8))

        self.nav_items = {}

        # Permanent trailing anchor, packed once and never re-packed. Every
        # collapsible menu block below is always inserted immediately ahead
        # of the next fixed sibling (either this anchor, or the header of
        # the following menu), so collapsing/expanding one section can
        # never reorder the others.
        bottom_anchor = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=1)
        bottom_anchor.pack(fill="x")

        coinsurance_header = self._build_collapsible_menu(
            sidebar, "Coinsurance", "\u2295",
            [
                ("coinsurance_pc_email", "Premium/Claims Email Automation"),
                ("coinsurance_premium_register", "Premium Register"),
                ("coinsurance_claims_register", "Claims Register"),
            ],
            before=bottom_anchor,
        )

        self._build_collapsible_menu(
            sidebar, "General Operations", "\u2699",
            [
                ("outstanding_premium", "Outstanding Premium"),
                ("loss_ratio", "Loss Ratio"),
                ("renewal", "Renewal"),
            ],
            before=coinsurance_header,
        )

        # Thin divider between the sidebar and content area
        tk.Frame(parent, bg=self.COLOR_BORDER, width=1).pack(side="left", fill="y")

    def _build_collapsible_menu(self, sidebar, label_text, icon_char, sub_items, before):
        """Generic collapsible parent menu (e.g. 'General Operations' or
        'Coinsurance'), shown inside a subtly tinted block with a left
        indent guide so its sub-items read as nested under the parent.
        Clicking the header only expands/collapses it (it isn't itself a
        page); clicking a sub-item navigates like any other nav item.

        `before`: a permanent fixed sibling widget, packed once and never
        forgotten/re-packed, that this menu's header and submenu are
        always inserted immediately ahead of via pack(before=...). This
        has to be a widget that never moves, otherwise re-expanding a
        submenu after collapsing it could land it in the wrong place
        relative to the other sidebar sections.

        Returns the header frame widget, so an earlier-built menu can use
        it as its own `before` anchor and keep menus stacked in the order
        they were built.
        """
        header = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, cursor="hand2")
        header.pack(fill="x", before=before)

        header_accent = tk.Frame(header, bg=self.COLOR_SIDEBAR, width=4)
        header_accent.pack(side="left", fill="y")

        header_inner = tk.Frame(header, bg=self.COLOR_SIDEBAR)
        header_inner.pack(side="left", fill="both", expand=True, padx=(14, 12), pady=14)

        icon_label = tk.Label(header_inner, text=icon_char, font=("Segoe UI", 10), width=2,
                               bg=self.COLOR_SIDEBAR, fg=self.COLOR_MUTED, anchor="w")
        icon_label.pack(side="left")

        header_label = tk.Label(header_inner, text=label_text, font=("Segoe UI", 11),
                                 bg=self.COLOR_SIDEBAR, fg=self.COLOR_MUTED, anchor="w")
        header_label.pack(side="left", fill="both", expand=True)

        caret = tk.Label(header_inner, text="\u25BE", font=("Segoe UI", 9),
                          bg=self.COLOR_SIDEBAR, fg=self.COLOR_MUTED)
        caret.pack(side="right")

        # Tinted wrapper block for the sub-items, with a thin vertical
        # indent guide on the left so the sub-items read as nested under
        # the parent rather than as flat, unrelated rows.
        submenu_block = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR_SUBBG)
        submenu_block.pack(fill="x", before=before)
        guide = tk.Frame(submenu_block, bg=self.COLOR_BORDER, width=1)
        guide.place(x=28, y=0, relheight=1.0)

        submenu_frame = tk.Frame(submenu_block, bg=self.COLOR_SIDEBAR_SUBBG)
        submenu_frame.pack(fill="x")  # starts expanded

        for key, label in sub_items:
            self.nav_items[key] = self._make_nav_item(
                submenu_frame, key, label, indent=True, base_bg=self.COLOR_SIDEBAR_SUBBG
            )

        def toggle_menu(event=None):
            if submenu_block.winfo_ismapped():
                submenu_block.pack_forget()
                caret.configure(text="\u25B8")
            else:
                submenu_block.pack(fill="x", before=before)
                caret.configure(text="\u25BE")

        for w in (header, header_accent, header_inner, icon_label, caret, header_label):
            w.bind("<Button-1>", toggle_menu)
            w.bind("<Enter>", lambda e: self._hover_menu_header(header, header_accent, header_inner, icon_label, header_label, True))
            w.bind("<Leave>", lambda e: self._hover_menu_header(header, header_accent, header_inner, icon_label, header_label, False))

        return header

    def _hover_menu_header(self, header, header_accent, header_inner, icon_label, header_label, entering):
        bg = self.COLOR_SIDEBAR_HOVER if entering else self.COLOR_SIDEBAR
        for w in (header, header_inner, icon_label, header_label):
            w.configure(bg=bg)
        header_accent.configure(bg=self.COLOR_SIDEBAR_HOVER if entering else self.COLOR_SIDEBAR)

    def _make_nav_item(self, parent, key, text, indent=False, before=None, base_bg=None):
        base_bg = base_bg or self.COLOR_SIDEBAR
        wrapper = tk.Frame(parent, bg=base_bg, cursor="hand2")
        pack_kwargs = {"fill": "x"}
        if before is not None:
            pack_kwargs["before"] = before
        wrapper.pack(**pack_kwargs)

        accent = tk.Frame(wrapper, bg=base_bg, width=4)
        accent.pack(side="left", fill="y")

        inner = tk.Frame(wrapper, bg=base_bg)
        inner.pack(side="left", fill="both", expand=True,
                    padx=(26, 10) if indent else (10, 10), pady=12 if indent else 14)

        icon = NAV_ICONS.get(key, "")
        icon_label = tk.Label(inner, text=icon, font=("Segoe UI", 9 if indent else 10), width=2,
                               bg=base_bg, fg=self.COLOR_MUTED, anchor="w")
        icon_label.pack(side="left")

        label = tk.Label(inner, text=text, font=("Segoe UI", 10 if indent else 11),
                          bg=base_bg, fg=self.COLOR_MUTED, anchor="w",
                          wraplength=160 if indent else None, justify="left")
        label.pack(side="left", fill="both", expand=True)

        def on_click(event=None, k=key):
            self.select_nav(k)

        def on_enter(event=None):
            if self.active_nav != key:
                hover_bg = self.COLOR_SIDEBAR_HOVER if not indent else "#E9EFF5"
                for w in (wrapper, accent, inner, icon_label, label):
                    w.configure(bg=hover_bg)

        def on_leave(event=None):
            if self.active_nav != key:
                for w in (wrapper, accent, inner, icon_label, label):
                    w.configure(bg=base_bg)

        for widget in (wrapper, accent, inner, icon_label, label):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)

        wrapper.accent = accent
        wrapper.inner = inner
        wrapper.icon_label = icon_label
        wrapper.label = label
        wrapper.indent = indent
        wrapper.base_bg = base_bg
        return wrapper

    def select_nav(self, key):
        """Switch the visible content page and restyle the sidebar so only
        the active section (top-level or Coinsurance sub-item) shows the
        accent bar / bold text / tinted bg."""
        self.active_nav = key
        for k, item in self.nav_items.items():
            selected = (k == key)
            bg = self.COLOR_SIDEBAR_ACTIVE if selected else item.base_bg
            base_size = 10 if item.indent else 11
            item.configure(bg=bg)
            item.accent.configure(bg=self.COLOR_ACCENT if selected else bg)
            item.inner.configure(bg=bg)
            item.icon_label.configure(bg=bg, fg=self.COLOR_PRIMARY if selected else self.COLOR_MUTED)
            item.label.configure(
                bg=bg,
                fg=self.COLOR_PRIMARY if selected else self.COLOR_MUTED,
                font=("Segoe UI", base_size, "bold" if selected else "normal"),
            )
        self.nav_pages[key].tkraise()

    def _build_content_area(self, parent):
        """Stacked pages sharing one grid cell; select_nav() raises the
        active one. Outstanding Premium, Loss Ratio, and the Coinsurance
        > Premium/Claims Email Automation page host real workflows; the
        rest are placeholders reserved for future modules."""
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

        pc_email_page = ttk.Frame(container)
        pc_email_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_pc_email"] = pc_email_page
        self._build_receivable_page(pc_email_page)

        premium_register_page = self._build_placeholder_page(container, "Premium Register")
        premium_register_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_premium_register"] = premium_register_page

        claims_register_page = self._build_placeholder_page(container, "Claims Register")
        claims_register_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_claims_register"] = claims_register_page

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
        self._enable_mousewheel(canvas, _on_mousewheel)

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
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 12))

        file_container = ttk.Frame(s2)
        file_container.pack(fill="both", expand=True)
        file_container.columnconfigure(0, weight=1)
        file_container.rowconfigure(0, weight=1)

        # ---- Individual mode: single-file picker card ----
        self.loss_individual_frame = ttk.Frame(file_container)
        self.loss_individual_frame.grid(row=0, column=0, sticky="nsew")

        indiv_card = tk.Frame(self.loss_individual_frame, bg=self.COLOR_PANEL_ALT,
                               highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        indiv_card.pack(fill="both", expand=True)
        indiv_inner = tk.Frame(indiv_card, bg=self.COLOR_PANEL_ALT)
        indiv_inner.pack(fill="x", padx=16, pady=14)

        self.loss_individual_icon = tk.Label(indiv_inner, text="\u25A4", font=("Segoe UI", 20),
                                              bg=self.COLOR_PANEL_ALT, fg=self.COLOR_BORDER, width=2)
        self.loss_individual_icon.pack(side="left", padx=(0, 14))

        indiv_text_col = tk.Frame(indiv_inner, bg=self.COLOR_PANEL_ALT)
        indiv_text_col.pack(side="left", fill="x", expand=True)

        self.loss_individual_file_var = tk.StringVar()
        self.loss_individual_name_var = tk.StringVar(value="No file selected")
        tk.Label(indiv_text_col, textvariable=self.loss_individual_name_var, font=("Segoe UI", 10, "bold"),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_TEXT, anchor="w", wraplength=420, justify="left")\
            .pack(fill="x", anchor="w")
        tk.Label(indiv_text_col, text="Excel (.xlsx) or PDF \u2014 one company", font=("Segoe UI", 8),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_MUTED, anchor="w").pack(fill="x", anchor="w")

        indiv_btn_col = tk.Frame(indiv_inner, bg=self.COLOR_PANEL_ALT)
        indiv_btn_col.pack(side="right", padx=(14, 0))
        self._make_pill_button(indiv_btn_col, "Browse...", self._browse_loss_individual_file,
                                bg=self.COLOR_ACCENT).pack(side="left")

        # ---- Group mode: multi-file list card ----
        self.loss_group_frame = ttk.Frame(file_container)
        self.loss_group_frame.grid(row=0, column=0, sticky="nsew")

        group_card = tk.Frame(self.loss_group_frame, bg=self.COLOR_PANEL_ALT,
                               highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        group_card.pack(fill="x")

        self.loss_group_count_var = tk.StringVar(value="No files added yet")
        count_row = tk.Frame(group_card, bg=self.COLOR_PANEL_ALT)
        count_row.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(count_row, textvariable=self.loss_group_count_var, font=("Segoe UI", 9, "bold"),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_MUTED).pack(side="left")

        list_outer = tk.Frame(group_card, bg=self.COLOR_PANEL_ALT)
        list_outer.pack(fill="x", padx=16, pady=(0, 4))

        list_frame = tk.Frame(list_outer, bg="white", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_frame.pack(fill="x")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.loss_group_listbox = tk.Listbox(list_frame, height=6, font=("Segoe UI", 10), selectmode="extended",
                                              activestyle="none", relief="flat", borderwidth=0,
                                              bg="white", selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                              highlightthickness=0, yscrollcommand=list_scroll.set)
        list_scroll.configure(command=self.loss_group_listbox.yview)
        self.loss_group_listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        # Empty-state placeholder, shown centered over the list until the
        # first file is added, then hidden - clearer than a bare empty box.
        # Wrapped across two short lines with a wraplength tied to the
        # box's actual width (via the <Configure> binding below) so the
        # text never runs wider than the bordered box it sits inside.
        empty_state = tk.Frame(list_frame, bg="white")
        tk.Label(empty_state, text="No files added", font=("Segoe UI", 9, "bold"),
                 bg="white", fg=self.COLOR_MUTED, justify="center").pack()
        self.loss_group_empty_sub = tk.Label(
            empty_state, text="Click Browse below to add Excel or PDF files",
            font=("Segoe UI", 8), bg="white", fg=self.COLOR_MUTED,
            justify="center", wraplength=220,
        )
        self.loss_group_empty_sub.pack(pady=(2, 0))
        self.loss_group_empty_label = empty_state
        self.loss_group_empty_label.place(relx=0.5, rely=0.5, anchor="center")

        def _resize_empty_label(event):
            # Keep the placeholder text within the box with some breathing
            # room on either side, instead of a fixed width that can spill
            # past the border on a narrower window.
            self.loss_group_empty_sub.configure(wraplength=max(140, event.width - 48))
        list_frame.bind("<Configure>", _resize_empty_label)

        # Action toolbar lives BELOW the list (not above it) so the natural
        # reading order is "here are your files, here's what you can do
        # with them" - and "Browse..." (not a separately-labeled "Add
        # File(s)...") matches the same action/label used in Individual
        # mode, since they do exactly the same thing (open a file picker).
        group_toolbar = tk.Frame(group_card, bg=self.COLOR_PANEL_ALT)
        group_toolbar.pack(fill="x", padx=16, pady=(6, 14))
        self._make_pill_button(group_toolbar, "Browse...", self._add_loss_group_files,
                                bg=self.COLOR_ACCENT).pack(side="left")
        self._make_pill_button(group_toolbar, "Remove Selected", self._remove_loss_group_file,
                                bg=self.COLOR_PANEL, fg=self.COLOR_TEXT, outline=True).pack(side="left", padx=(8, 0))
        self._make_pill_button(group_toolbar, "Clear All", self._clear_loss_group_files,
                                bg=self.COLOR_PANEL, fg=self.COLOR_WARN, outline=True).pack(side="left", padx=(8, 0))

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
        self._refresh_loss_group_count()  # sync the empty-state placeholder/count label

    def _enable_mousewheel(self, canvas, on_scroll):
        """canvas.bind_all() is application-wide, and each call silently
        REPLACES any previous global <MouseWheel> binding rather than
        adding to it. Every scrollable page (Outstanding Premium, Loss
        Ratio, Coinsurance) used to call bind_all() once, unconditionally,
        at build time - so whichever page was built LAST (Coinsurance)
        permanently owned the only working scroll handler, and two-finger
        scroll silently did nothing on every page built before it.
        Scoping the binding to <Enter>/<Leave> means each canvas only
        claims global scroll while the cursor is actually over it, and
        releases it immediately on <Leave> - so every scrollable page
        keeps working independently of build order."""
        def _on_enter(event):
            canvas.bind_all("<MouseWheel>", on_scroll)
        def _on_leave(event):
            canvas.unbind_all("<MouseWheel>")
        canvas.bind("<Enter>", _on_enter)
        canvas.bind("<Leave>", _on_leave)

    def _make_pill_button(self, parent, text, command, bg, fg="white", outline=False):
        """Small factory for the flat, rounded-look buttons already used
        elsewhere in the app (Save Signature, Calculate Loss Ratio, etc.)
        so every action button - across Individual and Group file modes -
        shares one consistent size/shape instead of each spot picking its
        own fixed character width."""
        btn = tk.Button(
            parent, text=text, command=command,
            font=("Segoe UI", 9, "bold"), bg=bg, fg=fg,
            activebackground=self.COLOR_SIDEBAR_ACTIVE if outline else self.COLOR_PRIMARY_DARK,
            activeforeground=fg,
            relief="solid" if outline else "flat",
            borderwidth=1 if outline else 0,
            highlightbackground=self.COLOR_BORDER,
            padx=14, pady=7, cursor="hand2",
        )
        return btn

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
            self.loss_individual_name_var.set(os.path.basename(path))
            self.loss_individual_icon.configure(fg=self.COLOR_ACCENT)

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
        self._refresh_loss_group_count()

    def _remove_loss_group_file(self):
        selected = list(self.loss_group_listbox.curselection())
        if not selected:
            messagebox.showinfo("Nothing selected", "Select one or more files in the list to remove.")
            return
        for idx in reversed(selected):
            self.loss_group_listbox.delete(idx)
            del self.loss_group_files[idx]
        self.status_var.set(f"{len(self.loss_group_files)} file(s) remaining in the group.")
        self._refresh_loss_group_count()

    def _clear_loss_group_files(self):
        if not self.loss_group_files:
            return
        if not messagebox.askyesno("Clear all files", f"Remove all {len(self.loss_group_files)} file(s) from the group list?"):
            return
        self.loss_group_listbox.delete(0, "end")
        self.loss_group_files.clear()
        self.status_var.set("Cleared all files from the Loss Ratio group.")
        self._refresh_loss_group_count()

    def _refresh_loss_group_count(self):
        n = len(self.loss_group_files)
        if n == 0:
            self.loss_group_count_var.set("No files added yet")
            self.loss_group_empty_label.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.loss_group_count_var.set(f"{n} file{'s' if n != 1 else ''} selected")
            self.loss_group_empty_label.place_forget()

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

    # ---------- Coinsurance > Premium/Claims Email Automation module ----------
    def _build_receivable_page(self, parent):
        """Premium & Claims Receivable automation, structured as three tabs
        that mirror the Outstanding Premium workflow exactly: Upload &
        Validate, Insurance Company Files, and Review & Send Emails."""
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self.tab_receivable_setup = ttk.Frame(notebook)
        self.tab_receivable_groups = ttk.Frame(notebook)
        self.tab_receivable_emails = ttk.Frame(notebook)
        notebook.add(self.tab_receivable_setup, text="1. Upload & Validate")
        notebook.add(self.tab_receivable_groups, text="2. Insurance Company Files")
        notebook.add(self.tab_receivable_emails, text="3. Review & Send Emails")

        self._build_receivable_setup_tab()
        self._build_receivable_groups_tab()
        self._build_receivable_emails_tab()

    def _build_receivable_setup_tab(self):
        f = self.tab_receivable_setup
        f.configure(style="TFrame")

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
        self._enable_mousewheel(canvas, _on_mousewheel)

        pad = {"padx": 10, "pady": 8}

        # ---- Section: Data Source ----
        source_box = ttk.Labelframe(canvas_frame, text="  Data Source  ", padding=14)
        source_box.pack(fill="x", padx=12, pady=(12, 8))
        source_box.columnconfigure(1, weight=1)

        ttk.Label(source_box, text="Master Workbook (.xlsx)", style="Panel.TLabel").grid(row=0, column=0, sticky="w", **pad)
        self.receivable_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.receivable_file_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_receivable_file).grid(row=0, column=2, **pad)

        ttk.Label(source_box, text="Output folder (workbooks written here)", style="Panel.TLabel").grid(row=1, column=0, sticky="w", **pad)
        self.receivable_output_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "receivable_output"))
        ttk.Entry(source_box, textvariable=self.receivable_output_dir_var).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_receivable_output_dir).grid(row=1, column=2, **pad)

        ttk.Label(source_box, text="Leader Ref Number is required on Premium Receivable rows only \u2014 it's the other "
                                    "company's own policy reference, so it's expected to be blank on Claims Receivable "
                                    "rows and is never flagged there.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9),
                  wraplength=560, justify="left").grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 2))

        self.receivable_load_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(source_box, text="Look up recipient emails from a lookup file",
                         variable=self.receivable_load_email_var).grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))

        ttk.Label(source_box, text="Email Lookup File (.xlsx/.csv, optional)", style="Panel.TLabel").grid(row=4, column=0, sticky="w", **pad)
        self.receivable_email_lookup_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.receivable_email_lookup_file_var).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_receivable_email_lookup_file).grid(row=4, column=2, **pad)
        ttk.Label(source_box, text="A Company-name column, an Email column, and an optional POC/Contact column. "
                                    "Rows are matched to the master file's 'Company' field by exact name. "
                                    "If left blank, drafted emails will be flagged \"missing email\" for manual entry.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9),
                  wraplength=560, justify="left").grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 2))

        ttk.Button(source_box, text="Load & Validate Data", style="Accent.TButton", width=24,
                   command=self.load_receivable_data).grid(row=6, column=0, sticky="w", padx=10, pady=(10, 4))

        # ---- Section: Filters ----
        filter_box = ttk.Labelframe(canvas_frame, text="  Filters  ", padding=14)
        filter_box.pack(fill="x", padx=12, pady=8)
        filter_box.columnconfigure(0, weight=1)

        ttk.Label(filter_box, text="Client filter (optional, comma-separated exact names)",
                  style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=6, pady=(0, 2))
        self.receivable_client_filter_var = tk.StringVar()
        ttk.Entry(filter_box, textvariable=self.receivable_client_filter_var, width=90).grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))

        ttk.Label(filter_box, text="Company filter (optional, comma-separated exact co-insurance company names)",
                  style="Panel.TLabel").grid(row=2, column=0, sticky="w", padx=6, pady=(0, 2))
        self.receivable_company_filter_var = tk.StringVar()
        ttk.Entry(filter_box, textvariable=self.receivable_company_filter_var, width=90).grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))

        ttk.Button(filter_box, text="Analyze / Build Groups", style="Accent.TButton", width=24,
                   command=self.analyze_receivable).grid(row=4, column=0, sticky="w", padx=6, pady=(4, 4))

        # ---- Section: Log ----
        log_box = ttk.Labelframe(canvas_frame, text="  Activity Log  ", padding=10)
        log_box.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.receivable_setup_log = scrolledtext.ScrolledText(log_box, height=10, font=("Consolas", 9),
                                                                bg="#FBFCFD", relief="flat", borderwidth=0)
        self.receivable_setup_log.pack(fill="both", expand=True)

    def _build_receivable_groups_tab(self):
        f = self.tab_receivable_groups

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Generate Excel Reports", style="Accent.TButton", width=24,
                   command=self.generate_receivable_reports).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Open Output Folder", command=self.open_receivable_output_folder, width=20).pack(side="left")
        self.receivable_results_summary_var = tk.StringVar(value="Run Analyze on tab 1 to see one row per output file here.")
        ttk.Label(top, textvariable=self.receivable_results_summary_var, style="Muted.TLabel").pack(side="right")

        hint = ttk.Label(f, text="One workbook is generated per insurance company (covering every client for that "
                                  "company); every company with \"OFC\" in its name is pooled into a single shared \"OFC\" file.",
                          style="Muted.TLabel", font=("Segoe UI", 9))
        hint.pack(anchor="w", padx=14, pady=(0, 6))

        tree_frame = ttk.Frame(f)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("company", "num_clients", "premium_total", "claims_total")
        self.receivable_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.receivable_results_tree.heading("company", text="Insurance Company (output file)")
        self.receivable_results_tree.heading("num_clients", text="Clients")
        self.receivable_results_tree.heading("premium_total", text="Premium Receivable Total")
        self.receivable_results_tree.heading("claims_total", text="Claims Receivable Total")
        self.receivable_results_tree.column("company", width=380)
        self.receivable_results_tree.column("num_clients", width=80, anchor="center")
        self.receivable_results_tree.column("premium_total", width=170, anchor="e")
        self.receivable_results_tree.column("claims_total", width=170, anchor="e")
        self.receivable_results_tree.tag_configure("oddrow", background="#F4F7FA")
        self.receivable_results_tree.tag_configure("evenrow", background="white")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.receivable_results_tree.yview)
        self.receivable_results_tree.configure(yscrollcommand=vsb.set)
        self.receivable_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_receivable_emails_tab(self):
        f = self.tab_receivable_emails

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Draft Emails", style="Accent.TButton", command=self.draft_receivable_emails, width=18).pack(side="left", padx=(0, 12))
        self.receivable_dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Dry run (preview only, do not actually send)", variable=self.receivable_dry_run_var).pack(side="left", padx=(0, 12))
        ttk.Button(top, text="Send All", style="Send.TButton", command=self.send_all_receivable_emails, width=14).pack(side="left")

        sig_frame = ttk.Labelframe(f, text="  Your Signature (appears after 'Regards,' in every email - shared with Outstanding Premium)  ", padding=10)
        sig_frame.pack(fill="x", padx=12, pady=(0, 8))
        sig_frame.columnconfigure(0, weight=1)
        self.receivable_signature_text = tk.Text(sig_frame, height=3, font=("Segoe UI", 10), relief="solid", borderwidth=1, wrap="word")
        self.receivable_signature_text.grid(row=0, column=0, sticky="ew")
        self.receivable_signature_text.insert("1.0", eg.load_signature())

        sig_btn_holder = tk.Frame(sig_frame, bg=self.COLOR_PANEL)
        sig_btn_holder.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        sig_btn = tk.Button(
            sig_btn_holder, text="Save Signature", command=self.save_receivable_signature,
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=10, pady=4, cursor="hand2", borderwidth=0,
        )
        sig_btn.pack(expand=True)

        body = ttk.Frame(f)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(left, text="Insurance Company Files", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        columns = ("company", "clients", "email", "status")
        self.receivable_email_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=26)
        self.receivable_email_tree.heading("company", text="Company")
        self.receivable_email_tree.heading("clients", text="Clients")
        self.receivable_email_tree.heading("email", text="Email")
        self.receivable_email_tree.heading("status", text="Status")
        self.receivable_email_tree.column("company", width=210)
        self.receivable_email_tree.column("clients", width=60, anchor="center")
        self.receivable_email_tree.column("email", width=170)
        self.receivable_email_tree.column("status", width=150)
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.receivable_email_tree.yview)
        self.receivable_email_tree.configure(yscrollcommand=tree_scroll.set)
        self.receivable_email_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.receivable_email_tree.bind("<<TreeviewSelect>>", self.on_select_receivable_email)

        right = ttk.Labelframe(body, text="  Email Preview  ", padding=14)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)

        ttk.Label(right, text="To (recipient email)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.receivable_email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.receivable_subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_subject_var, font=("Segoe UI", 10))\
            .grid(row=2, column=0, sticky="ew")

        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=3, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.receivable_body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1)
        self.receivable_body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        btn_row = ttk.Frame(right)
        btn_row.grid(row=4, column=0, sticky="w")

        save_btn = tk.Button(
            btn_row, text="Save Edits to Selected Email", command=self.save_receivable_email_edit,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_ACCENT, fg="white",
            activebackground=self.COLOR_PRIMARY_DARK, activeforeground="white",
            relief="flat", padx=14, pady=8, cursor="hand2", borderwidth=0,
        )
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = tk.Button(
            btn_row, text="Send This Email", command=self.send_individual_receivable_email,
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_SUCCESS, fg="white",
            activebackground="#1B5E20", activeforeground="white",
            relief="flat", padx=14, pady=8, cursor="hand2", borderwidth=0,
        )
        send_one_btn.pack(side="left")

    # ---------- Receivable actions ----------
    def browse_receivable_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.receivable_file_var.set(path)

    def browse_receivable_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.receivable_output_dir_var.set(path)

    def browse_receivable_email_lookup_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if path:
            self.receivable_email_lookup_file_var.set(path)

    def log_receivable(self, msg):
        self.receivable_setup_log.insert("end", str(msg) + "\n")
        self.receivable_setup_log.see("end")

    def load_receivable_data(self):
        path = self.receivable_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid master workbook.")
            return

        try:
            self.receivable_raw_df = rvp.load_raw_data(path)
            (self.receivable_clean_df, self.receivable_issues_df,
             dropped_columns, self.receivable_colmap) = rvp.validate_data(self.receivable_raw_df)
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))
            return

        self.receivable_setup_log.delete("1.0", "end")
        mode = self.receivable_colmap["mode"]
        self.log_receivable(f"Detected layout: {mode} mode "
                             f"({'Premium Receivable + Claim Receivable amount columns' if mode == 'dual' else 'Type/Category + single Amount column'}).")
        self.log_receivable(f"Loaded {len(self.receivable_raw_df)} rows, {len(self.receivable_clean_df)} valid after validation.")
        if dropped_columns:
            self.log_receivable(f"Removed {len(dropped_columns)} completely empty column(s): {', '.join(dropped_columns)}")
        if len(self.receivable_issues_df) > 0:
            self.log_receivable(f"{len(self.receivable_issues_df)} row(s) flagged with issues:")
            issue_counts = self.receivable_issues_df["issue"].value_counts()
            for issue_label, count in issue_counts.items():
                self.log_receivable(f"   {count}x  {issue_label}")
        else:
            self.log_receivable("No data issues found.")

        self.status_var.set("Receivable data loaded. Set filters and click Analyze / Build Groups.")

    def analyze_receivable(self):
        if self.receivable_clean_df is None:
            messagebox.showerror("Error", "Load data first (tab 1).")
            return

        client_col = self.receivable_colmap["client"]
        company_col = self.receivable_colmap["company"]
        df = self.receivable_clean_df

        client_text = self.receivable_client_filter_var.get().strip()
        if client_text:
            client_filter = [c.strip() for c in client_text.split(",") if c.strip()]
            df = df[df[client_col].isin(client_filter)]

        company_text = self.receivable_company_filter_var.get().strip()
        if company_text:
            company_filter = [c.strip() for c in company_text.split(",") if c.strip()]
            df = df[df[company_col].isin(company_filter)]

        self.receivable_filtered_df = df  # kept for the per-client chart breakdown when generating reports
        self.receivable_groups = rvp.build_company_groups(df)
        self.receivable_totals = rvp.company_totals(df, self.receivable_groups)

        self.receivable_results_tree.delete(*self.receivable_results_tree.get_children())
        for i, (bucket_label, _g) in enumerate(self.receivable_groups.items()):
            t = self.receivable_totals[bucket_label]
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.receivable_results_tree.insert("", "end", values=(
                bucket_label, t["num_clients"], f"{t['premium_total']:,.0f}", f"{t['claims_total']:,.0f}"
            ), tags=(tag,))
        self.receivable_results_summary_var.set(f"{len(self.receivable_groups)} output file(s) identified")

        self.log_receivable(f"\nAnalysis complete: {len(self.receivable_groups)} output file(s) "
                             f"(one per insurance company, OFC-named companies pooled into one shared file).")
        self.status_var.set(f"{len(self.receivable_groups)} files. See tab 2.")

    def generate_receivable_reports(self):
        if not getattr(self, "receivable_groups", None):
            messagebox.showwarning("Nothing to generate", "Run Analyze / Build Groups first (tab 1).")
            return

        output_dir = self.receivable_output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("No output folder", "Set an output folder on tab 1 first.")
            return
        os.makedirs(output_dir, exist_ok=True)

        colmap = self.receivable_colmap
        if colmap["mode"] == "dual":
            premium_amount_col, claims_amount_col = colmap["premium_amount"], colmap["claims_amount"]
        else:
            premium_amount_col = claims_amount_col = colmap["amount"]

        source_df = getattr(self, "receivable_filtered_df", self.receivable_clean_df)

        self.receivable_filename_lookup = {}
        generated = []
        errors = []
        for bucket_label, g in self.receivable_groups.items():
            try:
                breakdown = rvp.client_breakdown_for_bucket(source_df, bucket_label)
                fp = rvg.build_company_workbook(
                    bucket_label, g["premium_rows"], g["claims_rows"], breakdown, output_dir,
                    premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                )
                generated.append(fp)
                self.receivable_filename_lookup[bucket_label] = os.path.basename(fp)
            except Exception as e:
                errors.append(f"{bucket_label}: {e}")

        msg = f"Generated {len(generated)} workbook(s) in:\n{output_dir}"
        if errors:
            msg += f"\n\n{len(errors)} file(s) failed:\n" + "\n".join(errors[:10])
        messagebox.showinfo("Done", msg)
        self.status_var.set(f"Generated {len(generated)} Receivable workbook(s).")

    def open_receivable_output_folder(self):
        output_dir = self.receivable_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate reports first.")

    def draft_receivable_emails(self):
        if not getattr(self, "receivable_groups", None):
            messagebox.showwarning("Nothing to draft", "Run Analyze / Build Groups first (tab 1).")
            return

        self.receivable_email_lookup, self.receivable_poc_lookup = {}, {}
        if self.receivable_load_email_var.get():
            lookup_path = self.receivable_email_lookup_file_var.get().strip()
            if lookup_path and os.path.exists(lookup_path):
                try:
                    self.receivable_email_lookup, self.receivable_poc_lookup = ra.load_email_lookup_from_file(lookup_path)
                    self.log_receivable(f"\nLoaded {len(self.receivable_email_lookup)} company email mapping(s) "
                                         f"from '{os.path.basename(lookup_path)}'.")
                except Exception as e:
                    self.log_receivable(f"\nWARNING: could not load email lookup: {e}")
            else:
                self.log_receivable("\nNo email lookup file selected - all drafts will show 'Missing email'.")

        signature = self.receivable_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)  # shared signature store with Outstanding Premium

        self.receivable_drafts = rve.build_receivable_email_drafts(
            self.receivable_groups, self.receivable_totals, self.receivable_email_lookup,
            poc_lookup=self.receivable_poc_lookup, signature=signature,
        )
        self.receivable_email_tree.delete(*self.receivable_email_tree.get_children())
        for d in self.receivable_drafts:
            self.receivable_email_tree.insert("", "end", values=(
                d["company_name"], d["num_clients"], d["email"] or "(missing)", d["status"]
            ))
        self.status_var.set(f"Drafted {len(self.receivable_drafts)} Receivable emails.")

    def save_receivable_signature(self):
        signature = self.receivable_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)
        self.status_var.set("Signature saved (shared with Outstanding Premium). Click Draft Emails again to refresh existing drafts.")

    def on_select_receivable_email(self, event):
        sel = self.receivable_email_tree.selection()
        if not sel:
            return
        idx = self.receivable_email_tree.index(sel[0])
        d = self.receivable_drafts[idx]
        self.receivable_email_to_var.set(d.get("email") or "")
        self.receivable_subject_var.set(d["subject"])
        self.receivable_body_text.delete("1.0", "end")
        self.receivable_body_text.insert("1.0", d["body"])

    def save_receivable_email_edit(self):
        sel = self.receivable_email_tree.selection()
        if not sel:
            messagebox.showinfo("No company selected", "Select an insurance company from the list on the left first.")
            return
        idx = self.receivable_email_tree.index(sel[0])
        new_email = self.receivable_email_to_var.get().strip()
        self.receivable_drafts[idx]["email"] = new_email or None
        self.receivable_drafts[idx]["subject"] = self.receivable_subject_var.get()
        self.receivable_drafts[idx]["body"] = self.receivable_body_text.get("1.0", "end").strip()

        status = "Ready" if new_email else "Missing email - needs manual entry"
        row_id = self.receivable_email_tree.get_children()[idx]
        self.receivable_email_tree.set(row_id, "email", new_email or "(missing)")
        self.receivable_email_tree.set(row_id, "status", status)

        self.status_var.set(f"Updated draft for {self.receivable_drafts[idx]['company_name']}.")

    def send_all_receivable_emails(self):
        if not getattr(self, "receivable_drafts", None):
            messagebox.showwarning("Nothing to send", "Draft emails first.")
            return

        filename_lookup = getattr(self, "receivable_filename_lookup", {})
        missing = [d["company_name"] for d in self.receivable_drafts
                   if d["company_name"] not in filename_lookup]
        if missing:
            proceed = messagebox.askyesno(
                "Reports not generated",
                f"{len(missing)} of {len(self.receivable_drafts)} file(s) have no generated Excel workbook yet, "
                f"so those emails would go out with no attachment:\n\n"
                + "\n".join(missing[:10]) + (f"\n...and {len(missing) - 10} more" if len(missing) > 10 else "")
                + "\n\nGo to tab 2 and click 'Generate Excel Reports' first, "
                  "or continue anyway without attachments?\n\nContinue anyway?"
            )
            if not proceed:
                return

        dry_run = self.receivable_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send {len(self.receivable_drafts)} emails via Outlook now? This cannot be undone."):
                return

        output_dir = self.receivable_output_dir_var.get().strip()

        def worker():
            results = rve.send_receivable_emails(
                self.receivable_drafts, output_dir, filename_lookup, dry_run=dry_run
            )
            sent = sum(1 for r in results if r["sent"])
            failed = len(results) - sent
            missing_email = sum(1 for d in self.receivable_drafts if not d.get("email"))
            rvlog.log_run(
                num_groups=len(self.receivable_groups),
                num_files_generated=len(filename_lookup),
                num_sent=sent, num_failed=failed,
                num_missing_email=missing_email, dry_run=dry_run,
            )
            for i, r in enumerate(results):
                self.receivable_email_tree.set(self.receivable_email_tree.get_children()[i], "status",
                                                "Sent" if r["sent"] else f"Failed: {r['send_message']}")
            messagebox.showinfo("Done", f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")
            self.status_var.set(f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")

        threading.Thread(target=worker, daemon=True).start()

    def send_individual_receivable_email(self):
        sel = self.receivable_email_tree.selection()
        if not sel:
            messagebox.showinfo("No company selected", "Click an insurance company in the list on the left first.")
            return
        idx = self.receivable_email_tree.index(sel[0])

        self.save_receivable_email_edit()
        d = self.receivable_drafts[idx]

        if not d.get("email"):
            messagebox.showerror("Missing email", "Enter a recipient email address in the 'To' field before sending.")
            return

        output_dir = self.receivable_output_dir_var.get().strip()
        filename_lookup = getattr(self, "receivable_filename_lookup", {})
        fname = filename_lookup.get(d["company_name"])
        attachment_path = os.path.join(output_dir, fname) if fname and os.path.exists(os.path.join(output_dir, fname)) else None
        if attachment_path is None:
            if not messagebox.askyesno(
                "No report generated",
                f"No Excel workbook has been generated yet for {d['company_name']}. "
                f"Send this email without an attachment?\n\n"
                f"(Go to tab 2 and click 'Generate Excel Reports' first if you'd rather attach one.)"
            ):
                return

        dry_run = self.receivable_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send this email to {d['email']} ({d['company_name']}) now?"):
                return

        row_id = self.receivable_email_tree.get_children()[idx]

        def worker():
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run)
            self.receivable_email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
            prefix = "[DRY RUN] " if dry_run else ""
            messagebox.showinfo("Done", f"{prefix}{message}")
            self.status_var.set(f"{prefix}{message}")

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
        self._enable_mousewheel(canvas, _on_mousewheel)

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
                   command=self.load_data).grid(row=4, column=0, sticky="w", padx=10, pady=(10, 4))

        # ---- Section: Threshold & Filters ----
        filter_box = ttk.Labelframe(canvas_frame, text="  Threshold & Filters  ", padding=14)
        filter_box.pack(fill="x", padx=12, pady=8)
        filter_box.columnconfigure(0, weight=1)
        filter_box.columnconfigure(1, weight=1)
        filter_box.columnconfigure(2, weight=1)

        threshold_row = ttk.Frame(filter_box, style="Panel.TFrame")
        threshold_row.grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 10))
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
        self.dept_listbox = tk.Listbox(dept_frame, selectmode="extended", height=11, width=40,
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
        self.agent_listbox = tk.Listbox(agent_frame, selectmode="extended", height=11, width=40,
                                         exportselection=False, font=("Segoe UI", 10),
                                         activestyle="none", relief="flat", borderwidth=0,
                                         selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                         highlightthickness=0,
                                         yscrollcommand=agent_scroll.set)
        agent_scroll.config(command=self.agent_listbox.yview)
        self.agent_listbox.pack(side="left", fill="both", expand=True)
        agent_scroll.pack(side="right", fill="y")

        branch_col = ttk.Frame(filter_box, style="Panel.TFrame")
        branch_col.grid(row=1, column=2, sticky="nsew", padx=6, pady=4)
        ttk.Label(branch_col, text="Branch filter", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(branch_col, text="Ctrl+click to select multiple · none = all", style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
        branch_frame = tk.Frame(branch_col, highlightbackground="#D6DCE1", highlightthickness=1)
        branch_frame.pack(fill="both", expand=True)
        branch_scroll = ttk.Scrollbar(branch_frame, orient="vertical")
        self.branch_listbox = tk.Listbox(branch_frame, selectmode="extended", height=11, width=40,
                                          exportselection=False, font=("Segoe UI", 10),
                                          activestyle="none", relief="flat", borderwidth=0,
                                          selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                          highlightthickness=0,
                                          yscrollcommand=branch_scroll.set)
        branch_scroll.config(command=self.branch_listbox.yview)
        self.branch_listbox.pack(side="left", fill="both", expand=True)
        branch_scroll.pack(side="right", fill="y")

        company_row = ttk.Frame(filter_box, style="Panel.TFrame")
        company_row.grid(row=2, column=0, columnspan=3, sticky="ew", padx=6, pady=(10, 4))
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

        columns = ("client", "total", "agent", "branch")
        self.results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.results_tree.heading("client", text="Client / Participant")
        self.results_tree.heading("total", text="Grand Total (O/s)")
        self.results_tree.heading("agent", text="Agent(s) / Agency")
        self.results_tree.heading("branch", text="Branch(es)")
        self.results_tree.column("client", width=380)
        self.results_tree.column("total", width=150, anchor="e")
        self.results_tree.column("agent", width=220)
        self.results_tree.column("branch", width=220)
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

        self.branch_listbox.delete(0, "end")
        if "Branch desc" in self.clean_df.columns:
            for i, branch in enumerate(sorted(self.clean_df["Branch desc"].dropna().unique())):
                self.branch_listbox.insert("end", branch)
                self.branch_listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

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
        selected_branches = [self.branch_listbox.get(i) for i in self.branch_listbox.curselection()]
        company_text = self.company_filter_var.get().strip()
        company_filter = [c.strip() for c in company_text.split(",") if c.strip()] if company_text else None

        self.filtered_df, self.grand_totals = dp.build_pivot(
            self.clean_df, company_filter=company_filter,
            department_filter=selected_depts or None,
            agent_filter=selected_agents or None,
            branch_filter=selected_branches or None,
        )
        self.qualifying = dp.qualifying_clients(self.filtered_df, self.grand_totals, threshold)

        self.results_tree.delete(*self.results_tree.get_children())
        for i, (client, total) in enumerate(self.qualifying.items()):
            agents = ", ".join(dp.client_agents(self.filtered_df, client))
            branches = ", ".join(dp.client_branches(self.filtered_df, client))
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.results_tree.insert("", "end", values=(client, f"{total:,.0f}", agents, branches), tags=(tag,))
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
