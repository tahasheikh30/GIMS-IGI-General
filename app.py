"""
Outstanding Balance Reporting & Client Email System
Desktop GUI (Tkinter) - Finance department automation tool.

Run with:  python app.py
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import date, datetime

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
import premium_register_processor as prp
import premium_register_generator as prg
import claims_register_processor as crp
import claims_register_generator as crg
import claims_register_email as cre
import claims_register_run_log as crlog
import claims_check_processor as ccp
import claims_check_generator as ccg
import premium_register_company_processor as prcp
import premium_register_company_generator as prcg2
import premium_register_company_email as prce2
import premium_register_company_run_log as prclog2
import perf_utils
import upr_processor as uprp
import upr_generator as uprg
import upr_ledger_processor as uprlp
import reinsurance_processor as rip
import reinsurance_generator as rig
import deferred_commission_processor as dcp
import deferred_commission_generator as dcg

try:
    import keyring
    KEYRING_AVAILABLE = True
except ImportError:  # keyring not installed - password slots still work, just have to be
    KEYRING_AVAILABLE = False  # retyped each run instead of saved to the OS credential store


ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
GIMS_LOGO_PATH = os.path.join(ASSETS_DIR, "gims_logo_cropped.png")
COMPANY_LOGO_PATH = os.path.join(ASSETS_DIR, "company_logo_cropped.png")


class BalanceApp(tk.Tk):
    # Palette - a refined "deep navy + sapphire + muted gold" scheme,
    # chosen for a more premium, editorial feel than a generic flat-UI
    # blue/grey kit: the navy carries authority for chrome (header,
    # sidebar accents, table headers), sapphire is reserved for
    # interactive/actionable elements (buttons, links, active states),
    # and the muted gold is used sparingly as a single "signature" accent
    # (numbered badges, section rules, focus highlights) rather than
    # spread across the UI, which is what keeps an accent color reading
    # as elegant instead of decorative.
    COLOR_BG = "#F2F4F8"          # soft, cool neutral canvas
    COLOR_PANEL = "#FFFFFF"
    COLOR_PANEL_ALT = "#F7F9FC"   # subtle tint for nested/alternating panels
    COLOR_PRIMARY = "#11263F"     # deep navy - chrome, headers, emphasis text
    COLOR_PRIMARY_DARK = "#0A1B2E"
    COLOR_ACCENT = "#2E6E9E"      # sapphire - primary interactive color
    COLOR_ACCENT_DARK = "#1F5279"
    COLOR_ACCENT_SOFT = "#EAF1F8"  # very light sapphire wash for subtle fills
    COLOR_GOLD = "#B8925A"        # muted gold - the one "signature" accent
    COLOR_GOLD_DEEP = "#8C6633"   # deeper gold - used wherever WHITE text sits directly on gold
                                   # (e.g. badge numbers); the lighter COLOR_GOLD above reads fine
                                   # as text/icons on navy but only clears ~2.9:1 against white,
                                   # short of the ~4.5:1 needed for small bold text to stay crisp
    COLOR_GOLD_SOFT = "#F5EEE2"
    COLOR_SUCCESS = "#1E8354"
    COLOR_WARN = "#B4791C"
    COLOR_DANGER = "#B3402F"
    COLOR_TEXT = "#16212E"
    COLOR_MUTED = "#5E6E80"
    COLOR_BORDER = "#DCE3ED"
    # The sidebar is deliberately its own dark surface (not a tint of the
    # light content canvas) - a strong, well-known enterprise pattern
    # (dark nav rail + light workspace) that gives the app a clear
    # navigational anchor and reads noticeably more premium than a sidebar
    # that just blends into the page. Every text/icon color paired with it
    # below was chosen for contrast against this specific navy, not
    # reused from the light-surface palette above.
    COLOR_SIDEBAR = "#0E2138"           # base sidebar navy (deliberately distinct from COLOR_PRIMARY)
    COLOR_SIDEBAR_SUBBG = "#0A192C"     # recessed panel behind a submenu - reads as "nested"
    COLOR_SIDEBAR_HOVER = "#17324F"     # hover feedback, one clear step up from base
    COLOR_SIDEBAR_ACTIVE = "#1E3F60"    # active "pill" background, brightest of the three states
    COLOR_SIDEBAR_GUIDE = "#25405E"     # faint vertical indent guide under a submenu
    COLOR_SIDEBAR_TEXT = "#B8C6D9"      # inactive label/icon - muted but comfortably legible on navy
    COLOR_SIDEBAR_TEXT_DIM = "#7E90A6"  # de-emphasized elements (caret, decorative glyphs)
    COLOR_SIDEBAR_TEXT_ACTIVE = "#FFFFFF"  # active/hover label - full contrast
    COLOR_SIDEBAR_DIVIDER = "#081422"   # hairline seam between the sidebar and the content canvas
    COLOR_HEADER_BORDER = "#0A1B2E"   # hairline shadow strip under the header, elevates it off the page

    def __init__(self):
        super().__init__()
        self.title("GIMS - General Insurance Management Software")
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
        # Base folder is named after the module; every "Generate Excel
        # Reports" click writes into a fresh dated/timed subfolder under
        # it (see _new_run_folder), rather than every run's files landing
        # together in one flat folder. self.output_dir tracks whichever
        # folder is "current" for Draft/Send Emails (tab 3) to attach
        # from - the base folder until the first generation, then that
        # generation's own subfolder from then on.
        self.output_dir_base = os.path.join(os.getcwd(), "Outstanding Premium")
        self.output_dir = self.output_dir_base

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
        self.receivable_last_run_dir = None  # set to the current generation's own dated/timed subfolder once generated
        # Path the currently-cached receivable_raw_df/clean_df/colmap were
        # loaded from - lets "Load, Validate & Build Groups" reuse
        # whatever "Load Filter Options" already read (if it's for this
        # same file) instead of reading it a second time, same cache-reuse
        # pattern Claims Intimation's own filter loading already uses.
        self.receivable_loaded_path = None

        self.premreg_leader_file = ""
        self.premreg_register_file = ""
        self.premreg_result_df = None

        self.claimscheck_leader_file = ""
        self.claimscheck_register_file = ""
        self.claimscheck_result_df = None

        self.claims_intimation_file = ""
        self.claims_premium_register_file = ""
        self.claims_pmd_col = None  # real policy-reference column name from the last matched run - see load_and_match_claims
        self.claims_matched_df = None
        self.claims_unmatched_df = None
        self.claims_groups = None
        self.claims_totals = None
        self.claims_breakdowns = None
        self.claims_filename_lookup = {}
        self.claims_email_lookup = {}
        self.claims_poc_lookup = {}
        self.claims_drafts = []
        self.claims_output_dir = os.path.join(os.getcwd(), "Claims Register")
        self.claims_last_run_dir = None  # set to the current generation's own dated/timed subfolder once generated
        # self.claims_business_type_var ("conventional"/"takaful") is a
        # tk.StringVar created in _build_claims_setup_tab (see that method
        # for why every tk.Variable in this app is created there rather
        # than in __init__) - read at generate/draft time by
        # generate_claims_register_reports() / draft_claims_register_emails()
        # and threaded through to claims_register_generator.py /
        # claims_register_email.py so a Takaful run gets "Takaful" worked
        # into every generated filename, the email subject, and the email
        # body - a Conventional run (the default) is worded exactly as
        # before, with nothing added.
        # Cache for the Unit/Client Name/Department filter listboxes - set
        # by _load_claims_filter_options() (the "Load Filter Options"
        # button) so the file only has to be read once: load_and_match_claims()
        # reuses this cached dataframe instead of re-reading the same file
        # from disk, as long as the file path hasn't changed since.
        self.claims_intimation_df = None
        self.claims_intimation_pmd_col = None
        self.claims_intimation_colmap = None
        self.claims_intimation_loaded_path = None

        # Premium Register (Coinsurance company split) module state
        self.premregsplit_file = ""
        self.premregsplit_df = None
        self.premregsplit_colmap = None
        self.premregsplit_loaded_path = None
        self.premregsplit_groups = None
        self.premregsplit_totals = None
        self.premregsplit_breakdowns = None
        self.premregsplit_filename_lookup = {}
        self.premregsplit_email_lookup = {}
        self.premregsplit_poc_lookup = {}
        self.premregsplit_drafts = []
        self.premregsplit_output_dir = os.path.join(os.getcwd(), "Premium Register")
        self.premregsplit_last_run_dir = None
        # self.premregsplit_business_type_var ("conventional"/"takaful") -
        # same pattern as self.claims_business_type_var above (see that
        # comment): a tk.StringVar created in _build_premregsplit_setup_tab,
        # read at generate/draft time and threaded through to
        # premium_register_company_generator.py /
        # premium_register_company_email.py.

        # UPR Automation module state
        self.upr_register_paths = []           # list of uploaded premium register filepaths
        self.upr_loaded_df = None               # concatenated, cleaned register rows (Department/Comm Date/Gross Premium)
        self.upr_period_options = []             # [{"year","month","label"}, ...] populated after loading
        self.upr_last_result = None              # last run_upr_calculation() result dict
        self.upr_output_dir = os.path.join(os.getcwd(), "UPR Automation")
        self.upr_last_output_file = None
        self.upr_ledger_pdf_path = None          # last-browsed Ledger Balance PDF path
        self.upr_ledger_parse_result = None      # last upr_ledger_processor.parse_ledger_balance_pdf() result

        # Reinsurance Automation module state
        self.reins_register_paths = []           # list of uploaded RI premium ceded register filepaths
        self.reins_loaded_df = None               # concatenated, cleaned register rows (Department/Comm Date/ceded components)
        self.reins_period_options = []             # [{"year","month","label"}, ...] populated after loading
        self.reins_last_result = None              # last run_reinsurance_calculation() result dict
        self.reins_last_opening = 0.0              # Opening Prepaid Reinsurance Ceded entered on tab 2, used at generate time
        self.reins_output_dir = os.path.join(os.getcwd(), "Reinsurance Automation")
        self.reins_last_output_file = None

        # Deferred Commission Expense module state
        self.dc_register_paths = []              # list of uploaded Premium Register filepaths
        self.dc_direct_paths = []                # list of uploaded Commission Exp - Direct ledger filepaths
        self.dc_coinsurance_paths = []            # list of uploaded Commission Exp - Coinsurance Accepted ledger filepaths
        self.dc_service_charges_paths = []        # list of uploaded Service Charges - Coinsurance ledger filepaths
        self.dc_loaded_df = None                  # concatenated Premium Register rows (Department/Comm Date/Doc Ref/Gross Premium) - used only to populate the period dropdown
        self.dc_period_options = []               # [{"year","month","label"}, ...] populated after loading the register
        self.dc_last_result = None                # last run_deferred_commission_calculation() result dict
        self.dc_last_opening = 0.0                # Opening Deferred Commission Expense entered on tab 2, used at generate time
        self.dc_output_dir = os.path.join(os.getcwd(), "Deferred Commission")
        self.dc_last_output_file = None

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

        # Typography scale - one family throughout (Segoe UI) with a
        # deliberate weight/size hierarchy rather than ad-hoc sizes, so
        # every screen reads with the same rhythm: 17/12/10.5/9 for
        # page title / section heading / body / caption.
        base_font = ("Segoe UI", 10)
        heading_font = ("Segoe UI Semibold", 12)
        title_font = ("Segoe UI Semibold", 18)

        style.configure(".", font=base_font, background=self.COLOR_BG)
        style.configure("TFrame", background=self.COLOR_BG)
        style.configure("Panel.TFrame", background=self.COLOR_PANEL)
        style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=base_font)
        style.configure("Panel.TLabel", background=self.COLOR_PANEL, foreground=self.COLOR_TEXT, font=base_font)
        style.configure("Muted.TLabel", background=self.COLOR_BG, foreground=self.COLOR_MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=self.COLOR_PRIMARY, foreground="white", font=title_font)
        style.configure("SectionHeader.TLabel", background=self.COLOR_BG, foreground=self.COLOR_PRIMARY, font=heading_font)

        style.configure("TLabelframe", background=self.COLOR_PANEL, bordercolor=self.COLOR_BORDER, relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=self.COLOR_PANEL, foreground=self.COLOR_PRIMARY, font=heading_font)

        current_theme = style.theme_use()
        using_native = current_theme in ("vista", "xpnative")

        style.configure("TNotebook", background=self.COLOR_BG, borderwidth=0)
        if using_native:
            # Native Windows themes already render clean, readable tabs -
            # forcing custom colors here tends to look worse, not better.
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(18, 9))
        else:
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(20, 11))
            style.map("TNotebook.Tab",
                      background=[("selected", self.COLOR_PANEL), ("!selected", "#E7EBF1")],
                      foreground=[("selected", self.COLOR_PRIMARY), ("!selected", self.COLOR_MUTED)])

        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        if using_native:
            # Native theme buttons don't reliably support background-color
            # overrides, so keep them native-styled but bold/larger for emphasis.
            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9))
            style.configure("Send.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9))
        else:
            # Plain TButton ("Browse...", "Open Output Folder") gets a real
            # secondary-button treatment - a crisp navy border and label on
            # white, with a soft sapphire wash on hover/press - instead of
            # the flat, low-contrast grey a bare ttk button renders as by
            # default. This is what every "Browse..." picker button and
            # "Open Output Folder" button renders through app-wide.
            style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7),
                            background=self.COLOR_PANEL, foreground=self.COLOR_PRIMARY,
                            bordercolor=self.COLOR_BORDER, lightcolor=self.COLOR_PANEL,
                            darkcolor=self.COLOR_PANEL, focusthickness=1, focuscolor=self.COLOR_ACCENT)
            style.map("TButton",
                      background=[("pressed", self.COLOR_ACCENT_SOFT), ("active", self.COLOR_ACCENT_SOFT)],
                      bordercolor=[("pressed", self.COLOR_ACCENT), ("active", self.COLOR_ACCENT)],
                      foreground=[("disabled", "#A9B7C2")])

            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9),
                            background=self.COLOR_ACCENT, foreground="white", borderwidth=0)
            style.map("Accent.TButton",
                      background=[("pressed", self.COLOR_ACCENT_DARK), ("active", self._shade(self.COLOR_ACCENT, -18)),
                                  ("disabled", "#B7C3D1")])
            style.configure("Send.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9),
                            background=self.COLOR_SUCCESS, foreground="white", borderwidth=0)
            style.map("Send.TButton",
                      background=[("pressed", "#124A2C"), ("active", self._shade(self.COLOR_SUCCESS, -18)),
                                  ("disabled", "#B7C3D1")])

        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28,
                        background="white", fieldbackground="white",
                        bordercolor=self.COLOR_BORDER, borderwidth=0)
        if using_native:
            # Native Windows themes (vista/xpnative) don't reliably honor a
            # custom Treeview.Heading background - it silently keeps the
            # OS's own light heading chrome no matter what's set here. The
            # old unconditional config below still forced foreground="white"
            # on top of that untouched light background, which is exactly
            # what made every column header unreadable (white-on-white) on
            # native themes. So on native themes, only bold/size the
            # heading font and leave its colors alone, the same treatment
            # already given to buttons a few lines up.
            style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10, "bold"),
                            padding=(10, 8))
        else:
            style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10, "bold"),
                            background=self.COLOR_PRIMARY, foreground="white", relief="flat",
                            padding=(10, 8))
            style.map("Treeview.Heading", background=[("active", self.COLOR_PRIMARY_DARK)])
        style.map("Treeview", background=[("selected", self.COLOR_ACCENT)], foreground=[("selected", "white")])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # drop the sunken native border

        style.configure("TEntry", padding=6, bordercolor=self.COLOR_BORDER, relief="solid")
        style.map("TEntry", bordercolor=[("focus", self.COLOR_ACCENT)])
        style.configure("TCombobox", padding=6)
        style.configure("TCheckbutton", background=self.COLOR_BG, font=base_font)
        style.configure("TRadiobutton", background=self.COLOR_BG, font=base_font)
        style.configure("Horizontal.TSeparator", background=self.COLOR_BORDER)
        style.configure("Horizontal.TScrollbar", background=self.COLOR_PANEL_ALT, troughcolor=self.COLOR_BG, bordercolor=self.COLOR_BG)
        style.configure("Vertical.TScrollbar", background="#C7D2E0", troughcolor=self.COLOR_PANEL_ALT,
                        bordercolor=self.COLOR_PANEL_ALT, arrowsize=13, width=13)
        style.map("Vertical.TScrollbar", background=[("active", self.COLOR_ACCENT)])

        # Sidebar's own scrollbar needs dark colors to sit on the navy
        # sidebar instead of the light default above (which was designed
        # for the light content canvas) - same thumb/trough relationship,
        # just recolored for a dark surface.
        style.configure("Sidebar.Vertical.TScrollbar", background=self.COLOR_SIDEBAR_HOVER,
                        troughcolor=self.COLOR_SIDEBAR, bordercolor=self.COLOR_SIDEBAR,
                        arrowsize=11, width=10)
        style.map("Sidebar.Vertical.TScrollbar", background=[("active", self.COLOR_ACCENT)])

    def _bind_dynamic_wraplength(self, label, container, side_padding=56):
        """Keep `label`'s wraplength synced to `container`'s own current
        width (e.g. a Labelframe section that stretches with the window),
        instead of a fixed guess like wraplength=560 - which wraps
        correctly at the width it was tuned for, but leaves a description
        paragraph needlessly narrow (or, on a smaller window, still
        overflowing) at any other width. add="+" on the bind so several
        description labels can each track the same container without one
        binding overwriting another's.

        side_padding must cover everything eating into the container's
        raw width before text can actually render: a ttk.Labelframe's own
        `padding` (currently 14 per side = 28) PLUS the label's own grid
        `padx` (currently 10 per side = 20), for ~48px total at today's
        call sites - the default here is 56 for a small safety margin.
        `container.winfo_width()` (what the <Configure> event reports) is
        the frame's outer allocated width, which does NOT already have
        its own padding subtracted, so skipping this would leave the
        wraplength wider than the actually-visible text area and text
        would run past the right edge instead of wrapping - the current
        call sites all share this same Labelframe(padding=14) + grid
        padx=10 shape; pass a different side_padding if a future caller's
        container/grid padding differs from this."""
        def _update(event):
            label.configure(wraplength=max(200, event.width - side_padding))
        container.bind("<Configure>", _update, add="+")

    @staticmethod
    def _new_run_folder(base_dir):
        """Create and return a fresh, uniquely-named subfolder under a
        module's own base output folder (e.g. 'Outstanding Premium/2026-07-22
        14-35-02'), named for the date and time of this generation. Every
        'Generate Excel Reports'-style click gets its own folder this way,
        rather than every run's files landing together in one shared flat
        folder - each module's base folder (named after the module) ends
        up holding one dated/timed subfolder per generation."""
        run_dir = os.path.join(base_dir, datetime.now().strftime("%Y-%m-%d %H-%M-%S"))
        os.makedirs(run_dir, exist_ok=True)
        return run_dir

    @staticmethod
    def _format_period_text(from_var, to_var):
        """Turn a pair of YYYY-MM-DD date-filter StringVars into readable
        text for an email body, e.g. "the period from 01 January 2026 to
        31 July 2026". Falls back gracefully if either side is blank or
        unparseable (the raw text is used as-is rather than silently
        dropping it), and to a generic phrase if both are blank."""
        def _fmt(raw):
            raw = (raw or "").strip()
            if not raw:
                return None
            try:
                return datetime.strptime(raw, "%Y-%m-%d").strftime("%d %B %Y")
            except ValueError:
                return raw

        start_text = _fmt(from_var.get())
        end_text = _fmt(to_var.get())
        if start_text and end_text:
            return f"the period from {start_text} to {end_text}"
        if start_text:
            return f"the period from {start_text} onwards"
        if end_text:
            return f"the period through {end_text}"
        return None

    @staticmethod
    def _cleanup_if_empty_dir(path):
        """Remove `path` if it exists and is empty - used after a Save As
        dialog is cancelled for a module (Loss Ratio, Policy Check, Claims
        Check) whose default save location is a freshly-created
        _new_run_folder() subfolder created purely to seed the dialog's
        initialdir. Without this, cancelling the dialog left behind a
        permanently empty, timestamped folder under the module's base
        output folder every time - harmless to functionality but clutter
        that accumulates indefinitely. Only ever removes a directory
        that's genuinely empty (os.rmdir already refuses a non-empty one),
        so this can never delete a file the person actually saved."""
        try:
            if path and os.path.isdir(path) and not os.listdir(path):
                os.rmdir(path)
        except OSError:
            pass  # best-effort cleanup only - never worth failing the flow over

    # ---------- UI construction ----------
    def _build_ui(self):
        self._build_header()

        body = tk.Frame(self, bg=self.COLOR_BG)
        body.pack(fill="both", expand=True, side="top")

        self._build_sidebar(body)
        self._build_content_area(body)

        tk.Frame(self, bg=self.COLOR_BORDER, height=1).pack(fill="x", side="bottom")
        status_frame = tk.Frame(self, bg=self.COLOR_PANEL_ALT, height=30)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(status_frame, textvariable=self.status_var, bg=self.COLOR_PANEL_ALT, fg=self.COLOR_MUTED,
                 font=("Segoe UI", 9), anchor="w").pack(fill="both", padx=16, pady=5)

    def _build_header(self):
        """Fixed top header - GIMS mark + wordmark on the left, the
        company logo on the right. Stays visible across every sidebar
        section and never breaks on resize (fixed height, edge-anchored
        content via place(), no elements that depend on window width)."""
        header_wrap = tk.Frame(self, bg=self.COLOR_PRIMARY)
        header_wrap.pack(fill="x", side="top")

        header = tk.Frame(header_wrap, bg=self.COLOR_PRIMARY, height=74)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # A slender gold rule is the header's one "signature" accent -
        # it reads as an intentional, editorial flourish (think a gilt
        # line under a masthead) rather than another slab of color, and
        # sits above the darker drop-shadow strip so the header still
        # lifts cleanly off the page beneath it.
        tk.Frame(header_wrap, bg=self.COLOR_GOLD, height=2).pack(fill="x", side="top")
        tk.Frame(header_wrap, bg=self.COLOR_HEADER_BORDER, height=1).pack(fill="x", side="top")

        # ---- Left: GIMS mark + wordmark ----
        brand = tk.Frame(header, bg=self.COLOR_PRIMARY)
        brand.place(relx=0, rely=0.5, x=26, anchor="w")  # vertically centered, fixed left inset

        if self.logo_gims_img is not None:
            tk.Label(brand, image=self.logo_gims_img, bg=self.COLOR_PRIMARY).pack(side="left", padx=(0, 14))
        else:
            # Fallback drawn mark if Pillow / the asset file isn't available -
            # a slim gold ring around the navy disc echoes the header's own
            # gold accent line instead of reading as a generic flat icon.
            mark = tk.Canvas(brand, width=42, height=42, bg=self.COLOR_PRIMARY, highlightthickness=0)
            mark.pack(side="left", padx=(0, 14))
            mark.create_oval(1, 1, 41, 41, fill=self.COLOR_PRIMARY_DARK, outline=self.COLOR_GOLD, width=1.5)
            mark.create_text(21, 22, text="G", fill=self.COLOR_GOLD, font=("Segoe UI Semibold", 18, "bold"))

        wordmark_col = tk.Frame(brand, bg=self.COLOR_PRIMARY)
        wordmark_col.pack(side="left")
        tk.Label(wordmark_col, text="GIMS", font=("Segoe UI Semibold", 21, "bold"),
                 bg=self.COLOR_PRIMARY, fg="white").pack(anchor="w")
        tk.Label(wordmark_col, text="G E N E R A L   I N S U R A N C E   M A N A G E M E N T   S O F T W A R E",
                 font=("Segoe UI", 7, "bold"), bg=self.COLOR_PRIMARY, fg=self.COLOR_GOLD).pack(anchor="w")

        # ---- Right: company logo ----
        company_slot = tk.Frame(header, bg=self.COLOR_PRIMARY)
        company_slot.place(relx=1.0, rely=0.5, x=-26, anchor="e")  # vertically centered, fixed right inset

        if self.logo_company_img is not None:
            # White pill behind the logo so it reads cleanly against the
            # dark header regardless of the logo's own background.
            pill = tk.Frame(company_slot, bg="white", highlightbackground=self.COLOR_GOLD, highlightthickness=1)
            pill.pack()
            tk.Label(pill, image=self.logo_company_img, bg="white").pack(padx=13, pady=7)

    def _build_sidebar(self, parent):
        """Persistent left sidebar. Every section now lives under one of
        two collapsible parent menus: 'General Operations' (Outstanding
        Premium, Loss Ratio) and 'Coinsurance' (its existing
        three sub-items). The active section is shown with a gold accent
        bar, full-white bold text, and a clearly brighter "pill"
        background, so the current page is unmistakable at a glance.

        The module list itself is wrapped in a canvas+scrollbar (the same
        pattern every other scrollable page in the app already uses - see
        _enable_mousewheel) since expanding every collapsible menu (e.g.
        Coinsurance's nested Claims submenu) can make the list taller than
        a shorter window, and it needs to scroll rather than clip. The
        small footer strip stays pinned to the very bottom of the sidebar,
        outside the scrollable area, so it's always visible regardless of
        scroll position."""
        sidebar_outer = tk.Frame(parent, bg=self.COLOR_SIDEBAR, width=254)
        sidebar_outer.pack(side="left", fill="y")
        sidebar_outer.pack_propagate(False)

        # Footer packed first (side="bottom") so it permanently claims its
        # own slot at the bottom of the sidebar; the scrollable area below
        # is packed to fill whatever space is left above it.
        footer = tk.Frame(sidebar_outer, bg=self.COLOR_SIDEBAR_SUBBG)
        footer.pack(side="bottom", fill="x")
        tk.Frame(footer, bg=self.COLOR_SIDEBAR_GUIDE, height=1).pack(fill="x", side="top")
        tk.Label(footer, text="GIMS \u00b7 Finance Automation Suite", font=("Segoe UI", 8),
                 bg=self.COLOR_SIDEBAR_SUBBG, fg=self.COLOR_SIDEBAR_TEXT_DIM,
                 anchor="w").pack(fill="x", padx=24, pady=10)

        scroll_area = tk.Frame(sidebar_outer, bg=self.COLOR_SIDEBAR)
        scroll_area.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(scroll_area, bg=self.COLOR_SIDEBAR, highlightthickness=0)
        vscroll = ttk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview,
                                 style="Sidebar.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        sidebar = tk.Frame(canvas, bg=self.COLOR_SIDEBAR)
        sidebar_window = canvas.create_window((0, 0), window=sidebar, anchor="nw")

        def _on_sidebar_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        sidebar.bind("<Configure>", _on_sidebar_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(sidebar_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self._enable_mousewheel(canvas, _on_mousewheel)

        tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=24).pack(fill="x")
        tk.Label(sidebar, text="M O D U L E S", font=("Segoe UI", 8, "bold"), bg=self.COLOR_SIDEBAR,
                 fg=self.COLOR_GOLD).pack(anchor="w", padx=24, pady=(0, 12))

        self.nav_items = {}

        # Permanent trailing anchor, packed once and never re-packed. Every
        # collapsible menu block below is always inserted immediately ahead
        # of the next fixed sibling (either this anchor, or the header of
        # the following menu), so collapsing/expanding one section can
        # never reorder the others.
        bottom_anchor = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=1)
        bottom_anchor.pack(fill="x")

        coinsurance_header = self._build_collapsible_menu(
            sidebar, "Coinsurance",
            [
                ("coinsurance_pc_email", "Premium/Claims Email Automation"),
                # "Premium" nested submenu - holds the page that used to be
                # the plain "Premium Register" leaf item, renamed to
                # "Policy Check" and nested one level deeper. The page's
                # own key/wiring (coinsurance_premium_register -> its
                # nav_pages entry) is unchanged - only where it sits in
                # the sidebar changed, same treatment as "Claims" below.
                (None, "Premium", [
                    ("coinsurance_premium_register", "Policy Check"),
                    ("coinsurance_premium_register_split", "Premium Register"),
                ]),
                # Nested submenu marker (None, label, [(key, label), ...]) -
                # "Claims" is itself a further collapsible menu, currently
                # holding the one page that used to be the plain "Claims
                # Register" leaf item, renamed to "Claims Intimation" and
                # nested one level deeper. The page's own key/wiring
                # (coinsurance_claims_register -> its nav_pages entry) is
                # unchanged - only where it sits in the sidebar changed.
                (None, "Claims", [
                    ("coinsurance_claims_register", "Claims Intimation"),
                    ("coinsurance_claims_check", "Claims Check"),
                ]),
            ],
            before=bottom_anchor,
        )

        finance_header = self._build_collapsible_menu(
            sidebar, "Finance",
            [
                # "Management Accounts" nested submenu - holds the two
                # pages that used to sit directly under "Finance" as flat
                # leaf items, nested one level deeper. Each page's own
                # key/wiring (upr_automation / reinsurance_automation ->
                # their nav_pages entries) is unchanged - only where they
                # sit in the sidebar changed, same treatment already used
                # for Coinsurance's own "Premium"/"Claims" submenus.
                (None, "Management Accounts", [
                    ("upr_automation", "UPR Automation"),
                    ("reinsurance_automation", "Reinsurance Automation"),
                    ("deferred_commission", "Deferred Commission"),
                ]),
            ],
            before=coinsurance_header,
        )

        self._build_collapsible_menu(
            sidebar, "General Operations",
            [
                ("outstanding_premium", "Outstanding Premium"),
                ("loss_ratio", "Loss Ratio"),
            ],
            before=finance_header,
        )

        # A little breathing room below the last menu so it doesn't sit
        # flush against the bottom edge of the scrollable area.
        tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, height=12).pack(fill="x")

        # Hairline seam between the dark sidebar and the light content
        # canvas - crisper than letting the two surfaces butt directly
        # against each other.
        tk.Frame(parent, bg=self.COLOR_SIDEBAR_DIVIDER, width=1).pack(side="left", fill="y")

    def _build_collapsible_menu(self, sidebar, label_text, sub_items, before):
        """Generic collapsible parent menu (e.g. 'General Operations' or
        'Coinsurance'), shown inside a recessed, darker-navy block with a
        left indent guide so its sub-items read as clearly nested under
        the parent. Clicking the header only expands/collapses it (it
        isn't itself a page); clicking a sub-item navigates like any
        other nav item.

        `before`: a permanent fixed sibling widget, packed once and never
        forgotten/re-packed, that this menu's header and submenu are
        always inserted immediately ahead of via pack(before=...). This
        has to be a widget that never moves, otherwise re-expanding a
        submenu after collapsing it could land it in the wrong place
        relative to the other sidebar sections.

        Returns the header frame widget, so an earlier-built menu can use
        it as its own `before` anchor and keep menus stacked in the order
        they were built.

        sub_items entries are normally a plain (key, label) leaf tuple.
        An entry can instead be (None, label, nested_items) to nest a
        FURTHER collapsible submenu one level deeper (e.g. "Claims"
        nested inside "Coinsurance") - see _build_nested_submenu.
        """
        header = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR, cursor="hand2")
        header.pack(fill="x", before=before)

        header_accent = tk.Frame(header, bg=self.COLOR_SIDEBAR, width=4)
        header_accent.pack(side="left", fill="y")

        header_inner = tk.Frame(header, bg=self.COLOR_SIDEBAR)
        header_inner.pack(side="left", fill="both", expand=True, padx=(16, 14), pady=15)

        header_label = tk.Label(header_inner, text=label_text, font=("Segoe UI Semibold", 11, "bold"),
                                 bg=self.COLOR_SIDEBAR, fg=self.COLOR_SIDEBAR_TEXT_ACTIVE, anchor="w")
        header_label.pack(side="left", fill="both", expand=True)

        caret = tk.Label(header_inner, text="\u25BE", font=("Segoe UI", 9),
                          bg=self.COLOR_SIDEBAR, fg=self.COLOR_SIDEBAR_TEXT_DIM)
        caret.pack(side="right")

        # Recessed wrapper block for the sub-items - a visibly darker
        # navy "well" (not just a faint tint) so the nesting relationship
        # to the parent header is obvious at a glance, with a slim gold-
        # tinted guide line down the left edge instead of a barely-there
        # 1px border.
        submenu_block = tk.Frame(sidebar, bg=self.COLOR_SIDEBAR_SUBBG)
        submenu_block.pack(fill="x", before=before)
        guide = tk.Frame(submenu_block, bg=self.COLOR_SIDEBAR_GUIDE, width=2)
        guide.place(x=29, y=2, relheight=1.0, height=-8)

        submenu_frame = tk.Frame(submenu_block, bg=self.COLOR_SIDEBAR_SUBBG)
        submenu_frame.pack(fill="x", pady=(2, 6))  # starts expanded

        for entry in sub_items:
            if entry[0] is None:
                _, nested_label, nested_items = entry
                self._build_nested_submenu(submenu_frame, nested_label, nested_items)
            else:
                key, label = entry
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

        for w in (header, header_accent, header_inner, caret, header_label):
            w.bind("<Button-1>", toggle_menu)
            w.bind("<Enter>", lambda e: self._hover_menu_header(header, header_accent, header_inner, header_label, True))
            w.bind("<Leave>", lambda e: self._hover_menu_header(header, header_accent, header_inner, header_label, False))

        return header

    def _hover_menu_header(self, header, header_accent, header_inner, header_label, entering, base_bg=None):
        base_bg = base_bg or self.COLOR_SIDEBAR
        bg = self.COLOR_SIDEBAR_HOVER if entering else base_bg
        for w in (header, header_inner, header_label):
            w.configure(bg=bg)
        header_accent.configure(bg=bg)

    def _build_nested_submenu(self, parent_container, label_text, nested_items):
        """A second-level collapsible menu nested inside a top-level
        collapsible menu's own recessed block (e.g. 'Claims' nested inside
        'Coinsurance') - same expand/collapse behavior as
        _build_collapsible_menu, one level deeper: its own header sits at
        the same indent as a normal leaf item, and its leaf items sit in a
        further-indented, further-nested block below it.

        Built once, in order, directly into `parent_container` (the
        parent menu's own submenu_frame) - a plain pack() (append at the
        end) is what places this submenu's header/anchor correctly
        relative to whatever else has already been built into
        parent_container. But a plain pack() is NOT safe for re-showing
        inner_block after it's been collapsed (see toggle_menu below) -
        see anchor's own docstring for why.
        """
        header = tk.Frame(parent_container, bg=self.COLOR_SIDEBAR_SUBBG, cursor="hand2")
        header.pack(fill="x")

        header_accent = tk.Frame(header, bg=self.COLOR_SIDEBAR_SUBBG, width=4)
        header_accent.pack(side="left", fill="y")

        header_inner = tk.Frame(header, bg=self.COLOR_SIDEBAR_SUBBG)
        header_inner.pack(side="left", fill="both", expand=True, padx=(24, 14), pady=12)

        header_label = tk.Label(header_inner, text=label_text, font=("Segoe UI", 10, "bold"),
                                 bg=self.COLOR_SIDEBAR_SUBBG, fg=self.COLOR_SIDEBAR_TEXT_ACTIVE, anchor="w")
        header_label.pack(side="left", fill="both", expand=True)

        caret = tk.Label(header_inner, text="\u25BE", font=("Segoe UI", 8),
                          bg=self.COLOR_SIDEBAR_SUBBG, fg=self.COLOR_SIDEBAR_TEXT_DIM)
        caret.pack(side="right")

        # Permanent, zero-height anchor - packed once, right after the
        # header, and never touched again. inner_block is always
        # (re-)packed with before=anchor, so toggle_menu's else-branch
        # (re-showing inner_block after a collapse) restores it to
        # EXACTLY this position - immediately after its own header -
        # regardless of what other nested submenus get built into
        # parent_container afterward (e.g. "Claims" built right after
        # "Premium").
        #
        # CONFIRMED REAL BUG this fixes: without an anchor, toggle_menu's
        # else-branch called inner_block.pack(fill="x") with no before= -
        # a plain pack() call always APPENDS the widget to the end of its
        # parent's current pack order, not back into its original slot.
        # So collapsing then re-expanding "Premium" (built first) moved
        # its own inner_block (Policy Check / Premium Register) to the
        # END of the Coinsurance menu's own submenu_frame - after
        # "Claims" own header AND its inner_block - which is exactly
        # what made Premium's items visually appear below Claims instead
        # of directly under Premium's own header. The very first render
        # (before anyone ever toggled anything) looked correct purely by
        # accident, since the initial pack() calls happen to already run
        # in top-to-bottom order - the bug only ever showed up after an
        # actual collapse/expand interaction.
        anchor = tk.Frame(parent_container, bg=self.COLOR_SIDEBAR_SUBBG, height=1)
        anchor.pack(fill="x")

        inner_block = tk.Frame(parent_container, bg=self.COLOR_SIDEBAR_SUBBG)
        inner_block.pack(fill="x", before=anchor)
        guide = tk.Frame(inner_block, bg=self.COLOR_SIDEBAR_GUIDE, width=2)
        guide.place(x=37, y=2, relheight=1.0, height=-8)

        inner_frame = tk.Frame(inner_block, bg=self.COLOR_SIDEBAR_SUBBG)
        inner_frame.pack(fill="x", pady=(2, 6))  # starts expanded

        for key, label in nested_items:
            self.nav_items[key] = self._make_nav_item(
                inner_frame, key, label, indent=True, base_bg=self.COLOR_SIDEBAR_SUBBG, extra_indent=16
            )

        def toggle_menu(event=None):
            if inner_block.winfo_ismapped():
                inner_block.pack_forget()
                caret.configure(text="\u25B8")
            else:
                inner_block.pack(fill="x", before=anchor)
                caret.configure(text="\u25BE")

        for w in (header, header_accent, header_inner, caret, header_label):
            w.bind("<Button-1>", toggle_menu)
            w.bind("<Enter>", lambda e: self._hover_menu_header(
                header, header_accent, header_inner, header_label, True, base_bg=self.COLOR_SIDEBAR_SUBBG))
            w.bind("<Leave>", lambda e: self._hover_menu_header(
                header, header_accent, header_inner, header_label, False, base_bg=self.COLOR_SIDEBAR_SUBBG))

        return header

    def _make_nav_item(self, parent, key, text, indent=False, before=None, base_bg=None, extra_indent=0):
        base_bg = base_bg or self.COLOR_SIDEBAR
        wrapper = tk.Frame(parent, bg=base_bg, cursor="hand2")
        pack_kwargs = {"fill": "x"}
        if before is not None:
            pack_kwargs["before"] = before
        wrapper.pack(**pack_kwargs)

        accent = tk.Frame(wrapper, bg=base_bg, width=4)
        accent.pack(side="left", fill="y")

        inner = tk.Frame(wrapper, bg=base_bg)
        # extra_indent pushes a leaf item further right for each
        # additional level of nesting below the first (e.g. "Claims
        # Intimation" sitting under the nested "Claims" submenu, rather
        # than directly under "Coinsurance") - 0 for a normal sub-item.
        left_pad = (24 + extra_indent, 14) if indent else (16, 14)
        inner.pack(side="left", fill="both", expand=True,
                    padx=left_pad, pady=13 if indent else 15)

        label = tk.Label(inner, text=text, font=("Segoe UI", 10 if indent else 11),
                          bg=base_bg, fg=self.COLOR_SIDEBAR_TEXT, anchor="w",
                          wraplength=(190 - extra_indent) if indent else None, justify="left")
        label.pack(side="left", fill="both", expand=True)

        def on_click(event=None, k=key):
            self.select_nav(k)

        def on_enter(event=None):
            if self.active_nav != key:
                hover_bg = self.COLOR_SIDEBAR_HOVER
                for w in (wrapper, accent, inner):
                    w.configure(bg=hover_bg)
                label.configure(bg=hover_bg, fg=self.COLOR_SIDEBAR_TEXT_ACTIVE)

        def on_leave(event=None):
            if self.active_nav != key:
                for w in (wrapper, accent, inner):
                    w.configure(bg=base_bg)
                label.configure(bg=base_bg, fg=self.COLOR_SIDEBAR_TEXT)

        for widget in (wrapper, accent, inner, label):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)

        wrapper.accent = accent
        wrapper.inner = inner
        wrapper.label = label
        wrapper.indent = indent
        wrapper.base_bg = base_bg
        return wrapper

    def select_nav(self, key):
        """Switch the visible content page and restyle the sidebar so only
        the active section shows the gold accent bar, a clearly brighter
        "pill" background, and full-white bold text - the three states
        (default / hover / active) are visually distinct enough that the
        current page never has to be guessed at."""
        self.active_nav = key
        for k, item in self.nav_items.items():
            selected = (k == key)
            bg = self.COLOR_SIDEBAR_ACTIVE if selected else item.base_bg
            base_size = 10 if item.indent else 11
            item.configure(bg=bg)
            item.accent.configure(bg=self.COLOR_GOLD if selected else bg)
            item.inner.configure(bg=bg)
            item.label.configure(
                bg=bg,
                fg=self.COLOR_SIDEBAR_TEXT_ACTIVE if selected else self.COLOR_SIDEBAR_TEXT,
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

        upr_automation_page = ttk.Frame(container)
        upr_automation_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["upr_automation"] = upr_automation_page
        self._build_upr_automation_page(upr_automation_page)

        reinsurance_page = ttk.Frame(container)
        reinsurance_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["reinsurance_automation"] = reinsurance_page
        self._build_reinsurance_automation_page(reinsurance_page)

        deferred_commission_page = ttk.Frame(container)
        deferred_commission_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["deferred_commission"] = deferred_commission_page
        self._build_deferred_commission_page(deferred_commission_page)

        pc_email_page = ttk.Frame(container)
        pc_email_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_pc_email"] = pc_email_page
        self._build_receivable_page(pc_email_page)

        premium_register_page = ttk.Frame(container)
        premium_register_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_premium_register"] = premium_register_page
        self._build_premium_register_page(premium_register_page)

        claims_register_page = ttk.Frame(container)
        claims_register_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_claims_register"] = claims_register_page
        self._build_claims_register_page(claims_register_page)

        claims_check_page = ttk.Frame(container)
        claims_check_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_claims_check"] = claims_check_page
        self._build_claims_check_page(claims_check_page)

        premregsplit_page = ttk.Frame(container)
        premregsplit_page.grid(row=0, column=0, sticky="nsew")
        self.nav_pages["coinsurance_premium_register_split"] = premregsplit_page
        self._build_premium_register_split_page(premregsplit_page)

        self.select_nav("outstanding_premium")

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
        ttk.Label(header, text="Loss Ratio", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        ttk.Label(header, text="Calculate and summarize Loss Ratio for one company or a group, from Excel or PDF source files.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

        formula_badge = tk.Frame(header, bg=self.COLOR_ACCENT_SOFT)
        formula_badge.pack(anchor="w")
        tk.Frame(formula_badge, bg=self.COLOR_GOLD, width=3).pack(side="left", fill="y")
        tk.Label(formula_badge, text="Loss Ratio %  =  (Incurred Loss + Commission Accrual)  \u00f7  Gross Premium Written  \u00d7  100",
                 bg=self.COLOR_ACCENT_SOFT, fg=self.COLOR_PRIMARY, font=("Consolas", 10, "bold"),
                 padx=14, pady=8).pack(side="left")

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

        calc_btn = self._make_action_button(s3, "Calculate Loss Ratio", self._calculate_loss_ratio, kind="accent")
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
        own fixed character width. A genuine hover state (not just the
        native click-flash) is wired via Enter/Leave so buttons feel
        responsive under the cursor, matching the sidebar's own hover
        treatment."""
        hover_bg = self.COLOR_ACCENT_SOFT if outline else self._shade(bg, -18)
        btn = tk.Button(
            parent, text=text, command=command,
            font=("Segoe UI", 9, "bold"), bg=bg, fg=(self.COLOR_PRIMARY if outline else fg),
            activebackground=hover_bg,
            activeforeground=(self.COLOR_PRIMARY if outline else fg),
            relief="solid" if outline else "flat",
            borderwidth=1 if outline else 0,
            highlightbackground=self.COLOR_BORDER,
            padx=15, pady=8, cursor="hand2",
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    @staticmethod
    def _shade(hex_color, amount):
        """Lighten (amount > 0) or darken (amount < 0) a #RRGGBB color by
        `amount` per channel - used for lightweight, dependency-free hover
        states instead of hand-picking a second color for every button."""
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return "#" + hex_color
        r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
        r, g, b = (max(0, min(255, c + amount)) for c in (r, g, b))
        return f"#{r:02X}{g:02X}{b:02X}"

    def _make_action_button(self, parent, text, command, kind="accent", compact=False):
        """Shared factory for every primary call-to-action button in the
        app (Calculate Loss Ratio, Run Matching, Save Edits, Send This
        Email, ...) - previously each of these was its own hand-built
        tk.Button with slightly different padding/hover logic. Centralizing
        them here means one visual language (flat fill, generous padding,
        a genuine hover shade, a subtle pressed-in shade on click) across
        every page instead of small inconsistencies page to page.

        kind: 'accent' (sapphire - the default primary action), 'success'
            (green - sending/affirmative actions), or 'quiet' (navy - a
            lower-emphasis action like Save Signature that shouldn't
            visually compete with the page's main CTA).
        compact: smaller padding for a button that sits inline next to a
            text field (e.g. Save Signature) rather than anchoring a
            section on its own.
        """
        bg, fg = {
            "accent": (self.COLOR_ACCENT, "white"),
            "success": (self.COLOR_SUCCESS, "white"),
            "quiet": (self.COLOR_PRIMARY, "white"),
        }.get(kind, (self.COLOR_ACCENT, "white"))
        hover_bg = self._shade(bg, -18)
        press_bg = self._shade(bg, -32)
        padx, pady, font_size = (12, 6, 9) if compact else (20, 11, 10)

        btn = tk.Button(
            parent, text=text, command=command,
            font=("Segoe UI", font_size, "bold"), bg=bg, fg=fg,
            activebackground=press_bg, activeforeground=fg,
            relief="flat", borderwidth=0, cursor="hand2",
            padx=padx, pady=pady,
            highlightthickness=0,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        btn.bind("<ButtonPress-1>", lambda e: btn.configure(bg=press_bg))
        btn.bind("<ButtonRelease-1>", lambda e: btn.configure(bg=hover_bg))
        return btn

    def _loss_section_card(self, parent, number, title, pad_x, expand_bottom=False):
        """A numbered section card matching the app's '1. Upload & Filter'
        style numbering convention, for visual consistency across pages.
        Rendered with a soft drop shadow (a slightly darker frame peeking
        out from behind the panel) and a slim gold top rule, so each card
        reads as a distinct, lightly elevated surface rather than a flat
        bordered box - the same "quiet elegance" cue used in the header."""
        shadow = tk.Frame(parent, bg="#D6DDE7")
        shadow.pack(fill="both" if expand_bottom else "x", expand=expand_bottom,
                     padx=pad_x, pady=8)

        wrapper = tk.Frame(shadow, bg=self.COLOR_PANEL, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        wrapper.pack(fill="both", expand=True, padx=(0, 3), pady=(0, 3))

        tk.Frame(wrapper, bg=self.COLOR_GOLD, height=2).pack(fill="x", side="top")

        title_row = tk.Frame(wrapper, bg=self.COLOR_PANEL)
        title_row.pack(fill="x", padx=18, pady=(15, 10))

        badge = tk.Canvas(title_row, width=24, height=24, bg=self.COLOR_PANEL, highlightthickness=0)
        badge.pack(side="left", padx=(0, 10))
        badge.create_oval(1, 1, 23, 23, fill=self.COLOR_GOLD_DEEP, outline="")
        badge.create_text(12, 12, text=number, fill="white", font=("Segoe UI Semibold", 10, "bold"))

        tk.Label(title_row, text=title, bg=self.COLOR_PANEL, fg=self.COLOR_PRIMARY,
                 font=("Segoe UI Semibold", 12, "bold")).pack(side="left")

        content = tk.Frame(wrapper, bg=self.COLOR_PANEL)
        content.pack(fill="both" if expand_bottom else "x", expand=expand_bottom, padx=18, pady=(0, 18))
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
        # Default save location is a fresh dated/timed subfolder under this
        # module's own base folder, same "module folder -> one subfolder
        # per generation" pattern used everywhere else in the app - the
        # person can still browse elsewhere in the dialog if they prefer.
        default_dir = self._new_run_folder(os.path.join(os.getcwd(), "Loss Ratio"))
        output_path = filedialog.asksaveasfilename(
            title="Save Loss Ratio Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=default_dir,
            initialfile=default_name,
        )
        self._cleanup_if_empty_dir(default_dir)  # remove the seeded folder if nothing was ever written into it
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

    # ---------- UPR (Unearned Premium Revenue) Automation module ----------
    def _build_upr_automation_page(self, parent):
        """UPR Automation: Section 1 manual entries (Premium Direct /
        Coinsurance / Reinsurance / Admin Surcharge, Opening UPR), a
        Premium Register upload feeding a period dropdown, and the
        two-table (Gross Premium Distribution / Unearned Premium
        Working) calculation described in upr_processor.py - laid out as
        a 3-tab Notebook, same pattern used across GIMS's other modules."""
        header = ttk.Frame(parent)
        header.pack(fill="x", padx=20, pady=(20, 6))
        ttk.Label(header, text="UPR Automation", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        upr_subtitle = ttk.Label(
            header, text="Computes Closing Unearned Premium Revenue from the Premium Register using the "
                          "1/24 (most departments) and 1/6 (Marine/Travel) amortization methods, for a "
                          "selected 12-month period.",
            foreground=self.COLOR_MUTED, font=("Segoe UI", 10), justify="left")
        upr_subtitle.pack(anchor="w", fill="x", pady=(2, 10))

        def _wrap_upr_subtitle(event):
            upr_subtitle.configure(wraplength=max(200, event.width - 4))
        header.bind("<Configure>", _wrap_upr_subtitle)

        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        tab_register = ttk.Frame(notebook)
        tab_section1 = ttk.Frame(notebook)
        tab_generate = ttk.Frame(notebook)
        notebook.add(tab_register, text="1. Premium Register & Period")
        notebook.add(tab_section1, text="2. Section 1 Entries & Calculate")
        notebook.add(tab_generate, text="3. Generate & Save")

        self._build_upr_tab_register(tab_register)
        self._build_upr_tab_section1(tab_section1)
        self._build_upr_tab_generate(tab_generate)

    def _make_upr_scrollable_body(self, tab):
        outer = ttk.Frame(tab)
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

        # Stashed so _finalize_scroll_passthrough (called once, after all
        # of this tab's own widgets are built) can extend the same
        # Enter/Leave pair down onto every descendant - see that method's
        # own docstring for why the canvas-only binding above isn't
        # enough on a tab with several stacked Entry fields / a results
        # Treeview, both of which this module has plenty of.
        body._scroll_canvas = canvas
        body._scroll_handler = _on_mousewheel
        return body

    def _finalize_scroll_passthrough(self, body, skip_types=None):
        """
        Extends `body`'s own canvas-level mousewheel Enter/Leave binding
        (see _make_upr_scrollable_body / _enable_mousewheel) down onto
        every descendant widget, so scrolling keeps working no matter
        where inside the tab the cursor happens to be - not just over
        the bare background between widgets.

        WHY THIS IS NEEDED: Tk fires <Enter>/<Leave> on every real
        window-boundary crossing, and each child widget embedded in the
        canvas (a Label, Entry, Button, Frame...) is its own separate
        window. Moving the cursor from the canvas's own background onto
        ANY child fires <Leave> on the canvas - which tears down the
        page's global mousewheel binding - and nothing re-establishes it,
        since the child itself has no binding of its own. On a page with
        just a couple of small buttons this is barely noticeable; on a
        tab with several stacked Entry fields (this module's Section 1
        form) or a wide static Label, large parts of the visible area
        end up "dead" for scrolling - not something a person would think
        to work around by nudging the mouse a few pixels first.

        Call this ONCE, at the end of a tab-builder method, after every
        widget for that tab has already been created (binding has to
        walk the real widget tree, so it can't run before the tree
        exists). Not needed again afterward unless more widgets are
        added to that same body later.

        skip_types: widget classes deliberately left out of the
        pass-through, because they already scroll their OWN content on
        the wheel via their built-in class binding (Listbox, Treeview,
        Text/ScrolledText) - adding our page-level handler on top of
        that would fight with their native behavior instead of just
        letting them handle their own scrolling, which is what a person
        actually wants when hovering directly over one of them (e.g. the
        register file list, or the department results table).
        """
        canvas = getattr(body, "_scroll_canvas", None)
        handler = getattr(body, "_scroll_handler", None)
        if canvas is None or handler is None:
            return
        skip_types = skip_types or (tk.Listbox, ttk.Treeview, tk.Text)

        def _bind(widget):
            if not isinstance(widget, skip_types):
                widget.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", handler), add="+")
                widget.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"), add="+")
            for child in widget.winfo_children():
                _bind(child)

        for child in body.winfo_children():
            _bind(child)

    # ---- Tab 1: Premium Register upload + period selection ----
    def _build_upr_tab_register(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s1 = self._loss_section_card(body, "1", "Premium Register File(s)", PAD_X)
        upr_s1_note = ttk.Label(
            s1, text="Upload one consolidated file (with a sheet per year) or several separate per-year "
                      "files - either works, every qualifying sheet across every file is combined before "
                      "anything is calculated. Required columns: Department/Class, Comm Date, Gross "
                      "Premium/Contribution.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        upr_s1_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(upr_s1_note, s1, side_padding=10)

        list_card = tk.Frame(s1, bg=self.COLOR_PANEL_ALT, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_card.pack(fill="x")
        list_outer = tk.Frame(list_card, bg=self.COLOR_PANEL_ALT)
        list_outer.pack(fill="x", padx=12, pady=10)

        list_frame = tk.Frame(list_outer, bg="white", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_frame.pack(fill="x")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.upr_register_listbox = tk.Listbox(list_frame, height=5, font=("Segoe UI", 10), selectmode="extended",
                                                 activestyle="none", relief="flat", borderwidth=0,
                                                 bg="white", selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                                 highlightthickness=0, yscrollcommand=list_scroll.set)
        list_scroll.configure(command=self.upr_register_listbox.yview)
        self.upr_register_listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        toolbar = tk.Frame(list_outer, bg=self.COLOR_PANEL_ALT)
        toolbar.pack(fill="x", pady=(8, 0))
        self._make_pill_button(toolbar, "Browse...", self._add_upr_register_files, bg=self.COLOR_ACCENT).pack(side="left")
        self._make_pill_button(toolbar, "Remove Selected", self._remove_upr_register_file,
                                bg=self.COLOR_PANEL, fg=self.COLOR_TEXT, outline=True).pack(side="left", padx=(8, 0))
        self._make_pill_button(toolbar, "Clear All", self._clear_upr_register_files,
                                bg=self.COLOR_PANEL, fg=self.COLOR_WARN, outline=True).pack(side="left", padx=(8, 0))

        load_row = ttk.Frame(s1)
        load_row.pack(fill="x", pady=(10, 0))
        self.upr_load_btn = self._make_action_button(load_row, "Load Register(s)", self._load_upr_registers, kind="accent")
        self.upr_load_btn.pack(side="left")

        self.upr_load_progress = perf_utils.ProgressBarPanel(s1)
        self.upr_load_progress.frame.pack(fill="x", pady=(8, 0))

        s2 = self._loss_section_card(body, "2", "Period", PAD_X, expand_bottom=True)
        upr_s2_note = ttk.Label(
            s2, text="Select the period END month - the calculation covers exactly the 12 months ending "
                      "there (e.g. selecting \"June 2026\" covers July 2025 through June 2026). If the "
                      "uploaded register(s) don't reach back or forward far enough to cover the full 12 "
                      "months, an error is raised rather than guessing.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        upr_s2_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(upr_s2_note, s2, side_padding=10)

        period_row = ttk.Frame(s2)
        period_row.pack(fill="x")
        ttk.Label(period_row, text="Period (ending month)", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.upr_period_var = tk.StringVar()
        self.upr_period_combo = ttk.Combobox(period_row, textvariable=self.upr_period_var, state="readonly", width=22)
        self.upr_period_combo.pack(side="left")

        self.upr_register_log = scrolledtext.ScrolledText(s2, height=8, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.upr_register_log.pack(fill="both", expand=True, pady=(12, 0))

        self._finalize_scroll_passthrough(body)

    def _log_upr(self, msg):
        self.upr_register_log.insert("end", str(msg) + "\n")
        self.upr_register_log.see("end")

    def _add_upr_register_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        added = 0
        for p in paths:
            if p not in self.upr_register_paths:
                self.upr_register_paths.append(p)
                self.upr_register_listbox.insert("end", os.path.basename(p))
                added += 1
        if added:
            self.status_var.set(f"UPR Automation: added {added} file(s) ({len(self.upr_register_paths)} total).")

    def _remove_upr_register_file(self):
        selected = list(self.upr_register_listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            self.upr_register_listbox.delete(idx)
            del self.upr_register_paths[idx]

    def _clear_upr_register_files(self):
        self.upr_register_listbox.delete(0, "end")
        self.upr_register_paths.clear()

    def _load_upr_registers(self):
        if not self.upr_register_paths:
            messagebox.showerror("No files selected", "Browse to at least one Premium Register file first.")
            return

        self.upr_register_log.delete("1.0", "end")
        self.status_var.set("Loading premium register(s)...")
        self.upr_load_progress.start(total=None, text="Loading...")
        paths = list(self.upr_register_paths)

        def worker(report):
            df, meta = uprp.load_premium_registers(paths, report=report)
            return df, meta

        def on_done(result):
            df, meta = result
            self.upr_loaded_df = df
            self.upr_load_progress.finish(f"Loaded {len(df):,} row(s).")
            self._log_upr(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s): "
                          + ", ".join(f"{fn} [{sn}]" for fn, sn in meta["sheets_used"]))
            if meta["sheets_skipped"]:
                self._log_upr(f"Skipped {len(meta['sheets_skipped'])} sheet(s) that didn't look like a premium register:")
                for fn, sn, reason in meta["sheets_skipped"]:
                    self._log_upr(f"   {fn} [{sn}]: {reason}")
            if meta["excluded_rows"]:
                self._log_upr(f"{meta['excluded_rows']:,} row(s) excluded (missing Department, Comm Date, "
                              f"Gross Premium, or an implausibly OLD Comm Date).")
            if meta.get("implausible_date_rows"):
                self._log_upr(f"  Of those, {meta['implausible_date_rows']:,} had a Comm Date implausibly far in "
                              f"the past - excluded, since a corrupted date landing decades/centuries too early "
                              f"has no sensible bucket to land in. This is a source file data issue, worth fixing "
                              f"at the export level. Example(s):")
                for s in meta["implausible_date_samples"][:5]:
                    cd = s["comm_date"]
                    cd_text = cd.strftime("%d-%b-%Y") if hasattr(cd, "strftime") else str(cd)
                    doc_ref_text = f" ({s['doc_ref']})" if s.get("doc_ref") else ""
                    self._log_upr(f"    {s['department']}: {cd_text}{doc_ref_text}")
            if meta.get("far_future_rows"):
                self._log_upr(f"{meta['far_future_rows']:,} row(s) had a Comm Date implausibly far in the future "
                              f"(e.g. year 5026) - possibly a corrupted date value in the source export, but NOT "
                              f"excluded: they're kept and counted under \"Future Policies\" (fully unearned) "
                              f"wherever the selected period, no matter how far out the date is. They're also kept "
                              f"out of the Period dropdown itself, so a corrupted year never shows up there as a "
                              f"selectable period. Example(s):")
                for s in meta["far_future_date_samples"][:5]:
                    cd = s["comm_date"]
                    cd_text = cd.strftime("%d-%b-%Y") if hasattr(cd, "strftime") else str(cd)
                    doc_ref_text = f" ({s['doc_ref']})" if s.get("doc_ref") else ""
                    self._log_upr(f"    {s['department']}: {cd_text}{doc_ref_text}")

            options = uprp.get_available_period_options(df)
            self.upr_period_options = options
            self.upr_period_combo.configure(values=[o["label"] for o in options])
            if options:
                self.upr_period_combo.current(0)
            self.status_var.set(f"UPR Automation: {len(df):,} row(s) loaded. Choose a period, then go to tab 2.")

        def on_error(exc):
            self.upr_load_progress.finish("Failed - see the log for details.")
            self._log_upr(f"ERROR: {exc}")
            messagebox.showerror("Could not load register(s)", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_upr(msg),
            on_progress_value=self.upr_load_progress.update,
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.upr_load_btn],
        )

    # ---- Tab 2: Section 1 manual entries + Calculate ----
    def _build_upr_tab_section1(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s3 = self._loss_section_card(body, "3", "Section 1 - Premium Income Figures", PAD_X)
        upr_s3_note = ttk.Label(
            s3, text="Premium Direct / Coinsurance / Reinsurance / Administrative Surcharge can be loaded "
                      "automatically from the \"All Ledgers Balance\" PDF (Report ID GL_LB_LEDGERBALANCE) "
                      "below, or entered/edited by hand either way. Opening Unearned Premium Revenue isn't "
                      "on that report - it comes from the prior period's own Closing UPR, entered manually "
                      "for now until that's automated too.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        upr_s3_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(upr_s3_note, s3, side_padding=10)

        # ---- Optional: load the 4 premium-income figures from the Ledger Balance PDF ----
        pdf_card = tk.Frame(s3, bg=self.COLOR_PANEL_ALT, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        pdf_card.pack(fill="x", pady=(0, 14))
        pdf_inner = tk.Frame(pdf_card, bg=self.COLOR_PANEL_ALT)
        pdf_inner.pack(fill="x", padx=12, pady=10)

        pdf_row = tk.Frame(pdf_inner, bg=self.COLOR_PANEL_ALT)
        pdf_row.pack(fill="x")
        tk.Label(pdf_row, text="Ledger Balance PDF (optional)", font=("Segoe UI", 10, "bold"),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_TEXT, width=24, anchor="w").pack(side="left")
        self.upr_ledger_pdf_name_var = tk.StringVar(value="No file selected")
        tk.Label(pdf_row, textvariable=self.upr_ledger_pdf_name_var, font=("Segoe UI", 9), bg=self.COLOR_PANEL_ALT,
                 fg=self.COLOR_MUTED, anchor="w").pack(side="left", fill="x", expand=True, padx=(6, 10))
        self._make_pill_button(pdf_row, "Browse...", self._browse_upr_ledger_pdf, bg=self.COLOR_ACCENT).pack(side="left", padx=(0, 8))
        self._make_pill_button(pdf_row, "Load from PDF", self._load_upr_ledger_pdf,
                                bg=self.COLOR_PANEL, fg=self.COLOR_TEXT, outline=True).pack(side="left")

        self.upr_ledger_status_var = tk.StringVar(value="")
        tk.Label(pdf_inner, textvariable=self.upr_ledger_status_var, font=("Segoe UI", 9), bg=self.COLOR_PANEL_ALT,
                 fg=self.COLOR_MUTED, anchor="w", justify="left", wraplength=760).pack(anchor="w", pady=(8, 0))

        self.upr_section1_vars = {}
        fields = [
            ("premium_direct", "Premium Direct"),
            ("premium_coinsurance", "Premium Income Coinsurance"),
            ("premium_reinsurance", "Premium Income Reinsurance"),
            ("admin_surcharge", "Administrative Surcharge"),
            ("opening_upr", "Opening Unearned Premium Revenue"),
        ]
        form = ttk.Frame(s3)
        form.pack(fill="x")
        for i, (key, label) in enumerate(fields):
            ttk.Label(form, text=label, style="Panel.TLabel").grid(row=i, column=0, sticky="w", padx=(0, 10), pady=5)
            var = tk.StringVar()
            ttk.Entry(form, textvariable=var, width=22, font=("Segoe UI", 10)).grid(row=i, column=1, sticky="w", pady=5)
            self.upr_section1_vars[key] = var
        ttk.Label(form, text="(values above can still be typed/edited by hand at any point)",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 8, "italic"))\
            .grid(row=len(fields), column=0, columnspan=2, sticky="w", pady=(4, 0))

        s4 = self._loss_section_card(body, "4", "Run Calculation", PAD_X, expand_bottom=True)
        ttk.Label(s4, text="Builds both tables (Gross Premium Distribution and Unearned Premium Working) "
                            "for the period selected on tab 1.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        calc_btn = self._make_action_button(s4, "Calculate", self._run_upr_calculation, kind="accent")
        calc_btn.pack(anchor="w")

        self.upr_calc_progress = perf_utils.ProgressBarPanel(s4)
        self.upr_calc_progress.frame.pack(fill="x", pady=(10, 0))

        self.upr_calc_summary_var = tk.StringVar(value="")
        ttk.Label(s4, textvariable=self.upr_calc_summary_var, foreground=self.COLOR_SUCCESS,
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 8))

        tree_frame = ttk.Frame(s4)
        tree_frame.pack(fill="both", expand=True)
        columns = ("department", "method", "gross_total", "unearned_total")
        self.upr_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)
        self.upr_results_tree.heading("department", text="Department")
        self.upr_results_tree.heading("method", text="Method")
        self.upr_results_tree.heading("gross_total", text="Gross Premium Total")
        self.upr_results_tree.heading("unearned_total", text="Unearned Total")
        self.upr_results_tree.column("department", width=180)
        self.upr_results_tree.column("method", width=80, anchor="center")
        self.upr_results_tree.column("gross_total", width=170, anchor="e")
        self.upr_results_tree.column("unearned_total", width=170, anchor="e")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.upr_results_tree.yview)
        self.upr_results_tree.configure(yscrollcommand=vsb.set)
        self.upr_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._finalize_scroll_passthrough(body)

    def _browse_upr_ledger_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.upr_ledger_pdf_path = path
            self.upr_ledger_pdf_name_var.set(os.path.basename(path))

    def _load_upr_ledger_pdf(self):
        if not self.upr_ledger_pdf_path:
            messagebox.showerror("No file selected", "Browse to the Ledger Balance PDF first.")
            return
        try:
            result = uprlp.parse_ledger_balance_pdf(self.upr_ledger_pdf_path)
        except uprlp.UPRLedgerError as e:
            self.upr_ledger_status_var.set(f"Could not load: {e}")
            messagebox.showerror("Could not read PDF", str(e))
            return

        self.upr_ledger_parse_result = result
        matched_labels = []
        for field, label in uprlp.FIELD_LABELS.items():
            if result[field] is not None:
                self.upr_section1_vars[field].set(f"{result[field]:,.2f}")
                matched_labels.append(label)
            else:
                matched_labels.append(f"{label} (not found - left as-is)")

        status_lines = [f"Loaded: {', '.join(matched_labels)}."]

        if result["unmatched_lines"]:
            names = ", ".join(f"{u['description']} ({u['amount']:,.0f})" for u in result["unmatched_lines"][:5])
            status_lines.append(f"Note: {len(result['unmatched_lines'])} other account line(s) on the report "
                                 f"weren't one of the four expected accounts and were ignored: {names}.")

        if result["grand_total_reported"] is not None:
            diff = result["grand_total_matched"] - result["grand_total_reported"]
            if abs(diff) > 1.0:
                status_lines.append(f"WARNING: the report's own Grand Total ({result['grand_total_reported']:,.2f}) "
                                     f"doesn't match the sum of the four matched accounts "
                                     f"({result['grand_total_matched']:,.2f}) - difference {diff:,.2f}. This usually "
                                     f"means the report covered other GL accounts too - double-check before relying "
                                     f"on these figures.")

        if result["as_at_date"] and self.upr_period_options:
            selected_label = self.upr_period_var.get()
            option = next((o for o in self.upr_period_options if o["label"] == selected_label), None)
            if option and (result["as_at_date"].year, result["as_at_date"].month) != (option["year"], option["month"]):
                status_lines.append(f"WARNING: this PDF is as at {result['as_at_date']:%d-%b-%Y}, but the period "
                                     f"selected on tab 1 is \"{selected_label}\" - these don't match. Confirm "
                                     f"you've uploaded the right PDF for this period before generating.")

        self.upr_ledger_status_var.set("\n".join(status_lines))
        self.status_var.set("UPR Automation: Section 1 figures loaded from the Ledger Balance PDF.")

    def _collect_upr_section1_inputs(self):
        out = {}
        for key, var in self.upr_section1_vars.items():
            text = var.get().strip().replace(",", "")
            if not text:
                out[key] = 0.0
                continue
            try:
                out[key] = float(text)
            except ValueError:
                raise ValueError(f"'{text}' isn't a valid number for that field.")
        return out

    def _run_upr_calculation(self):
        if self.upr_loaded_df is None:
            messagebox.showerror("Error", "Load the Premium Register(s) on tab 1 first.")
            return
        if not self.upr_period_options:
            messagebox.showerror("Error", "No period options available - load a register with usable Comm Date values first.")
            return
        selected_label = self.upr_period_var.get()
        option = next((o for o in self.upr_period_options if o["label"] == selected_label), None)
        if option is None:
            messagebox.showerror("Error", "Select a period on tab 1 first.")
            return

        try:
            self._collect_upr_section1_inputs()
        except ValueError as e:
            messagebox.showerror("Invalid entry", str(e))
            return

        df = self.upr_loaded_df
        period_year, period_month = option["year"], option["month"]
        self.upr_calc_summary_var.set("")
        self.upr_calc_progress.start(total=None, text="Calculating...")

        def worker(report):
            period_start, period_end = uprp.compute_period_bounds(period_year, period_month)
            report(f"Selected period: {period_start:%d-%b-%Y} to {period_end:%d-%b-%Y}.")
            uprp.validate_coverage(df, period_start, period_end)
            bucketed = uprp.assign_buckets(df, period_start, period_end)
            departments, warnings = uprp.resolve_department_list(bucketed)
            gross_table = uprp.build_gross_premium_table(bucketed, departments)
            unearned_table, methods = uprp.build_unearned_table(gross_table, departments, period_start)
            closing_upr = uprp.grand_total(unearned_table, departments)
            return {
                "period_start": period_start, "period_end": period_end, "departments": departments,
                "methods": methods, "gross_table": gross_table, "unearned_table": unearned_table,
                "closing_upr": closing_upr, "warnings": warnings,
                "excluded_rows": 0, "sheets_used": [],
            }

        def on_done(result):
            self.upr_last_result = result
            self.upr_calc_progress.finish("Done.")
            self.upr_calc_summary_var.set(f"Closing Unearned Premium Revenue: PKR {result['closing_upr']:,.0f}")

            self.upr_results_tree.delete(*self.upr_results_tree.get_children())
            for dept in result["departments"]:
                gt = uprp.row_total(result["gross_table"][dept])
                ut = uprp.row_total(result["unearned_table"][dept])
                self.upr_results_tree.insert("", "end", values=(
                    dept, result["methods"][dept], f"{gt:,.0f}", f"{ut:,.0f}"
                ))
            for w in result["warnings"]:
                self._log_upr(f"WARNING: {w}")
            self.status_var.set(f"UPR Automation: Closing UPR = PKR {result['closing_upr']:,.0f}. See tab 3 to generate.")

        def on_error(exc):
            self.upr_calc_progress.finish("Failed - see the message for details.")
            messagebox.showerror("Calculation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.upr_calc_progress.update,
            on_done=on_done, on_error=on_error,
        )

    # ---- Tab 3: Generate & Save ----
    def _build_upr_tab_generate(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s5 = self._loss_section_card(body, "5", "Generate & Save", PAD_X, expand_bottom=True)
        out_row = ttk.Frame(s5)
        out_row.pack(fill="x", pady=(0, 8))
        ttk.Label(out_row, text="Output folder", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.upr_output_dir_var = tk.StringVar(value=self.upr_output_dir)
        ttk.Entry(out_row, textvariable=self.upr_output_dir_var, width=50).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Browse...", command=self._browse_upr_output_dir).pack(side="left")

        gen_row = ttk.Frame(s5)
        gen_row.pack(fill="x", pady=(4, 8))
        self._make_action_button(gen_row, "Generate UPR Workbook", self._generate_upr_workbook, kind="accent").pack(side="left")
        ttk.Button(gen_row, text="Open Output Folder", command=self._open_upr_output_folder).pack(side="left", padx=(8, 0))

        self.upr_gen_progress = perf_utils.ProgressBarPanel(s5)
        self.upr_gen_progress.frame.pack(fill="x", pady=(4, 8))

        self.upr_gen_log = scrolledtext.ScrolledText(s5, height=10, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.upr_gen_log.pack(fill="both", expand=True, pady=(8, 0))

        self._finalize_scroll_passthrough(body)

    def _log_upr_gen(self, msg):
        self.upr_gen_log.insert("end", str(msg) + "\n")
        self.upr_gen_log.see("end")

    def _browse_upr_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.upr_output_dir_var.set(path)

    def _generate_upr_workbook(self):
        if self.upr_last_result is None:
            messagebox.showerror("Nothing to generate", "Run Calculate on tab 2 first.")
            return
        base_output_dir = self.upr_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder first.")
            return

        try:
            section1 = self._collect_upr_section1_inputs()
        except ValueError as e:
            messagebox.showerror("Invalid entry", str(e))
            return

        output_dir = self._new_run_folder(base_output_dir)
        result = self.upr_last_result
        filename = uprg.build_output_filename(result["period_end"])
        output_file = os.path.join(output_dir, filename)

        self.upr_gen_log.delete("1.0", "end")
        self.status_var.set("Generating UPR workbook...")
        self.upr_gen_progress.start(total=None, text="Generating...")

        def worker(report):
            return uprg.build_upr_workbook(result, section1, output_file)

        def on_done(fp):
            self.upr_last_output_file = fp
            self.upr_gen_progress.finish("Done.")
            self._log_upr_gen(f"Generated: {fp}")
            messagebox.showinfo("Done", f"UPR workbook generated:\n{fp}")
            self.status_var.set(f"UPR Automation: generated {os.path.basename(fp)}.")

        def on_error(exc):
            self.upr_gen_progress.finish("Failed - see the log for details.")
            self._log_upr_gen(f"ERROR: {exc}")
            messagebox.showerror("Generation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.upr_gen_progress.update,
            on_done=on_done, on_error=on_error,
        )

    def _open_upr_output_folder(self):
        output_dir = os.path.dirname(self.upr_last_output_file) if self.upr_last_output_file else self.upr_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate the workbook first.")

    # ---------- Reinsurance Automation module ----------
    def _build_reinsurance_automation_page(self, parent):
        """Reinsurance Automation: upload one register file per
        department (filename determines department), select a period,
        enter the Opening Prepaid Reinsurance Ceded balance, and generate
        the Prepaid RI Premium Ceded working workbook - same 3-tab shape
        as UPR Automation, since the underlying calculation reuses UPR's
        own period/bucketing/amortization logic directly."""
        header = ttk.Frame(parent)
        header.pack(fill="x", padx=20, pady=(20, 6))
        ttk.Label(header, text="Reinsurance Automation", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        reins_subtitle = ttk.Label(
            header, text="Computes Total Reinsurance Ceded (Treaty + Local + Foreign Premium) per policy from "
                          "each department's own RI Premium Ceded register (one file per department - the "
                          "department is read from the file name), then amortizes it using the same 1/24 "
                          "(most departments) / 1/6 (Marine, Travel) method as UPR Automation to produce the "
                          "Closing Prepaid Reinsurance Premium Ceded figure for a selected 12-month period.",
            foreground=self.COLOR_MUTED, font=("Segoe UI", 10), justify="left")
        reins_subtitle.pack(anchor="w", fill="x", pady=(2, 10))

        def _wrap_reins_subtitle(event):
            reins_subtitle.configure(wraplength=max(200, event.width - 4))
        header.bind("<Configure>", _wrap_reins_subtitle)

        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        tab_register = ttk.Frame(notebook)
        tab_opening = ttk.Frame(notebook)
        tab_generate = ttk.Frame(notebook)
        notebook.add(tab_register, text="1. Register Files & Period")
        notebook.add(tab_opening, text="2. Opening Balance & Calculate")
        notebook.add(tab_generate, text="3. Generate & Save")

        self._build_reins_tab_register(tab_register)
        self._build_reins_tab_opening(tab_opening)
        self._build_reins_tab_generate(tab_generate)

    # ---- Tab 1: register files + period ----
    def _build_reins_tab_register(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s1 = self._loss_section_card(body, "1", "RI Premium Ceded Register File(s)", PAD_X)
        reins_s1_note = ttk.Label(
            s1, text="Upload one file per department (e.g. \"RI Fire 2026.xlsx\", \"Eng__2025_v2.xls\") - "
                      "there is no Department column in these registers at all, so the department is "
                      "determined from each file's own name. Upload as many years/files per department as "
                      "needed; they're combined automatically. Required columns: Comm Date, Treaty "
                      "Premiumium, Local Premium, Foreign Premium.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        reins_s1_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(reins_s1_note, s1, side_padding=10)

        list_card = tk.Frame(s1, bg=self.COLOR_PANEL_ALT, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_card.pack(fill="x")
        list_outer = tk.Frame(list_card, bg=self.COLOR_PANEL_ALT)
        list_outer.pack(fill="x", padx=12, pady=10)

        list_frame = tk.Frame(list_outer, bg="white", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_frame.pack(fill="x")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.reins_register_listbox = tk.Listbox(list_frame, height=5, font=("Segoe UI", 10), selectmode="extended",
                                                   activestyle="none", relief="flat", borderwidth=0,
                                                   bg="white", selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                                   highlightthickness=0, yscrollcommand=list_scroll.set)
        list_scroll.configure(command=self.reins_register_listbox.yview)
        self.reins_register_listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        toolbar = tk.Frame(list_outer, bg=self.COLOR_PANEL_ALT)
        toolbar.pack(fill="x", pady=(8, 0))
        self._make_pill_button(toolbar, "Browse...", self._add_reins_register_files, bg=self.COLOR_ACCENT).pack(side="left")
        self._make_pill_button(toolbar, "Remove Selected", self._remove_reins_register_file,
                                bg=self.COLOR_PANEL, fg=self.COLOR_TEXT, outline=True).pack(side="left", padx=(8, 0))
        self._make_pill_button(toolbar, "Clear All", self._clear_reins_register_files,
                                bg=self.COLOR_PANEL, fg=self.COLOR_WARN, outline=True).pack(side="left", padx=(8, 0))

        load_row = ttk.Frame(s1)
        load_row.pack(fill="x", pady=(10, 0))
        self.reins_load_btn = self._make_action_button(load_row, "Load Register(s)", self._load_reins_registers, kind="accent")
        self.reins_load_btn.pack(side="left")

        self.reins_load_progress = perf_utils.ProgressBarPanel(s1)
        self.reins_load_progress.frame.pack(fill="x", pady=(8, 0))

        s2 = self._loss_section_card(body, "2", "Period", PAD_X, expand_bottom=True)
        reins_s2_note = ttk.Label(
            s2, text="Select the period END month - the calculation covers exactly the 12 months ending "
                      "there. If the uploaded register(s) don't reach back or forward far enough to cover "
                      "the full 12 months, an error is raised rather than guessing.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        reins_s2_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(reins_s2_note, s2, side_padding=10)

        period_row = ttk.Frame(s2)
        period_row.pack(fill="x")
        ttk.Label(period_row, text="Period (ending month)", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.reins_period_var = tk.StringVar()
        self.reins_period_combo = ttk.Combobox(period_row, textvariable=self.reins_period_var, state="readonly", width=22)
        self.reins_period_combo.pack(side="left")

        self.reins_register_log = scrolledtext.ScrolledText(s2, height=8, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.reins_register_log.pack(fill="both", expand=True, pady=(12, 0))

        self._finalize_scroll_passthrough(body)

    def _log_reins(self, msg):
        self.reins_register_log.insert("end", str(msg) + "\n")
        self.reins_register_log.see("end")

    def _add_reins_register_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        added = 0
        for p in paths:
            if p not in self.reins_register_paths:
                self.reins_register_paths.append(p)
                self.reins_register_listbox.insert("end", os.path.basename(p))
                added += 1
        if added:
            self.status_var.set(f"Reinsurance Automation: added {added} file(s) ({len(self.reins_register_paths)} total).")

    def _remove_reins_register_file(self):
        selected = list(self.reins_register_listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            self.reins_register_listbox.delete(idx)
            del self.reins_register_paths[idx]

    def _clear_reins_register_files(self):
        self.reins_register_listbox.delete(0, "end")
        self.reins_register_paths.clear()

    def _load_reins_registers(self):
        if not self.reins_register_paths:
            messagebox.showerror("No files selected", "Browse to at least one RI Premium Ceded register file first.")
            return

        self.reins_register_log.delete("1.0", "end")
        self.status_var.set("Loading RI premium ceded register(s)...")
        self.reins_load_progress.start(total=None, text="Loading...")
        paths = list(self.reins_register_paths)

        def worker(report):
            df, meta = rip.load_reinsurance_registers(paths, report=report)
            return df, meta

        def on_done(result):
            df, meta = result
            self.reins_loaded_df = df
            self.reins_load_progress.finish(f"Loaded {len(df):,} row(s).")
            self._log_reins(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s):")
            for fn, dept in meta["departments_by_file"].items():
                self._log_reins(f"   {fn}  ->  {dept}")
            if meta["excluded_rows"]:
                self._log_reins(f"{meta['excluded_rows']:,} row(s) excluded (missing or implausible Comm Date).")

            options = rip.get_available_period_options(df)
            self.reins_period_options = options
            self.reins_period_combo.configure(values=[o["label"] for o in options])
            if options:
                self.reins_period_combo.current(0)
            self.status_var.set(f"Reinsurance Automation: {len(df):,} row(s) loaded. Choose a period, then go to tab 2.")

        def on_error(exc):
            self.reins_load_progress.finish("Failed - see the log for details.")
            self._log_reins(f"ERROR: {exc}")
            messagebox.showerror("Could not load register(s)", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_reins(msg),
            on_progress_value=self.reins_load_progress.update,
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.reins_load_btn],
        )

    # ---- Tab 2: Opening balance + Calculate ----
    def _build_reins_tab_opening(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s3 = self._loss_section_card(body, "3", "Opening Balance", PAD_X)
        ttk.Label(s3, text="Opening Prepaid Reinsurance Premium Ceded",
                  style="Panel.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(s3, text="This isn't in any register - it's the prior period's own Closing figure, entered "
                            "by hand for now.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 6))
        self.reins_opening_var = tk.StringVar()
        ttk.Entry(s3, textvariable=self.reins_opening_var, width=22, font=("Segoe UI", 10)).pack(anchor="w")

        s4 = self._loss_section_card(body, "4", "Run Calculation", PAD_X, expand_bottom=True)
        ttk.Label(s4, text="Computes Total Reinsurance Ceded per row and builds both tables (Reinsurance "
                            "Premium Ceded Distribution and Prepaid Reinsurance Premium Ceded) for the period "
                            "selected on tab 1.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        calc_btn = self._make_action_button(s4, "Calculate", self._run_reins_calculation, kind="accent")
        calc_btn.pack(anchor="w")

        self.reins_calc_progress = perf_utils.ProgressBarPanel(s4)
        self.reins_calc_progress.frame.pack(fill="x", pady=(10, 0))

        self.reins_calc_summary_var = tk.StringVar(value="")
        ttk.Label(s4, textvariable=self.reins_calc_summary_var, foreground=self.COLOR_SUCCESS,
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 8))

        tree_frame = ttk.Frame(s4)
        tree_frame.pack(fill="both", expand=True)
        columns = ("department", "method", "ceded_total", "prepaid_total")
        self.reins_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)
        self.reins_results_tree.heading("department", text="Department")
        self.reins_results_tree.heading("method", text="Method")
        self.reins_results_tree.heading("ceded_total", text="Total Reinsurance Ceded")
        self.reins_results_tree.heading("prepaid_total", text="Prepaid (Unamortized) Total")
        self.reins_results_tree.column("department", width=180)
        self.reins_results_tree.column("method", width=80, anchor="center")
        self.reins_results_tree.column("ceded_total", width=180, anchor="e")
        self.reins_results_tree.column("prepaid_total", width=190, anchor="e")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.reins_results_tree.yview)
        self.reins_results_tree.configure(yscrollcommand=vsb.set)
        self.reins_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._finalize_scroll_passthrough(body)

    def _run_reins_calculation(self):
        if self.reins_loaded_df is None:
            messagebox.showerror("Error", "Load the RI Premium Ceded register(s) on tab 1 first.")
            return
        if not self.reins_period_options:
            messagebox.showerror("Error", "No period options available - load register(s) with usable Comm Date values first.")
            return
        selected_label = self.reins_period_var.get()
        option = next((o for o in self.reins_period_options if o["label"] == selected_label), None)
        if option is None:
            messagebox.showerror("Error", "Select a period on tab 1 first.")
            return

        opening_text = self.reins_opening_var.get().strip().replace(",", "")
        try:
            opening_prepaid = float(opening_text) if opening_text else 0.0
        except ValueError:
            messagebox.showerror("Invalid entry", f"'{opening_text}' isn't a valid number for Opening Prepaid Reinsurance Ceded.")
            return

        paths = list(self.reins_register_paths)
        period_year, period_month = option["year"], option["month"]
        self.reins_calc_summary_var.set("")
        self.reins_calc_progress.start(total=None, text="Calculating...")

        def worker(report):
            return rip.run_reinsurance_calculation(paths, period_year, period_month, report=report)

        def on_done(result):
            self.reins_last_result = result
            self.reins_last_opening = opening_prepaid
            self.reins_calc_progress.finish("Done.")
            self.reins_calc_summary_var.set(
                f"Closing Prepaid Reinsurance Premium Ceded: PKR {result['closing_prepaid']:,.0f}")

            self.reins_results_tree.delete(*self.reins_results_tree.get_children())
            for dept in result["departments"]:
                ct = rip.row_total(result["ceded_table"][dept])
                pt = rip.row_total(result["prepaid_table"][dept])
                self.reins_results_tree.insert("", "end", values=(
                    dept, result["methods"][dept], f"{ct:,.0f}", f"{pt:,.0f}"
                ))
            for w in result["warnings"]:
                self._log_reins(f"WARNING: {w}")
            self.status_var.set(f"Reinsurance Automation: Closing Prepaid = PKR {result['closing_prepaid']:,.0f}. See tab 3 to generate.")

        def on_error(exc):
            self.reins_calc_progress.finish("Failed - see the message for details.")
            messagebox.showerror("Calculation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.reins_calc_progress.update,
            on_done=on_done, on_error=on_error,
        )

    # ---- Tab 3: Generate & Save ----
    def _build_reins_tab_generate(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s5 = self._loss_section_card(body, "5", "Generate & Save", PAD_X, expand_bottom=True)
        out_row = ttk.Frame(s5)
        out_row.pack(fill="x", pady=(0, 8))
        ttk.Label(out_row, text="Output folder", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.reins_output_dir_var = tk.StringVar(value=self.reins_output_dir)
        ttk.Entry(out_row, textvariable=self.reins_output_dir_var, width=50).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Browse...", command=self._browse_reins_output_dir).pack(side="left")

        gen_row = ttk.Frame(s5)
        gen_row.pack(fill="x", pady=(4, 8))
        self._make_action_button(gen_row, "Generate Reinsurance Workbook", self._generate_reins_workbook, kind="accent").pack(side="left")
        ttk.Button(gen_row, text="Open Output Folder", command=self._open_reins_output_folder).pack(side="left", padx=(8, 0))

        self.reins_gen_progress = perf_utils.ProgressBarPanel(s5)
        self.reins_gen_progress.frame.pack(fill="x", pady=(4, 8))

        self.reins_gen_log = scrolledtext.ScrolledText(s5, height=10, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.reins_gen_log.pack(fill="both", expand=True, pady=(8, 0))

        self._finalize_scroll_passthrough(body)

    def _log_reins_gen(self, msg):
        self.reins_gen_log.insert("end", str(msg) + "\n")
        self.reins_gen_log.see("end")

    def _browse_reins_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.reins_output_dir_var.set(path)

    def _generate_reins_workbook(self):
        if self.reins_last_result is None:
            messagebox.showerror("Nothing to generate", "Run Calculate on tab 2 first.")
            return
        base_output_dir = self.reins_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder first.")
            return

        output_dir = self._new_run_folder(base_output_dir)
        result = self.reins_last_result
        opening_prepaid = getattr(self, "reins_last_opening", 0.0)
        filename = rig.build_output_filename(result["period_end"])
        output_file = os.path.join(output_dir, filename)

        self.reins_gen_log.delete("1.0", "end")
        self.status_var.set("Generating Reinsurance workbook...")
        self.reins_gen_progress.start(total=None, text="Generating...")

        def worker(report):
            return rig.build_reinsurance_workbook(result, opening_prepaid, output_file)

        def on_done(fp):
            self.reins_last_output_file = fp
            self.reins_gen_progress.finish("Done.")
            self._log_reins_gen(f"Generated: {fp}")
            messagebox.showinfo("Done", f"Reinsurance workbook generated:\n{fp}")
            self.status_var.set(f"Reinsurance Automation: generated {os.path.basename(fp)}.")

        def on_error(exc):
            self.reins_gen_progress.finish("Failed - see the log for details.")
            self._log_reins_gen(f"ERROR: {exc}")
            messagebox.showerror("Generation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.reins_gen_progress.update,
            on_done=on_done, on_error=on_error,
        )

    def _open_reins_output_folder(self):
        output_dir = os.path.dirname(self.reins_last_output_file) if self.reins_last_output_file else self.reins_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate the workbook first.")

    # ---------- Deferred Commission Expense module ----------
    def _build_deferred_commission_page(self, parent):
        """Deferred Commission Expense: upload the Premium Register(s)
        plus the three commission/service-charge ledger exports
        (Commission Exp - Direct, Commission Exp - Coinsurance Accepted,
        Service Charges - Coinsurance), select a period, enter the
        Opening Deferred Commission Expense balance, and generate the
        Deferred Commission working workbook - same 3-tab shape as UPR
        and Reinsurance Automation, since the underlying calculation
        reuses UPR's own period/bucketing/amortization logic directly
        once each ledger's balance has been matched onto its own policy
        in the Premium Register."""
        header = ttk.Frame(parent)
        header.pack(fill="x", padx=20, pady=(20, 6))
        ttk.Label(header, text="Deferred Commission", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        dc_subtitle = ttk.Label(
            header, text="Matches each of the three commission/service-charge ledgers to its own policy in the "
                          "Premium Register (by Doc Ref #, after keeping only the reference before a '~' in the "
                          "ledger's own Manual / Ref. No.), sums them into a Total Commission per policy, then "
                          "amortizes that total using the same 1/24 (most departments) / 1/6 (Marine, Travel) "
                          "method as UPR Automation to produce the Closing Deferred Commission Expense figure "
                          "for a selected 12-month period.",
            foreground=self.COLOR_MUTED, font=("Segoe UI", 10), justify="left")
        dc_subtitle.pack(anchor="w", fill="x", pady=(2, 10))

        def _wrap_dc_subtitle(event):
            dc_subtitle.configure(wraplength=max(200, event.width - 4))
        header.bind("<Configure>", _wrap_dc_subtitle)

        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        tab_files = ttk.Frame(notebook)
        tab_opening = ttk.Frame(notebook)
        tab_generate = ttk.Frame(notebook)
        notebook.add(tab_files, text="1. Register, Ledgers & Period")
        notebook.add(tab_opening, text="2. Opening Balance & Calculate")
        notebook.add(tab_generate, text="3. Generate & Save")

        self._build_dc_tab_files(tab_files)
        self._build_dc_tab_opening(tab_opening)
        self._build_dc_tab_generate(tab_generate)

    # ---- Small generic file-list helpers, shared by the Premium
    # Register card and all three ledger cards below - each card gets
    # its own listbox/paths pair (e.g. "dc_register_listbox" /
    # "dc_register_paths"), looked up by attribute name so the same
    # Browse/Remove/Clear handlers work for all four without repeating
    # near-identical code four times over. ----
    def _build_dc_file_list_card(self, parent, listbox_attr, paths_attr, height=4):
        list_card = tk.Frame(parent, bg=self.COLOR_PANEL_ALT, highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_card.pack(fill="x")
        list_outer = tk.Frame(list_card, bg=self.COLOR_PANEL_ALT)
        list_outer.pack(fill="x", padx=12, pady=10)

        list_frame = tk.Frame(list_outer, bg="white", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        list_frame.pack(fill="x")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(list_frame, height=height, font=("Segoe UI", 10), selectmode="extended",
                              activestyle="none", relief="flat", borderwidth=0,
                              bg="white", selectbackground=self.COLOR_ACCENT, selectforeground="white",
                              highlightthickness=0, yscrollcommand=list_scroll.set)
        list_scroll.configure(command=listbox.yview)
        listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")
        setattr(self, listbox_attr, listbox)

        toolbar = tk.Frame(list_outer, bg=self.COLOR_PANEL_ALT)
        toolbar.pack(fill="x", pady=(8, 0))
        self._make_pill_button(toolbar, "Browse...", lambda: self._dc_add_files(listbox_attr, paths_attr),
                                bg=self.COLOR_ACCENT).pack(side="left")
        self._make_pill_button(toolbar, "Remove Selected", lambda: self._dc_remove_file(listbox_attr, paths_attr),
                                bg=self.COLOR_PANEL, fg=self.COLOR_TEXT, outline=True).pack(side="left", padx=(8, 0))
        self._make_pill_button(toolbar, "Clear All", lambda: self._dc_clear_files(listbox_attr, paths_attr),
                                bg=self.COLOR_PANEL, fg=self.COLOR_WARN, outline=True).pack(side="left", padx=(8, 0))

    def _dc_add_files(self, listbox_attr, paths_attr):
        listbox = getattr(self, listbox_attr)
        paths = getattr(self, paths_attr)
        selected = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        added = 0
        for p in selected:
            if p not in paths:
                paths.append(p)
                listbox.insert("end", os.path.basename(p))
                added += 1
        if added:
            self.status_var.set(f"Deferred Commission: added {added} file(s) ({len(paths)} total).")

    def _dc_remove_file(self, listbox_attr, paths_attr):
        listbox = getattr(self, listbox_attr)
        paths = getattr(self, paths_attr)
        selected = list(listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            listbox.delete(idx)
            del paths[idx]

    def _dc_clear_files(self, listbox_attr, paths_attr):
        getattr(self, listbox_attr).delete(0, "end")
        getattr(self, paths_attr).clear()

    # ---- Tab 1: Premium Register + 3 ledgers + period ----
    def _build_dc_tab_files(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s1 = self._loss_section_card(body, "1", "Premium Register File(s)", PAD_X)
        dc_s1_note = ttk.Label(
            s1, text="The same conventional Premium Register file(s) UPR Automation uses - one consolidated "
                      "file or several per-year files, either works. Required columns: Department/Class, "
                      "Comm Date, Gross Premium/Contribution, Doc Ref # (used to match ledger balances).",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        dc_s1_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(dc_s1_note, s1, side_padding=10)
        self._build_dc_file_list_card(s1, "dc_register_listbox", "dc_register_paths", height=4)

        s2 = self._loss_section_card(body, "2", "Commission Exp \u2013 Direct Ledger File(s)", PAD_X)
        dc_s2_note = ttk.Label(
            s2, text="Account ledger export(s) for Commission Expense - Direct. Upload as many periods as "
                      "needed (e.g. a 2025 file and a Jan-Jun 2026 file); they're combined automatically.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        dc_s2_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(dc_s2_note, s2, side_padding=10)
        self._build_dc_file_list_card(s2, "dc_direct_listbox", "dc_direct_paths", height=3)

        s3 = self._loss_section_card(body, "3", "Commission Exp \u2013 Coinsurance Accepted Ledger File(s)", PAD_X)
        dc_s3_note = ttk.Label(
            s3, text="Account ledger export(s) for Commission Expense - Coinsurance Accepted.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        dc_s3_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(dc_s3_note, s3, side_padding=10)
        self._build_dc_file_list_card(s3, "dc_coinsurance_listbox", "dc_coinsurance_paths", height=3)

        s4 = self._loss_section_card(body, "4", "Service Charges \u2013 Coinsurance Ledger File(s)", PAD_X)
        dc_s4_note = ttk.Label(
            s4, text="Account ledger export(s) for Service Charges - Coinsurance.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        dc_s4_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(dc_s4_note, s4, side_padding=10)
        self._build_dc_file_list_card(s4, "dc_service_charges_listbox", "dc_service_charges_paths", height=3)

        s5 = self._loss_section_card(body, "5", "Load & Period", PAD_X, expand_bottom=True)
        dc_s5_note = ttk.Label(
            s5, text="Loads the Premium Register(s) to populate the period dropdown below (the three ledgers "
                      "are read on Calculate, tab 2). Select the period END month - the calculation covers "
                      "exactly the 12 months ending there. If the register(s) don't reach back or forward far "
                      "enough to cover the full 12 months, an error is raised rather than guessing.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        dc_s5_note.pack(anchor="w", pady=(0, 10))
        self._bind_dynamic_wraplength(dc_s5_note, s5, side_padding=10)

        load_row = ttk.Frame(s5)
        load_row.pack(fill="x", pady=(0, 10))
        self.dc_load_btn = self._make_action_button(load_row, "Load Register(s)", self._load_dc_registers, kind="accent")
        self.dc_load_btn.pack(side="left")

        self.dc_load_progress = perf_utils.ProgressBarPanel(s5)
        self.dc_load_progress.frame.pack(fill="x", pady=(0, 10))

        period_row = ttk.Frame(s5)
        period_row.pack(fill="x")
        ttk.Label(period_row, text="Period (ending month)", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.dc_period_var = tk.StringVar()
        self.dc_period_combo = ttk.Combobox(period_row, textvariable=self.dc_period_var, state="readonly", width=22)
        self.dc_period_combo.pack(side="left")

        self.dc_register_log = scrolledtext.ScrolledText(s5, height=8, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.dc_register_log.pack(fill="both", expand=True, pady=(12, 0))

        self._finalize_scroll_passthrough(body)

    def _log_dc(self, msg):
        self.dc_register_log.insert("end", str(msg) + "\n")
        self.dc_register_log.see("end")

    def _load_dc_registers(self):
        if not self.dc_register_paths:
            messagebox.showerror("No files selected", "Browse to at least one Premium Register file first.")
            return

        self.dc_register_log.delete("1.0", "end")
        self.status_var.set("Loading premium register(s)...")
        self.dc_load_progress.start(total=None, text="Loading...")
        paths = list(self.dc_register_paths)

        def worker(report):
            df, meta = dcp.up.load_premium_registers(paths, report=report)
            return df, meta

        def on_done(result):
            df, meta = result
            self.dc_loaded_df = df
            self.dc_load_progress.finish(f"Loaded {len(df):,} row(s).")
            self._log_dc(f"Loaded {len(df):,} usable row(s) from {len(meta['sheets_used'])} sheet(s): "
                         + ", ".join(f"{fn} [{sn}]" for fn, sn in meta["sheets_used"]))
            if meta["sheets_skipped"]:
                self._log_dc(f"Skipped {len(meta['sheets_skipped'])} sheet(s) that didn't look like a premium register:")
                for fn, sn, reason in meta["sheets_skipped"]:
                    self._log_dc(f"   {fn} [{sn}]: {reason}")
            if meta["excluded_rows"]:
                self._log_dc(f"{meta['excluded_rows']:,} row(s) excluded (missing Department, Comm Date, "
                             f"Gross Premium, or an implausible Comm Date).")
            if "Doc Ref" not in df.columns:
                self._log_dc("WARNING: no 'Doc Ref #' / 'Document No.' column was recognized in this register - "
                             "ledger balances won't be able to match to any policy. Add the real header text to "
                             "upr_processor.ALIASES's own 'Doc Ref' entry if this is just a naming mismatch.")

            options = dcp.get_available_period_options(df)
            self.dc_period_options = options
            self.dc_period_combo.configure(values=[o["label"] for o in options])
            if options:
                self.dc_period_combo.current(0)
            self.status_var.set(f"Deferred Commission: {len(df):,} register row(s) loaded. Choose a period, then go to tab 2.")

        def on_error(exc):
            self.dc_load_progress.finish("Failed - see the log for details.")
            self._log_dc(f"ERROR: {exc}")
            messagebox.showerror("Could not load register(s)", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_dc(msg),
            on_progress_value=self.dc_load_progress.update,
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.dc_load_btn],
        )

    # ---- Tab 2: Opening balance + Calculate ----
    def _build_dc_tab_opening(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s6 = self._loss_section_card(body, "6", "Opening Balance", PAD_X)
        ttk.Label(s6, text="Opening Deferred Commission Expense",
                  style="Panel.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(s6, text="This isn't in any register or ledger - it's the prior period's own Closing figure, "
                            "entered by hand for now.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 6))
        self.dc_opening_var = tk.StringVar()
        ttk.Entry(s6, textvariable=self.dc_opening_var, width=22, font=("Segoe UI", 10)).pack(anchor="w")

        s7 = self._loss_section_card(body, "7", "Run Calculation", PAD_X, expand_bottom=True)
        ttk.Label(s7, text="Re-reads the Premium Register(s) and all three ledgers, matches each ledger's "
                            "balance onto its own policy by Doc Ref #, and builds both tables (Commission "
                            "Paid/Payable Distribution and Deferred Commission Expense) for the period "
                            "selected on tab 1.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        calc_btn = self._make_action_button(s7, "Calculate", self._run_dc_calculation, kind="accent")
        calc_btn.pack(anchor="w")

        self.dc_calc_progress = perf_utils.ProgressBarPanel(s7)
        self.dc_calc_progress.frame.pack(fill="x", pady=(10, 0))

        self.dc_calc_summary_var = tk.StringVar(value="")
        ttk.Label(s7, textvariable=self.dc_calc_summary_var, foreground=self.COLOR_SUCCESS,
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 8))

        tree_frame = ttk.Frame(s7)
        tree_frame.pack(fill="both", expand=True)
        columns = ("department", "method", "commission_total", "deferred_total")
        self.dc_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)
        self.dc_results_tree.heading("department", text="Department")
        self.dc_results_tree.heading("method", text="Method")
        self.dc_results_tree.heading("commission_total", text="Commission Paid/Payable")
        self.dc_results_tree.heading("deferred_total", text="Deferred Commission Expense")
        self.dc_results_tree.column("department", width=180)
        self.dc_results_tree.column("method", width=80, anchor="center")
        self.dc_results_tree.column("commission_total", width=190, anchor="e")
        self.dc_results_tree.column("deferred_total", width=200, anchor="e")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.dc_results_tree.yview)
        self.dc_results_tree.configure(yscrollcommand=vsb.set)
        self.dc_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._finalize_scroll_passthrough(body)

    def _run_dc_calculation(self):
        if self.dc_loaded_df is None:
            messagebox.showerror("Error", "Load the Premium Register(s) on tab 1 first.")
            return
        if not self.dc_period_options:
            messagebox.showerror("Error", "No period options available - load register(s) with usable Comm Date values first.")
            return
        selected_label = self.dc_period_var.get()
        option = next((o for o in self.dc_period_options if o["label"] == selected_label), None)
        if option is None:
            messagebox.showerror("Error", "Select a period on tab 1 first.")
            return
        if not self.dc_direct_paths and not self.dc_coinsurance_paths and not self.dc_service_charges_paths:
            messagebox.showerror("No ledger files selected", "Browse to at least one commission/service-charge "
                                  "ledger file on tab 1 first.")
            return

        opening_text = self.dc_opening_var.get().strip().replace(",", "")
        try:
            opening_deferred = float(opening_text) if opening_text else 0.0
        except ValueError:
            messagebox.showerror("Invalid entry", f"'{opening_text}' isn't a valid number for Opening Deferred Commission Expense.")
            return

        register_paths = list(self.dc_register_paths)
        direct_paths = list(self.dc_direct_paths)
        coinsurance_paths = list(self.dc_coinsurance_paths)
        service_charge_paths = list(self.dc_service_charges_paths)
        period_year, period_month = option["year"], option["month"]
        self.dc_calc_summary_var.set("")
        self.dc_calc_progress.start(total=None, text="Calculating...")

        def worker(report):
            return dcp.run_deferred_commission_calculation(
                register_paths, direct_paths, coinsurance_paths, service_charge_paths,
                period_year, period_month, report=report)

        def on_done(result):
            self.dc_last_result = result
            self.dc_last_opening = opening_deferred
            self.dc_calc_progress.finish("Done.")
            self.dc_calc_summary_var.set(
                f"Closing Deferred Commission Expense: PKR {result['closing_deferred_commission']:,.0f}")

            self.dc_results_tree.delete(*self.dc_results_tree.get_children())
            for dept in result["departments"]:
                ct = dcp.row_total(result["commission_table"][dept])
                dt = dcp.row_total(result["deferred_table"][dept])
                self.dc_results_tree.insert("", "end", values=(
                    dept, result["methods"][dept], f"{ct:,.0f}", f"{dt:,.0f}"
                ))
            for w in result["warnings"]:
                self._log_dc(f"WARNING: {w}")
            self.status_var.set(f"Deferred Commission: Closing = PKR {result['closing_deferred_commission']:,.0f}. See tab 3 to generate.")

        def on_error(exc):
            self.dc_calc_progress.finish("Failed - see the message for details.")
            messagebox.showerror("Calculation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.dc_calc_progress.update,
            on_done=on_done, on_error=on_error,
        )

    # ---- Tab 3: Generate & Save ----
    def _build_dc_tab_generate(self, tab):
        body = self._make_upr_scrollable_body(tab)
        PAD_X = 16

        s8 = self._loss_section_card(body, "8", "Generate & Save", PAD_X, expand_bottom=True)
        out_row = ttk.Frame(s8)
        out_row.pack(fill="x", pady=(0, 8))
        ttk.Label(out_row, text="Output folder", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.dc_output_dir_var = tk.StringVar(value=self.dc_output_dir)
        ttk.Entry(out_row, textvariable=self.dc_output_dir_var, width=50).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Browse...", command=self._browse_dc_output_dir).pack(side="left")

        gen_row = ttk.Frame(s8)
        gen_row.pack(fill="x", pady=(4, 8))
        self._make_action_button(gen_row, "Generate Deferred Commission Workbook", self._generate_dc_workbook, kind="accent").pack(side="left")
        ttk.Button(gen_row, text="Open Output Folder", command=self._open_dc_output_folder).pack(side="left", padx=(8, 0))

        self.dc_gen_progress = perf_utils.ProgressBarPanel(s8)
        self.dc_gen_progress.frame.pack(fill="x", pady=(4, 8))

        self.dc_gen_log = scrolledtext.ScrolledText(s8, height=10, font=("Consolas", 9), relief="solid", borderwidth=1)
        self.dc_gen_log.pack(fill="both", expand=True, pady=(8, 0))

        self._finalize_scroll_passthrough(body)

    def _log_dc_gen(self, msg):
        self.dc_gen_log.insert("end", str(msg) + "\n")
        self.dc_gen_log.see("end")

    def _browse_dc_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.dc_output_dir_var.set(path)

    def _generate_dc_workbook(self):
        if self.dc_last_result is None:
            messagebox.showerror("Nothing to generate", "Run Calculate on tab 2 first.")
            return
        base_output_dir = self.dc_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder first.")
            return

        output_dir = self._new_run_folder(base_output_dir)
        result = self.dc_last_result
        opening_deferred = getattr(self, "dc_last_opening", 0.0)
        filename = dcg.build_output_filename(result["period_end"])
        output_file = os.path.join(output_dir, filename)

        self.dc_gen_log.delete("1.0", "end")
        self.status_var.set("Generating Deferred Commission workbook...")
        self.dc_gen_progress.start(total=None, text="Generating...")

        def worker(report):
            return dcg.build_deferred_commission_workbook(result, opening_deferred, output_file)

        def on_done(fp):
            self.dc_last_output_file = fp
            self.dc_gen_progress.finish("Done.")
            self._log_dc_gen(f"Generated: {fp}")
            messagebox.showinfo("Done", f"Deferred Commission workbook generated:\n{fp}")
            self.status_var.set(f"Deferred Commission: generated {os.path.basename(fp)}.")

        def on_error(exc):
            self.dc_gen_progress.finish("Failed - see the log for details.")
            self._log_dc_gen(f"ERROR: {exc}")
            messagebox.showerror("Generation failed", str(exc))

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.dc_gen_progress.update,
            on_done=on_done, on_error=on_error,
        )

    def _open_dc_output_folder(self):
        output_dir = os.path.dirname(self.dc_last_output_file) if self.dc_last_output_file else self.dc_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate the workbook first.")


    # ---------- Coinsurance > Premium Register module ----------
    def _build_premium_register_page(self, parent):
        """Premium Register (Coinsurance) - matches a co-insurance leader's
        own policy numbers against IGI's Premium Register and pulls in the
        matched IGI Policy Number, flagging unmatched rows for manual
        entry. Single-page workflow (no multi-tab notebook needed, since
        there's no grouping/email step) - same scrollable numbered-card
        pattern as the Loss Ratio page."""
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
        ttk.Label(header, text="Policy Check", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        ttk.Label(header, text="Match a co-insurance leader's own policy numbers against IGI's Premium "
                                "Register, and pull in the matched IGI Policy Number automatically.",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

        info_badge = tk.Frame(header, bg=self.COLOR_ACCENT_SOFT)
        info_badge.pack(anchor="w")
        tk.Frame(info_badge, bg=self.COLOR_GOLD, width=3).pack(side="left", fill="y")
        tk.Label(info_badge, text="Fuzzy digit-run matching \u2014 unmatched rows are flagged red for manual entry, not guessed",
                 bg=self.COLOR_ACCENT_SOFT, fg=self.COLOR_PRIMARY, font=("Segoe UI", 9, "bold"),
                 padx=14, pady=8).pack(side="left")

        # --- Section 1: Input files ---
        s1 = self._loss_section_card(body, "1", "Input Files", PAD_X)

        # Register file type - "Premium Register" (the original, plain
        # export) or "Outstanding Premium File" (IGI's own "Co Insurance
        # Outstanding Position" export, e.g. a "Co_Insurance_<date>.xlsx"
        # workbook's "Detail" sheet). Both are accepted and matched
        # identically either way (this only ever changes what's PULLED
        # THROUGH into the output, never how matching itself works) -
        # picking "Outstanding Premium File" additionally looks up that
        # file's own "Premium Payable" column and carries it into the
        # match output as an extra column, since that figure only exists
        # on this file shape, never on a plain Premium Register.
        filetype_row = ttk.Frame(s1)
        filetype_row.pack(fill="x", pady=(0, 4))
        ttk.Label(filetype_row, text="Register file type:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 14))
        self.premreg_file_type_var = tk.StringVar(value="premium_register")
        ttk.Radiobutton(filetype_row, text="Premium Register", variable=self.premreg_file_type_var,
                         value="premium_register").pack(side="left", padx=(0, 20))
        ttk.Radiobutton(filetype_row, text="Outstanding Premium File", variable=self.premreg_file_type_var,
                         value="outstanding_premium").pack(side="left")
        premreg_filetype_note = ttk.Label(
            s1, text="",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premreg_filetype_note.pack(anchor="w", pady=(0, 12))
        self._bind_dynamic_wraplength(premreg_filetype_note, s1, side_padding=10)
        self.premreg_filetype_note_var = tk.StringVar()
        premreg_filetype_note.configure(textvariable=self.premreg_filetype_note_var)

        premreg_s1_note = ttk.Label(s1, text="Both files accept .xlsx or .xls with auto-detected headers. The right sheet "
                            "is found automatically in a multi-tab workbook, and column naming (\"NUMS\", \"Policy No.\", "
                            "\"Leader#\", etc.) is matched automatically too.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premreg_s1_note.pack(anchor="w", pady=(0, 12))
        self._bind_dynamic_wraplength(premreg_s1_note, s1, side_padding=10)

        leader_row = self._premreg_file_row(s1, "Leader's Policy File", "")
        (self.premreg_leader_icon, self.premreg_leader_name_var, leader_btn_col,
         self.premreg_leader_title_var, self.premreg_leader_subtitle_var) = leader_row
        self._make_pill_button(leader_btn_col, "Browse...", self._browse_premreg_leader_file,
                                bg=self.COLOR_ACCENT).pack(side="left")

        register_row = self._premreg_file_row(s1, "", "")
        (self.premreg_register_icon, self.premreg_register_name_var, register_btn_col,
         self.premreg_register_title_var, self.premreg_register_subtitle_var) = register_row
        self._make_pill_button(register_btn_col, "Browse...", self._browse_premreg_register_file,
                                bg=self.COLOR_ACCENT).pack(side="left")

        # Fills in the filetype note + both card titles/subtitles for
        # whichever mode is currently selected, and re-runs automatically
        # whenever the radio button changes - see _update_premreg_file_hints.
        self.premreg_file_type_var.trace_add("write", lambda *a: self._update_premreg_file_hints())
        self._update_premreg_file_hints()

        # --- Section 2: Run matching ---
        s2 = self._loss_section_card(body, "2", "Run Matching", PAD_X)
        premreg_s2_note = ttk.Label(s2, text="Adds the matched IGI Policy Number, matched reference, match score, "
                            "and Matched/Not Matched status to every row.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premreg_s2_note.pack(anchor="w", pady=(0, 12))
        self._bind_dynamic_wraplength(premreg_s2_note, s2, side_padding=10)
        run_btn = self._make_action_button(s2, "Run Matching", self._run_premium_register_match, kind="accent")
        run_btn.pack(anchor="w")
        self.premreg_summary_var = tk.StringVar(value="")
        ttk.Label(s2, textvariable=self.premreg_summary_var, foreground=self.COLOR_SUCCESS,
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 0))

        self.premreg_progress = perf_utils.ProgressBarPanel(s2)
        self.premreg_progress.frame.pack(fill="x", pady=(10, 0))

        # --- Section 3: Log ---
        s3 = self._loss_section_card(body, "3", "Processing Log", PAD_X, expand_bottom=True)
        self.premreg_log = scrolledtext.ScrolledText(s3, height=10, font=("Consolas", 9),
                                                       relief="solid", borderwidth=1)
        self.premreg_log.pack(fill="both", expand=True)

    def _premreg_file_row(self, parent, title, subtitle):
        """One file-picker card, matching the Loss Ratio Individual-mode
        card style. Returns (icon_label, name_var, button_holder_frame,
        title_var, subtitle_var) - title_var/subtitle_var are tk.StringVars
        so a caller can swap the card's own title/description text later
        (e.g. Policy Check's own leader/register cards update their text
        depending on the "Register file type" radio button - see
        _update_premreg_file_hints), without needing to know the label
        widgets themselves. Existing callers that only ever use
        row[0]/row[1]/row[2] are unaffected by these extra elements."""
        card = tk.Frame(parent, bg=self.COLOR_PANEL_ALT,
                         highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=self.COLOR_PANEL_ALT)
        inner.pack(fill="x", padx=16, pady=14)

        icon = tk.Label(inner, text="\u25A4", font=("Segoe UI", 20),
                         bg=self.COLOR_PANEL_ALT, fg=self.COLOR_BORDER, width=2)
        icon.pack(side="left", padx=(0, 14))

        text_col = tk.Frame(inner, bg=self.COLOR_PANEL_ALT)
        text_col.pack(side="left", fill="x", expand=True)
        title_var = tk.StringVar(value=title)
        tk.Label(text_col, textvariable=title_var, font=("Segoe UI", 10, "bold"),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_TEXT, anchor="w").pack(fill="x", anchor="w")
        name_var = tk.StringVar(value="No file selected")
        tk.Label(text_col, textvariable=name_var, font=("Segoe UI", 9),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_MUTED, anchor="w",
                 wraplength=420, justify="left").pack(fill="x", anchor="w")
        subtitle_var = tk.StringVar(value=subtitle)
        tk.Label(text_col, textvariable=subtitle_var, font=("Segoe UI", 8),
                 bg=self.COLOR_PANEL_ALT, fg=self.COLOR_MUTED, anchor="w",
                 wraplength=420, justify="left").pack(fill="x", anchor="w", pady=(2, 0))

        btn_col = tk.Frame(inner, bg=self.COLOR_PANEL_ALT)
        btn_col.pack(side="right", padx=(14, 0))
        return icon, name_var, btn_col, title_var, subtitle_var

    def _update_premreg_file_hints(self):
        """
        Keeps Policy Check's own filetype note and both file cards
        (Leader's Policy File / register card) in sync with whichever
        "Register file type" radio button is currently selected -
        called once at build time and again every time that variable
        changes (see the trace_add wired up right after both cards are
        created). Nothing about matching itself depends on this; it's
        purely so the two upload areas always describe what they
        actually expect and what the run will actually do, instead of
        one fixed, one-size-fits-all description that doesn't change
        when the person switches modes.
        """
        mode = self.premreg_file_type_var.get()
        if mode == "outstanding_premium":
            self.premreg_filetype_note_var.set(
                "Match against IGI's Co-Insurance Outstanding Position export (\"Detail\" sheet). Also pulls in "
                "Premium Payable and Total Payable for each matched row."
            )
            self.premreg_leader_title_var.set("Leader's Policy File")
            self.premreg_leader_subtitle_var.set(
                "The leader's own policy number column (e.g. \"NUMS\") is matched against \"Leader Ref. No\" below."
            )
            self.premreg_register_title_var.set("Outstanding Premium File")
            self.premreg_register_subtitle_var.set(
                "Carries the leader's reference alongside our IGI Policy Number, Premium Payable, and Total Payable."
            )
        else:
            self.premreg_filetype_note_var.set(
                "Match against a plain Premium Register export (default). Switch to \"Outstanding Premium File\" "
                "above if matching against the Outstanding Position export instead."
            )
            self.premreg_leader_title_var.set("Leader's Policy File")
            self.premreg_leader_subtitle_var.set(
                "The leader's own policy number column (e.g. \"NUMS\") is matched against our Leader Ref Number."
            )
            self.premreg_register_title_var.set("IGI Premium Register")
            self.premreg_register_subtitle_var.set(
                "Carries the leader's reference (Leader#) alongside our IGI Policy Number (Doc Ref #)."
            )

    def _browse_premreg_leader_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls"), ("All files", "*.*"),
        ])
        if path:
            self.premreg_leader_file = path
            self.premreg_leader_name_var.set(os.path.basename(path))
            self.premreg_leader_icon.configure(fg=self.COLOR_ACCENT)

    def _browse_premreg_register_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls"), ("All files", "*.*"),
        ])
        if path:
            self.premreg_register_file = path
            self.premreg_register_name_var.set(os.path.basename(path))
            self.premreg_register_icon.configure(fg=self.COLOR_ACCENT)

    def _log_premreg(self, msg):
        self.premreg_log.insert("end", str(msg) + "\n")
        self.premreg_log.see("end")

    def _run_premium_register_match(self):
        if not self.premreg_leader_file:
            messagebox.showerror("No leader file selected", "Browse to the leader's policy file first.")
            return
        if not self.premreg_register_file:
            messagebox.showerror("No register file selected", "Browse to IGI's Premium Register file first.")
            return
        if not os.path.exists(self.premreg_leader_file):
            messagebox.showerror("File not found", f"Could not find:\n{self.premreg_leader_file}")
            return
        if not os.path.exists(self.premreg_register_file):
            messagebox.showerror("File not found", f"Could not find:\n{self.premreg_register_file}")
            return

        self.premreg_log.delete("1.0", "end")
        self.premreg_summary_var.set("")
        self._log_premreg(f"Leader file:   {os.path.basename(self.premreg_leader_file)}")
        self._log_premreg(f"Register file: {os.path.basename(self.premreg_register_file)}\n")
        self.status_var.set("Matching policy numbers...")
        self.premreg_progress.start(total=None, text="Loading files...")

        leader_file = self.premreg_leader_file
        register_file = self.premreg_register_file
        # Read on the main thread, same rule as every other Tkinter
        # variable this worker closure captures - Tkinter widgets/vars
        # must never be queried from a background thread.
        register_file_type = self.premreg_file_type_var.get()

        def worker(report):
            report("Loading leader file...")
            leader_df, nums_col = prp.load_leader_file(leader_file)
            report(f"  {len(leader_df):,} row(s) loaded. Policy-number column detected as '{nums_col}'.")

            report("Loading IGI Premium Register..." if register_file_type == "premium_register"
                   else "Loading Outstanding Premium file...")
            register_df, leader_ref_col, igi_policy_col = prp.load_register_file(register_file)
            report(f"  {len(register_df):,} row(s) loaded. Leader Ref column='{leader_ref_col}', "
                   f"IGI Policy Number column='{igi_policy_col}'.")

            premium_payable_col = None
            total_payable_col = None
            if register_file_type == "outstanding_premium":
                premium_payable_col = prp.find_premium_payable_column(register_df)
                total_payable_col = prp.find_total_payable_column(register_df)
                if premium_payable_col:
                    report(f"  Outstanding Premium file selected - found a 'Premium Payable' column "
                           f"('{premium_payable_col}'); it will be pulled into the output for every matched row.")
                else:
                    report("  WARNING: \"Outstanding Premium File\" was selected, but no 'Premium Payable' "
                           "column was found in this file - continuing without it.")
                if total_payable_col:
                    report(f"  Also found a 'Total Payable' column ('{total_payable_col}') - pulled into the "
                           f"output too, since on some files the real outstanding balance sits under Claims/"
                           f"Commission Payable rather than Premium Payable alone, and only shows up in this "
                           f"combined figure.")
                else:
                    report("  WARNING: no 'Total Payable' column was found in this file either - continuing "
                           "without it.")

            report(f"\nMatching {len(leader_df):,} leader row(s) against {len(register_df):,} register row(s)...")
            report.progress(0, len(leader_df), text=f"Matching 0 / {len(leader_df):,} leader row(s)")
            result = prp.match_leader_file(leader_df, nums_col, register_df, leader_ref_col, igi_policy_col,
                                            report=report, premium_payable_col=premium_payable_col,
                                            total_payable_col=total_payable_col)
            return result

        def on_done(result):
            self.premreg_result_df = result
            matched = int((result["Match Status"] == "Matched").sum())
            total = len(result)
            pct = (matched / total) if total else 0
            self._log_premreg(f"\nMatched {matched} of {total} ({pct:.1%}).")
            if matched < total:
                self._log_premreg(f"{total - matched} row(s) unmatched - these will be highlighted red in the output for manual entry.")

            self.premreg_summary_var.set(f"Matched {matched} / {total} ({pct:.1%})")
            self.premreg_progress.finish(f"Matched {matched} / {total} ({pct:.1%}).")
            self.status_var.set(f"Premium Register: matched {matched}/{total} ({pct:.1%}).")

            default_name = prg.build_output_filename(leader_file)
            # Own base folder ("Policy Check"), deliberately distinct from
            # the Coinsurance > Premium > Premium Register (company split)
            # module's own base folder ("Premium Register") - both used to
            # default to the identical "Premium Register" folder name,
            # which meant every run of either module wrote into the same
            # shared folder, with no way to tell which module a given
            # dated/timed subfolder underneath it had come from.
            default_dir = self._new_run_folder(os.path.join(os.getcwd(), "Policy Check"))
            output_path = filedialog.asksaveasfilename(
                title="Save Policy Check Match Results As",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=default_dir,
                initialfile=default_name,
            )
            self._cleanup_if_empty_dir(default_dir)  # remove the seeded folder if nothing was ever written into it
            if not output_path:
                self._log_premreg("\nSave cancelled - results were not written to a file (still available above).")
                return
            self._save_premreg_output(result, output_path, leader_file, matched, total, pct)

        def on_error(exc):
            self.premreg_progress.finish("Matching failed - see the Processing Log for details.")
            self._log_premreg(f"\nERROR: {exc}")
            messagebox.showerror("Matching failed", str(exc))
            self.status_var.set("Premium Register matching failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_premreg(msg),
            on_progress_value=self.premreg_progress.update,
            on_done=on_done, on_error=on_error,
        )

    def _save_premreg_output(self, result, output_path, leader_file, matched, total, pct):
        """Writing the highlighted output workbook is itself a per-row
        openpyxl write (see premium_register_generator.build_match_output)
        - kept on a second background pass so a huge match result doesn't
        freeze the save step either."""
        self.premreg_progress.start(total=len(result), text=f"Writing 0 / {len(result):,} row(s)")

        def worker(report):
            report(f"Writing {len(result):,} row(s) to {os.path.basename(output_path)}...")
            prg.build_match_output(result, output_path, leader_source_filename=leader_file, report=report)

        def on_done(_):
            self.premreg_progress.finish(f"Saved: {os.path.basename(output_path)}")
            self._log_premreg(f"\nSaved: {output_path}")
            self.status_var.set(f"Premium Register saved: {os.path.basename(output_path)}")
            messagebox.showinfo("Done", f"Matched {matched} of {total} ({pct:.1%}).\n\nSaved to:\n{output_path}")

        def on_error(exc):
            self.premreg_progress.finish("Save failed - see the Processing Log for details.")
            self._log_premreg(f"\nERROR writing output file: {exc}")
            messagebox.showerror("Could not save report", f"An error occurred while writing the Excel file:\n\n{exc}")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_premreg(msg),
            on_progress_value=self.premreg_progress.update,
            on_done=on_done, on_error=on_error,
        )

    # ---------- Coinsurance > Claims > Claims Check module ----------
    def _build_claims_check_page(self, parent):
        """Claims Check (Coinsurance) - checks whether a co-insurance
        leader's own claims file references a policy that actually
        exists in IGI's own Premium Register, by matching the claim's
        own Policy/PMD # against the register's Base Doc No. (falling
        back to Document No. as a secondary column). Same single-page,
        numbered-card layout as Policy Check, since this is its
        claims-side counterpart and does the same kind of job."""
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
        ttk.Label(header, text="Claims Check", font=("Segoe UI Semibold", 18, "bold"),
                  foreground=self.COLOR_PRIMARY).pack(anchor="w")
        ttk.Label(header, text="Check whether a co-insurance leader's own claims file references a policy "
                                "that actually exists in our own books, by matching each claim's Policy/PMD # "
                                "against IGI's Premium Register's own Leader Ref Number (the same column Policy "
                                "Check matches against) and pulling in the matched IGI Policy Number (Doc Ref #).",
                  foreground=self.COLOR_MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

        info_badge = tk.Frame(header, bg=self.COLOR_ACCENT_SOFT)
        info_badge.pack(anchor="w")
        tk.Frame(info_badge, bg=self.COLOR_GOLD, width=3).pack(side="left", fill="y")
        tk.Label(info_badge, text="Same fuzzy digit-run matching engine as Policy Check \u2014 unmatched claims are flagged red for manual review, not guessed",
                 bg=self.COLOR_ACCENT_SOFT, fg=self.COLOR_PRIMARY, font=("Segoe UI", 9, "bold"),
                 padx=14, pady=8).pack(side="left")

        # --- Section 1: Input files ---
        s1 = self._loss_section_card(body, "1", "Input Files", PAD_X)
        claimscheck_s1_note = ttk.Label(s1, text="Both files accept .xlsx or .xls with auto-detected headers. The claims "
                            "sheet is found automatically in a multi-tab workbook, by name first.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        claimscheck_s1_note.pack(anchor="w", pady=(0, 12))
        self._bind_dynamic_wraplength(claimscheck_s1_note, s1, side_padding=10)

        leader_row = self._premreg_file_row(
            s1, "Leader's Claims File",
            "The leader's own claims sheet (e.g. \"IGI Claims\") \u2014 Policy/PMD # is matched against our Base Doc No.",
        )
        self.claimscheck_leader_icon, self.claimscheck_leader_name_var = leader_row[0], leader_row[1]
        self._make_pill_button(leader_row[2], "Browse...", self._browse_claimscheck_leader_file,
                                bg=self.COLOR_ACCENT).pack(side="left")

        register_row = self._premreg_file_row(
            s1, "IGI Premium Register",
            "Carries the leader's own reference (Leader Ref Number) alongside our IGI Policy Number (Doc Ref #).",
        )
        self.claimscheck_register_icon, self.claimscheck_register_name_var = register_row[0], register_row[1]
        self._make_pill_button(register_row[2], "Browse...", self._browse_claimscheck_register_file,
                                bg=self.COLOR_ACCENT).pack(side="left")

        # --- Section 2: Run matching ---
        s2 = self._loss_section_card(body, "2", "Run Matching", PAD_X)
        claimscheck_s2_note = ttk.Label(s2, text="Adds the matched IGI Policy Number (Doc Ref #), matched reference, "
                            "match score, and Matched/Not Matched status to every row.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        claimscheck_s2_note.pack(anchor="w", pady=(0, 12))
        self._bind_dynamic_wraplength(claimscheck_s2_note, s2, side_padding=10)
        run_btn = self._make_action_button(s2, "Run Matching", self._run_claims_check_match, kind="accent")
        run_btn.pack(anchor="w")
        self.claimscheck_summary_var = tk.StringVar(value="")
        ttk.Label(s2, textvariable=self.claimscheck_summary_var, foreground=self.COLOR_SUCCESS,
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 0))

        self.claimscheck_progress = perf_utils.ProgressBarPanel(s2)
        self.claimscheck_progress.frame.pack(fill="x", pady=(10, 0))

        # --- Section 3: Log ---
        s3 = self._loss_section_card(body, "3", "Processing Log", PAD_X, expand_bottom=True)
        self.claimscheck_log = scrolledtext.ScrolledText(s3, height=10, font=("Consolas", 9),
                                                           relief="solid", borderwidth=1)
        self.claimscheck_log.pack(fill="both", expand=True)

    def _browse_claimscheck_leader_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls"), ("All files", "*.*"),
        ])
        if path:
            self.claimscheck_leader_file = path
            self.claimscheck_leader_name_var.set(os.path.basename(path))
            self.claimscheck_leader_icon.configure(fg=self.COLOR_ACCENT)

    def _browse_claimscheck_register_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls"), ("All files", "*.*"),
        ])
        if path:
            self.claimscheck_register_file = path
            self.claimscheck_register_name_var.set(os.path.basename(path))
            self.claimscheck_register_icon.configure(fg=self.COLOR_ACCENT)

    def _log_claimscheck(self, msg):
        self.claimscheck_log.insert("end", str(msg) + "\n")
        self.claimscheck_log.see("end")

    def _run_claims_check_match(self):
        leader_file = getattr(self, "claimscheck_leader_file", "")
        register_file = getattr(self, "claimscheck_register_file", "")
        if not leader_file:
            messagebox.showerror("No leader file selected", "Browse to the leader's claims file first.")
            return
        if not register_file:
            messagebox.showerror("No register file selected", "Browse to IGI's Premium Register file first.")
            return
        if not os.path.exists(leader_file):
            messagebox.showerror("File not found", f"Could not find:\n{leader_file}")
            return
        if not os.path.exists(register_file):
            messagebox.showerror("File not found", f"Could not find:\n{register_file}")
            return

        self.claimscheck_log.delete("1.0", "end")
        self.claimscheck_summary_var.set("")
        self._log_claimscheck(f"Leader claims file: {os.path.basename(leader_file)}")
        self._log_claimscheck(f"Register file:      {os.path.basename(register_file)}\n")
        self.status_var.set("Checking claims against the Premium Register...")
        self.claimscheck_progress.start(total=None, text="Loading files...")

        def worker(report):
            report("Loading leader's claims file...")
            leader_df, pmd_col = ccp.load_leader_claims_file(leader_file)
            report(f"  {len(leader_df):,} row(s) loaded. Policy/PMD # column detected as '{pmd_col}'.")

            report("Loading IGI Premium Register...")
            register_df, leader_ref_col, igi_policy_col = ccp.load_register_file(register_file)
            report(f"  {len(register_df):,} row(s) loaded. Leader Ref Number column='{leader_ref_col}', "
                   f"IGI Policy Number column='{igi_policy_col}'.")

            report(f"\nChecking {len(leader_df):,} claim(s) against {len(register_df):,} register row(s)...")
            report.progress(0, len(leader_df), text=f"Checking 0 / {len(leader_df):,} claim(s)")
            # Auto-select the matching engine based on the claims file's
            # own reference shape (see claims_check_processor's own
            # module-level note above match_leader_claims_file_strict):
            # a purely numeric reference (no letters at all - e.g. a
            # GIMS/Centegy system-generated 'POLICY' column) makes the
            # normal engine's branch/class-code safeguard a silent no-op,
            # confirmed to produce real false positives (a claim's own
            # cert/endorsement number coincidentally matching an
            # unrelated policy) - the strict, Doc-No-only matcher is used
            # instead for that shape. Any reference with letters in it
            # (e.g. 'YYYY/MM/CODE#####'-style leader files) keeps using
            # the normal engine, since the strict matcher's own
            # assumption - the leading segment IS the real Doc-No -
            # doesn't hold for that shape.
            if ccp._reference_has_no_letters(leader_df[pmd_col]):
                report("  Reference column has no letters (a purely numeric Doc-No/Cert/Endt/MM/YYYY "
                       "style reference) - using the strict Doc-No-only matcher to avoid false positives "
                       "from a coincidentally-shared cert/endorsement number.")
                result = ccp.match_leader_claims_file_strict(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col, report=report)
            else:
                result = ccp.match_leader_claims_file(leader_df, pmd_col, register_df, leader_ref_col, igi_policy_col, report=report)
            return result

        def on_done(result):
            self.claimscheck_result_df = result
            matched = int((result["Match Status"] == "Matched").sum())
            total = len(result)
            pct = (matched / total) if total else 0
            self._log_claimscheck(f"\nMatched {matched} of {total} ({pct:.1%}).")
            if matched < total:
                self._log_claimscheck(f"{total - matched} claim(s) unmatched - these will be highlighted red in "
                                       f"the output for manual review (no policy found in our own Premium Register).")

            self.claimscheck_summary_var.set(f"Matched {matched} / {total} ({pct:.1%})")
            self.claimscheck_progress.finish(f"Matched {matched} / {total} ({pct:.1%}).")
            self.status_var.set(f"Claims Check: matched {matched}/{total} ({pct:.1%}).")

            default_name = ccg.build_output_filename(leader_file)
            default_dir = self._new_run_folder(os.path.join(os.getcwd(), "Claims Check"))
            output_path = filedialog.asksaveasfilename(
                title="Save Claims Check Results As",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=default_dir,
                initialfile=default_name,
            )
            self._cleanup_if_empty_dir(default_dir)  # remove the seeded folder if nothing was ever written into it
            if not output_path:
                self._log_claimscheck("\nSave cancelled - results were not written to a file (still available above).")
                return
            self._save_claimscheck_output(result, output_path, leader_file, matched, total, pct)

        def on_error(exc):
            self.claimscheck_progress.finish("Matching failed - see the Processing Log for details.")
            self._log_claimscheck(f"\nERROR: {exc}")
            messagebox.showerror("Matching failed", str(exc))
            self.status_var.set("Claims Check matching failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_claimscheck(msg),
            on_progress_value=self.claimscheck_progress.update,
            on_done=on_done, on_error=on_error,
        )

    def _save_claimscheck_output(self, result, output_path, leader_file, matched, total, pct):
        """Writing the highlighted output workbook is itself a per-row
        openpyxl write (see claims_check_generator.build_match_output) -
        kept on a second background pass so a huge result doesn't freeze
        the save step either."""
        self.claimscheck_progress.start(total=len(result), text=f"Writing 0 / {len(result):,} row(s)")

        def worker(report):
            report(f"Writing {len(result):,} row(s) to {os.path.basename(output_path)}...")
            ccg.build_match_output(result, output_path, leader_source_filename=leader_file, report=report)

        def on_done(_):
            self.claimscheck_progress.finish(f"Saved: {os.path.basename(output_path)}")
            self._log_claimscheck(f"\nSaved: {output_path}")
            self.status_var.set(f"Claims Check saved: {os.path.basename(output_path)}")
            messagebox.showinfo("Done", f"Matched {matched} of {total} ({pct:.1%}).\n\nSaved to:\n{output_path}")

        def on_error(exc):
            self.claimscheck_progress.finish("Save failed - see the Processing Log for details.")
            self._log_claimscheck(f"\nERROR writing output file: {exc}")
            messagebox.showerror("Could not save report", f"An error occurred while writing the Excel file:\n\n{exc}")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self._log_claimscheck(msg),
            on_progress_value=self.claimscheck_progress.update,
            on_done=on_done, on_error=on_error,
        )

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
        self.receivable_output_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "Premium Claims Receivable"))
        ttk.Entry(source_box, textvariable=self.receivable_output_dir_var).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_receivable_output_dir).grid(row=1, column=2, **pad)

        receivable_leader_ref_note = ttk.Label(
            source_box, text="Leader Ref Number is required on Premium Receivable rows only \u2014 it's the other "
                              "company's own policy reference, so it's expected to be blank on Claims Receivable "
                              "rows and is never flagged there.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        receivable_leader_ref_note.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(2, 2))
        self._bind_dynamic_wraplength(receivable_leader_ref_note, source_box)

        self.receivable_load_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(source_box, text="Look up recipient emails from a lookup file",
                         variable=self.receivable_load_email_var).grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))

        ttk.Label(source_box, text="Email Lookup File (.xlsx/.csv, optional)", style="Panel.TLabel").grid(row=4, column=0, sticky="w", **pad)
        self.receivable_email_lookup_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.receivable_email_lookup_file_var).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_receivable_email_lookup_file).grid(row=4, column=2, **pad)
        receivable_email_lookup_note = ttk.Label(
            source_box, text="A Company-name column, an Email column, and an optional POC/Contact column. "
                              "Rows are matched to the master file's 'Company' field by exact name. "
                              "If left blank, drafted emails will be flagged \"missing email\" for manual entry.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        receivable_email_lookup_note.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 2))
        self._bind_dynamic_wraplength(receivable_email_lookup_note, source_box)

        # ---- Section: Filters (multi-select, optional) ----
        # Same "Load Filter Options" + multi-select-listbox pattern used by
        # Claims Intimation's own Unit/Client Name/Department filters -
        # populated straight from the master file's own real distinct
        # values, so there's nothing to mistype and nothing to go stale
        # relative to the actual source data. This is now the ONLY filter
        # on this tab (the old exact-name text-entry Filters section was
        # removed as redundant with this one) and can be used before the
        # data is formally "loaded" for analysis - "Load Filter Options"
        # reads the file itself to populate these lists, independent of
        # "Load, Validate & Build Groups" below, so the filter can be set
        # up first if that's the more natural order for a given run.
        multiselect_filter_box = ttk.Labelframe(canvas_frame, text="  Filters (multi-select, optional)  ", padding=14)
        multiselect_filter_box.pack(fill="x", padx=12, pady=8)
        multiselect_filter_box.columnconfigure(0, weight=1)
        multiselect_filter_box.columnconfigure(1, weight=1)
        multiselect_filter_box.columnconfigure(2, weight=1)

        receivable_multiselect_note = ttk.Label(
            multiselect_filter_box, text="Click \"Load Filter Options\" above to populate these lists, then select "
                                          "before clicking \"Load, Validate & Build Groups\" below.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        receivable_multiselect_note.grid(row=0, column=0, columnspan=3, sticky="ew", padx=6, pady=(0, 8))
        self._bind_dynamic_wraplength(receivable_multiselect_note, multiselect_filter_box, side_padding=10)

        load_receivable_filters_row = ttk.Frame(multiselect_filter_box, style="Panel.TFrame")
        load_receivable_filters_row.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 10))
        self.receivable_load_filters_btn = ttk.Button(load_receivable_filters_row, text="Load Filter Options",
                                                        command=self._load_receivable_filter_options)
        self.receivable_load_filters_btn.pack(side="left")
        self.receivable_filter_status_var = tk.StringVar(value="")
        ttk.Label(load_receivable_filters_row, textvariable=self.receivable_filter_status_var, style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(side="left", padx=(10, 0))

        def _receivable_filter_listbox_col(parent, col_idx, title):
            col = ttk.Frame(parent, style="Panel.TFrame")
            col.grid(row=2, column=col_idx, sticky="nsew", padx=6, pady=4)
            ttk.Label(col, text=title, style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
            ttk.Label(col, text="Ctrl+click to select multiple \u00b7 none = all", style="Muted.TLabel",
                      background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
            frame = tk.Frame(col, highlightbackground="#D6DCE1", highlightthickness=1)
            frame.pack(fill="both", expand=True)
            scroll = ttk.Scrollbar(frame, orient="vertical")
            listbox = tk.Listbox(frame, selectmode="extended", height=10, width=32,
                                  exportselection=False, font=("Segoe UI", 10),
                                  activestyle="none", relief="flat", borderwidth=0,
                                  selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                  highlightthickness=0, yscrollcommand=scroll.set)
            scroll.config(command=listbox.yview)
            listbox.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
            return listbox

        self.receivable_class_listbox = _receivable_filter_listbox_col(multiselect_filter_box, 0, "Class filter")
        self.receivable_filter_client_listbox = _receivable_filter_listbox_col(multiselect_filter_box, 1, "Client Name filter")
        self.receivable_filter_company_listbox = _receivable_filter_listbox_col(multiselect_filter_box, 2, "Company filter")

        # ---- Section: Load, Validate & Build Groups ----
        # Deliberately placed AFTER the filter section above, not before
        # it - this single action now does what used to take two separate
        # button clicks ("Load & Validate Data" then "Analyze / Build
        # Groups"), and sitting at the end lets a person set their filter
        # selections first (if they have any) before triggering the load,
        # rather than having to load, then filter, then load again.
        load_box = ttk.Labelframe(canvas_frame, text="  Load, Validate & Build Groups  ", padding=14)
        load_box.pack(fill="x", padx=12, pady=8)
        ttk.Label(load_box, text="Loads the master workbook, validates it, applies any filter selections above, "
                                  "and builds the per-company output groups shown on tab 2 - all in one step.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        self.receivable_load_analyze_btn = self._make_action_button(
            load_box, "Load, Validate & Build Groups", self.load_and_analyze_receivable_data, kind="accent")
        self.receivable_load_analyze_btn.pack(anchor="w")

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

        merge_row = ttk.Frame(f)
        merge_row.pack(fill="x", padx=12, pady=(0, 8))
        self.receivable_merge_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(merge_row, text="Merge files for the same company (combine before generating)",
                         variable=self.receivable_merge_var).pack(side="left")
        merge_note = ttk.Label(
            f, text="Combines companies sharing a leading name (e.g. \"Atlas_Insurance_Faisalabad\" / "
                    "\"Atlas_Insurance_ClaimRecovery\") or a known abbreviation (AIC, JGE/JGI, UBL). No data is lost.",
            style="Muted.TLabel", font=("Segoe UI", 9), wraplength=900, justify="left")
        merge_note.pack(anchor="w", padx=14, pady=(0, 6))

        self.receivable_reports_progress = perf_utils.ProgressBarPanel(f)
        self.receivable_reports_progress.frame.pack(fill="x", padx=12, pady=(0, 8))

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
        self.receivable_results_tree.tag_configure("oddrow", background="#F6F8FB")
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
        sig_btn = self._make_action_button(sig_btn_holder, "Save Signature", self.save_receivable_signature,
                                            kind="quiet", compact=True)
        sig_btn.pack(expand=True)

        # ---- Scrollable body ----
        # Same fix as the Claims Register tab's own email preview panel
        # (see _build_claims_emails_tab): wrapped in a canvas+scrollbar so
        # the Body field and the Save/Send buttons stay reachable on a
        # window shorter than the panel's natural height, instead of being
        # silently clipped by the window edge.
        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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

        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1, minsize=560)

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
        right.rowconfigure(8, weight=1)

        ttk.Label(right, text="To (recipient email - separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.receivable_email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Cc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.receivable_email_cc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_email_cc_var, font=("Segoe UI", 10))\
            .grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Bcc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.receivable_email_bcc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_email_bcc_var, font=("Segoe UI", 10))\
            .grid(row=5, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=6, column=0, sticky="w", pady=(0, 2))
        self.receivable_subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.receivable_subject_var, font=("Segoe UI", 10))\
            .grid(row=7, column=0, sticky="ew")

        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=8, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.receivable_body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1, height=14)
        self.receivable_body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        btn_row = ttk.Frame(right)
        btn_row.grid(row=9, column=0, sticky="w", pady=(10, 0))

        save_btn = self._make_action_button(btn_row, "Save Edits to Selected Email", self.save_receivable_email_edit, kind="accent")
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = self._make_action_button(btn_row, "Send This Email", self.send_individual_receivable_email, kind="success")
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

    def _load_receivable_filter_options(self):
        """
        'Load Filter Options' button - reads the master workbook currently
        browsed to (Data Source, above) and populates the Class / Client
        Name / Company multi-select listboxes with that file's own real
        distinct values, same "load once, populate a multi-select
        Listbox" interaction as Claims Intimation's own Unit/Client
        Name/Department filters (see _load_claims_filter_options).

        Deliberately independent of "Load, Validate & Build Groups" below
        - this reads and validates the file itself rather than requiring
        that button to have been clicked first, so the filter can be set
        up before the load/build-groups step if that's the more natural
        order for a given run. The loaded dataframe and colmap are cached
        (self.receivable_raw_df/clean_df/issues_df/colmap plus
        self.receivable_loaded_path) so "Load, Validate & Build Groups"
        can reuse this same read instead of re-reading the file from disk,
        as long as the file path hasn't changed since.
        """
        path = self.receivable_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Browse to a valid master workbook first.")
            return

        self.receivable_filter_status_var.set("Loading...")

        def worker(report):
            report(f"Reading {os.path.basename(path)}...")
            raw_df = rvp.load_raw_data(path, report=report)
            report(f"Loaded {len(raw_df):,} rows. Validating...")
            clean_df, issues_df, dropped_columns, colmap = rvp.validate_data(raw_df)
            return raw_df, clean_df, issues_df, dropped_columns, colmap

        def on_done(result):
            raw_df, clean_df, issues_df, dropped_columns, colmap = result
            self.receivable_raw_df = raw_df
            self.receivable_clean_df = clean_df
            self.receivable_issues_df = issues_df
            self.receivable_colmap = colmap
            self.receivable_loaded_path = path

            def _populate(listbox, orig_col):
                listbox.delete(0, "end")
                if not orig_col or orig_col not in clean_df.columns:
                    return
                values = sorted(clean_df[orig_col].dropna().astype(str).str.strip().unique())
                values = [v for v in values if v]
                for i, v in enumerate(values):
                    listbox.insert("end", v)
                    listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

            _populate(self.receivable_class_listbox, colmap.get("class"))
            _populate(self.receivable_filter_client_listbox, colmap.get("client"))
            _populate(self.receivable_filter_company_listbox, colmap.get("company"))

            if not colmap.get("class"):
                self.receivable_filter_status_var.set(
                    f"Loaded {len(clean_df):,} row(s). No 'Class' column found in this file - that filter is unavailable.")
            else:
                self.receivable_filter_status_var.set(f"Loaded {len(clean_df):,} row(s).")
            self.status_var.set("Receivable: filter options loaded.")

        def on_error(exc):
            self.receivable_filter_status_var.set("Load failed - see error.")
            messagebox.showerror("Could not load file", str(exc))

        perf_utils.run_in_background(
            self, worker, on_done=on_done, on_error=on_error,
            busy_widgets=[self.receivable_load_filters_btn],
        )

    def load_and_analyze_receivable_data(self):
        """
        'Load, Validate & Build Groups' button - the single action that
        replaces what used to be two separate clicks ("Load & Validate
        Data" then "Analyze / Build Groups"): loads the master workbook
        (or reuses whatever "Load Filter Options" already read, if it's
        for this same file), validates it, applies any Class/Client
        Name/Company selections made in the multi-select Filters section
        above, and builds the per-company output groups shown on tab 2 -
        all in one step.
        """
        path = self.receivable_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid master workbook.")
            return

        # Multi-select Class/Client Name/Company filters - read from the
        # listboxes here, on the MAIN thread, since Tkinter widgets must
        # never be queried from a background thread (same rule Claims
        # Intimation's own listbox filters follow - see
        # load_and_match_claims).
        selected_classes = [self.receivable_class_listbox.get(i) for i in self.receivable_class_listbox.curselection()]
        selected_clients_multi = [self.receivable_filter_client_listbox.get(i) for i in self.receivable_filter_client_listbox.curselection()]
        selected_companies_multi = [self.receivable_filter_company_listbox.get(i) for i in self.receivable_filter_company_listbox.curselection()]

        self.receivable_setup_log.delete("1.0", "end")
        self.status_var.set("Loading, validating, and building groups...")

        def worker(report):
            # Reuse the cached load from "Load Filter Options" (above) if
            # it's for this exact same file - skips reading a potentially
            # huge file from disk a second time. Falls back to loading
            # fresh here if that button was never clicked, or was clicked
            # for a different file than what's currently browsed to.
            if self.receivable_clean_df is not None and self.receivable_loaded_path == path:
                raw_df = self.receivable_raw_df
                clean_df = self.receivable_clean_df
                issues_df = self.receivable_issues_df
                colmap = self.receivable_colmap
                report(f"Using already-loaded file: {os.path.basename(path)} ({len(raw_df):,} row(s)).")
            else:
                report(f"Reading {os.path.basename(path)}...")
                raw_df = rvp.load_raw_data(path, report=report)
                report(f"Loaded {len(raw_df):,} rows. Validating...")
                clean_df, issues_df, dropped_columns, colmap = rvp.validate_data(raw_df)
                if dropped_columns:
                    report(f"Removed {len(dropped_columns)} completely empty column(s): {', '.join(dropped_columns)}")

            df = clean_df
            client_col = colmap["client"]
            company_col = colmap["company"]
            class_col = colmap.get("class")

            # Combines with AND against each other, via
            # receivable_processor.filter_receivable_by_fields.
            field_filters = []
            if selected_classes:
                if not class_col:
                    report("\nWARNING: Class filter requested, but no 'Class' column was found in this file - filter skipped.")
                else:
                    field_filters.append((class_col, selected_classes))
            if selected_clients_multi:
                field_filters.append((client_col, selected_clients_multi))
            if selected_companies_multi:
                field_filters.append((company_col, selected_companies_multi))

            if field_filters:
                before_count = len(df)
                df = rvp.filter_receivable_by_fields(df, field_filters)
                report(f"\nClass/Client Name/Company filter(s) applied: {before_count:,} -> {len(df):,} row(s) kept.")

            # analyze_buckets() computes the company/OFC bucket assignment
            # ONCE and derives groups, totals, AND every bucket's client
            # breakdown from that single pass - see receivable_processor.py.
            # Buckets here are grouped by each row's own exact Company
            # name - merging branch-name/spelling variants together is a
            # separate, optional step (the "Merge files" checkbox on tab 2).
            report("\nBuilding groups...")
            groups, totals, breakdowns = rvp.analyze_buckets(df)
            return raw_df, clean_df, issues_df, colmap, df, groups, totals, breakdowns

        def on_done(result):
            (self.receivable_raw_df, self.receivable_clean_df, self.receivable_issues_df, self.receivable_colmap,
             self.receivable_filtered_df, self.receivable_groups, self.receivable_totals,
             self.receivable_breakdowns) = result
            self.receivable_loaded_path = path

            source_sheet = self.receivable_raw_df.attrs.get("source_sheet")
            if source_sheet is not None:
                self.log_receivable(f"Auto-detected data sheet: '{source_sheet}' (out of every sheet in the workbook).")
            mode = self.receivable_colmap["mode"]
            self.log_receivable(f"Detected layout: {mode} mode "
                                 f"({'Premium Receivable + Claim Receivable amount columns' if mode == 'dual' else 'Type/Category + single Amount column'}).")
            self.log_receivable(f"Loaded {len(self.receivable_raw_df)} rows, {len(self.receivable_clean_df)} valid after validation.")
            if len(self.receivable_issues_df) > 0:
                self.log_receivable(f"{len(self.receivable_issues_df)} row(s) flagged with issues:")
                issue_counts = self.receivable_issues_df["issue"].value_counts()
                for issue_label, count in issue_counts.items():
                    self.log_receivable(f"   {count}x  {issue_label}")
            else:
                self.log_receivable("No data issues found.")

            self._refresh_receivable_results_tree()

            self.log_receivable(f"\nDone: {len(self.receivable_groups)} output file(s) identified "
                                 f"(one per insurance company, OFC-named companies pooled into one shared file).")
            self.status_var.set(f"{len(self.receivable_groups)} files. See tab 2.")

        def on_error(exc):
            messagebox.showerror("Error", str(exc))
            self.status_var.set("Load failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self.log_receivable(msg),
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.receivable_load_analyze_btn],
        )

    def _refresh_receivable_results_tree(self):
        """Repopulate tab 2's insurance-company-file tree from
        self.receivable_groups/totals - shared by Analyze (tab 1) and by
        the merge checkbox (below)."""
        self.receivable_results_tree.delete(*self.receivable_results_tree.get_children())
        for i, (bucket_label, _g) in enumerate(self.receivable_groups.items()):
            t = self.receivable_totals[bucket_label]
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.receivable_results_tree.insert("", "end", values=(
                bucket_label, t["num_clients"], f"{t['premium_total']:,.0f}", f"{t['claims_total']:,.0f}"
            ), tags=(tag,))
        self.receivable_results_summary_var.set(f"{len(self.receivable_groups)} output file(s) identified")

    def _run_receivable_self_merge(self):
        """The 'Merge files for the same company' checkbox's actual work.

        Groups self.receivable_groups' own bucket labels (each already the
        exact, unmodified Company name a row was recorded under) by
        company_alias_resolver.self_merge_bucket_labels: two labels merge
        whenever they share the same leading name before the first
        underscore/space (e.g. "Atlas_Insurance_Faisalabad" and
        "Atlas_Insurance_ClaimRecovery" both start with "Atlas"), or when
        one of them is a known abbreviation (AIC/JGE/JGI/UBL). Fully
        deterministic - nothing is ever ambiguous or held back for manual
        review. Every merged bucket's premium/claims rows are concatenated
        (combine_group below), so no data from any of the original,
        separately-named files is ever lost."""
        import company_alias_resolver as car

        company_col = self.receivable_colmap["company"]
        weights = {
            label: len(g["premium_rows"]) + len(g["claims_rows"])
            for label, g in self.receivable_groups.items()
        }
        merge_map = car.self_merge_bucket_labels(list(self.receivable_groups.keys()), weight_by=weights)

        if all(v == k for k, v in merge_map.items()):
            return True

        def combine_group(group_values, new_label):
            import pandas as pd
            premium_frames = [g["premium_rows"] for g in group_values]
            claims_frames = [g["claims_rows"] for g in group_values]
            premium = pd.concat(premium_frames, ignore_index=True)
            claims = pd.concat(claims_frames, ignore_index=True)
            if company_col in premium.columns:
                premium[company_col] = new_label
            if company_col in claims.columns:
                claims[company_col] = new_label
            return {"premium_rows": premium, "claims_rows": claims}

        merged_groups, merged_totals, merged_breakdowns = car.merge_buckets(
            self.receivable_groups, self.receivable_totals, self.receivable_breakdowns, merge_map, combine_group)

        n_before, n_after = len(self.receivable_groups), len(merged_groups)
        self.receivable_groups = merged_groups
        self.receivable_totals = merged_totals
        self.receivable_breakdowns = merged_breakdowns
        self._refresh_receivable_results_tree()
        self.log_receivable(f"\nMerge checkbox: combined {n_before} company file(s) into {n_after} by shared company name.")
        return True

    def generate_receivable_reports(self):
        if not getattr(self, "receivable_groups", None):
            messagebox.showwarning("Nothing to generate", "Run Analyze / Build Groups first (tab 1).")
            return

        if self.receivable_merge_var.get():
            if not self._run_receivable_self_merge():
                return

        base_output_dir = self.receivable_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder on tab 1 first.")
            return
        # Fresh dated/timed subfolder under the configured base folder for
        # this generation - tracked separately from the entry field so
        # repeated generations don't nest inside each other, and so Send
        # Emails (tab 3) attaches from this specific run's own files.
        output_dir = self._new_run_folder(base_output_dir)
        self.receivable_last_run_dir = output_dir

        colmap = self.receivable_colmap
        if colmap["mode"] == "dual":
            premium_amount_col, claims_amount_col = colmap["premium_amount"], colmap["claims_amount"]
        else:
            premium_amount_col = claims_amount_col = colmap["amount"]

        source_df = getattr(self, "receivable_filtered_df", self.receivable_clean_df)
        groups = self.receivable_groups
        # Precomputed once by analyze_receivable() (see analyze_buckets() in
        # receivable_processor.py) - falls back to the old per-bucket call
        # only if Analyze hasn't been re-run since an older session state,
        # which shouldn't normally happen since analyze_receivable() always
        # sets this alongside receivable_groups.
        breakdowns = getattr(self, "receivable_breakdowns", None)
        self.status_var.set(f"Generating {len(groups)} Receivable workbook(s)...")
        self.receivable_reports_progress.start(total=len(groups), text=f"0 / {len(groups)} workbook(s) generated")

        def worker(report):
            # build_company_workbooks_parallel spreads workbook generation
            # across CPU cores (each company's file is fully independent -
            # its own DataFrame slice in, its own .xlsx out) when the batch
            # is large enough to be worth it, and transparently falls back
            # to the original one-at-a-time loop otherwise (a small batch,
            # a single-core machine, or a process pool that fails to start
            # for some environment-specific reason) - see its own
            # docstring in receivable_generator.py. Per-bucket client
            # breakdowns are still taken from the precomputed
            # analyze_buckets() result when available, same as before.
            resolved_breakdowns = breakdowns
            if resolved_breakdowns is None:
                resolved_breakdowns = {
                    bucket_label: rvp.client_breakdown_for_bucket(source_df, bucket_label)
                    for bucket_label in groups
                }
            return rvg.build_company_workbooks_parallel(
                groups, resolved_breakdowns, output_dir,
                premium_amount_col=premium_amount_col, claims_amount_col=claims_amount_col,
                report=report,
            )

        def on_done(result):
            generated, filename_lookup, errors = result
            self.receivable_filename_lookup = filename_lookup
            self.receivable_reports_progress.finish(f"Done: {len(generated)} / {len(groups)} workbook(s) generated.")
            msg = f"Generated {len(generated)} workbook(s) in:\n{output_dir}"
            if errors:
                msg += f"\n\n{len(errors)} file(s) failed:\n" + "\n".join(errors[:10])
            messagebox.showinfo("Done", msg)
            self.status_var.set(f"Generated {len(generated)} Receivable workbook(s).")

        def on_error(exc):
            self.receivable_reports_progress.finish("Failed - see the Activity Log for details.")
            messagebox.showerror("Report generation failed", str(exc))
            self.status_var.set("Report generation failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.receivable_reports_progress.update,
            on_progress=lambda msg: self.log_receivable(msg),
            on_done=on_done, on_error=on_error,
        )

    def open_receivable_output_folder(self):
        output_dir = self.receivable_last_run_dir or self.receivable_output_dir_var.get().strip()
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
        self.receivable_email_cc_var.set(d.get("cc") or "")
        self.receivable_email_bcc_var.set(d.get("bcc") or "")
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
        self.receivable_drafts[idx]["cc"] = self.receivable_email_cc_var.get().strip()
        self.receivable_drafts[idx]["bcc"] = self.receivable_email_bcc_var.get().strip()
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

        output_dir = self.receivable_last_run_dir or self.receivable_output_dir_var.get().strip()

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

        output_dir = self.receivable_last_run_dir or self.receivable_output_dir_var.get().strip()
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
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run,
                                                 cc=d.get("cc"), bcc=d.get("bcc"))
            self.receivable_email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
            prefix = "[DRY RUN] " if dry_run else ""
            messagebox.showinfo("Done", f"{prefix}{message}")
            self.status_var.set(f"{prefix}{message}")

        threading.Thread(target=worker, daemon=True).start()

    # ---------- Coinsurance > Claims Register module ----------
    def _build_claims_register_page(self, parent):
        """Claims Register (Coinsurance) - matches the Claim Intimation
        Register against IGI's own Premium Register to recover each
        claim's coinsurer name(s) (a policy can have several coinsurers
        at once, so a claim can fan out to several of their files), then
        writes one Excel workbook per coinsurer and drafts/sends a
        recovery-statement email for each. Same three-tab structure as
        the Premium/Claims Email Automation page."""
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self.tab_claims_setup = ttk.Frame(notebook)
        self.tab_claims_groups = ttk.Frame(notebook)
        self.tab_claims_emails = ttk.Frame(notebook)
        notebook.add(self.tab_claims_setup, text="1. Upload & Match")
        notebook.add(self.tab_claims_groups, text="2. Coinsurer Files")
        notebook.add(self.tab_claims_emails, text="3. Review & Send Emails")

        self._build_claims_setup_tab()
        self._build_claims_groups_tab()
        self._build_claims_emails_tab()

    def _build_claims_setup_tab(self):
        f = self.tab_claims_setup
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

        ttk.Label(source_box, text="Claim Intimation Register (.xlsx/.xls)", style="Panel.TLabel").grid(row=0, column=0, sticky="w", **pad)
        self.claims_intimation_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.claims_intimation_file_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_claims_intimation_file).grid(row=0, column=2, **pad)

        ttk.Label(source_box, text="IGI Premium Register (.xlsx/.xls)", style="Panel.TLabel").grid(row=1, column=0, sticky="w", **pad)
        self.claims_premium_register_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.claims_premium_register_file_var).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_claims_premium_register_file).grid(row=1, column=2, **pad)

        claims_matching_note = ttk.Label(
            source_box, text="Matches 'Policy/PMD #' against the register's 'Base Doc No.' (fallback: 'Document No.'). "
                              "A policy shared by several coinsurers fans a claim out to each of their files.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        claims_matching_note.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(2, 2))
        self._bind_dynamic_wraplength(claims_matching_note, source_box)

        ttk.Label(source_box, text="Output folder (workbooks written here)", style="Panel.TLabel").grid(row=3, column=0, sticky="w", **pad)
        self.claims_output_dir_var = tk.StringVar(value=self.claims_output_dir)
        ttk.Entry(source_box, textvariable=self.claims_output_dir_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_claims_output_dir).grid(row=3, column=2, **pad)

        self.claims_load_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(source_box, text="Look up recipient emails from a lookup file",
                         variable=self.claims_load_email_var).grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))

        ttk.Label(source_box, text="Email Lookup File (.xlsx/.csv, optional)", style="Panel.TLabel").grid(row=5, column=0, sticky="w", **pad)
        self.claims_email_lookup_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.claims_email_lookup_file_var).grid(row=5, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_claims_email_lookup_file).grid(row=5, column=2, **pad)
        claims_email_lookup_note = ttk.Label(
            source_box, text="Company Name, Email, and optional POC columns, matched by exact coinsurer name "
                              "(pooled OFC file uses the key 'OFC'). Left blank, drafts are flagged \"missing email\".",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        claims_email_lookup_note.grid(row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 2))
        self._bind_dynamic_wraplength(claims_email_lookup_note, source_box)

        # ---- Section: Business Type ----
        # Conventional (default) or Takaful - determines whether "Takaful"
        # is worked into every generated coinsurer filename, the email
        # subject, and the email body for this run (see
        # claims_register_generator.build_output_filename and
        # claims_register_email.build_claims_register_email_drafts for the
        # exact wording). A Conventional run is worded exactly as before -
        # nothing is added.
        business_type_box = ttk.Labelframe(canvas_frame, text="  Business Type  ", padding=14)
        business_type_box.pack(fill="x", padx=12, pady=8)
        ttk.Label(business_type_box,
                  text="Applies to the coinsurer file names, the email subject, and the email body "
                       "generated for this run.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        self.claims_business_type_var = tk.StringVar(value="conventional")
        business_type_row = ttk.Frame(business_type_box)
        business_type_row.pack(fill="x")
        ttk.Radiobutton(business_type_row, text="Conventional", variable=self.claims_business_type_var,
                         value="conventional").pack(side="left", padx=(0, 24))
        ttk.Radiobutton(business_type_row, text="Takaful", variable=self.claims_business_type_var,
                         value="takaful").pack(side="left")

        # ---- Section: Date Filter ----
        # Scopes a run to a selected period even though the Claim Intimation
        # Register (and the Premium Register it's matched against) can span
        # a much wider date range. Filtering happens on the CLAIM's own date
        # field, BEFORE matching - so the matched/unmatched counts, the
        # per-coinsurer buckets, and every generated Excel workbook only
        # ever cover claims inside the selected window, rather than the
        # window being applied after the fact. The Premium Register itself
        # is never date-filtered - it's only used as a lookup table to
        # resolve each in-range claim's coinsurer(s), so it still needs to
        # be searched over its own full range regardless of the period
        # selected here.
        date_filter_box = ttk.Labelframe(canvas_frame, text="  Date Filter (optional)  ", padding=14)
        date_filter_box.pack(fill="x", padx=12, pady=8)

        ttk.Label(date_filter_box,
                  text="Leave blank to process the whole register. When set, only claims in this date range are matched.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9),
                  wraplength=900, justify="left").pack(anchor="w", pady=(0, 10))

        date_filter_row = ttk.Frame(date_filter_box)
        date_filter_row.pack(fill="x")

        ttk.Label(date_filter_row, text="Date field", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.claims_date_field_var = tk.StringVar(value="Intimation Date")
        ttk.Combobox(date_filter_row, textvariable=self.claims_date_field_var,
                     values=["Intimation Date", "Loss Date"], state="readonly", width=14).pack(side="left", padx=(0, 20))

        ttk.Label(date_filter_row, text="From (YYYY-MM-DD)", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.claims_date_from_var = tk.StringVar()
        ttk.Entry(date_filter_row, textvariable=self.claims_date_from_var, width=14).pack(side="left", padx=(0, 20))

        ttk.Label(date_filter_row, text="To (YYYY-MM-DD)", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.claims_date_to_var = tk.StringVar()
        ttk.Entry(date_filter_row, textvariable=self.claims_date_to_var, width=14).pack(side="left")

        # ---- Section: Filters (Unit / Client Name / Department) ----
        # Same style/interaction as the Outstanding Premium module's own
        # "Threshold & Filters" section (General Operations sidebar menu):
        # multi-select listboxes (Ctrl+click for several values, none
        # selected = all), populated from the real distinct values in the
        # loaded Claim Intimation Register rather than typed by hand.
        # Applied to the Claim Intimation Register BEFORE matching (same
        # point in the pipeline as the Date Filter above), so - like the
        # Date Filter - everything downstream (matched/unmatched counts,
        # coinsurer buckets, every generated workbook) only ever sees
        # claims that passed these filters. The Premium Register itself is
        # never filtered by these either, for the same reason the Date
        # Filter leaves it alone - it's only a lookup table for resolving
        # coinsurers, so it still needs its own full range/scope.
        claims_filter_box = ttk.Labelframe(canvas_frame, text="  Filters (optional)  ", padding=14)
        claims_filter_box.pack(fill="x", padx=12, pady=8)
        claims_filter_box.columnconfigure(0, weight=1)
        claims_filter_box.columnconfigure(1, weight=1)
        claims_filter_box.columnconfigure(2, weight=1)

        claims_filter_note = ttk.Label(
            claims_filter_box, text="Click \"Load Filter Options\" above to populate these lists from the file's own values.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        claims_filter_note.grid(row=0, column=0, columnspan=3, sticky="ew", padx=6, pady=(0, 8))
        self._bind_dynamic_wraplength(claims_filter_note, claims_filter_box, side_padding=10)

        load_filters_row = ttk.Frame(claims_filter_box, style="Panel.TFrame")
        load_filters_row.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 10))
        self.claims_load_filters_btn = ttk.Button(load_filters_row, text="Load Filter Options",
                                                    command=self._load_claims_filter_options)
        self.claims_load_filters_btn.pack(side="left")
        self.claims_filter_status_var = tk.StringVar(value="")
        ttk.Label(load_filters_row, textvariable=self.claims_filter_status_var, style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(side="left", padx=(10, 0))

        def _claims_filter_listbox_col(parent, col_idx, title):
            col = ttk.Frame(parent, style="Panel.TFrame")
            col.grid(row=2, column=col_idx, sticky="nsew", padx=6, pady=4)
            ttk.Label(col, text=title, style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
            ttk.Label(col, text="Ctrl+click to select multiple \u00b7 none = all", style="Muted.TLabel",
                      background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
            frame = tk.Frame(col, highlightbackground="#D6DCE1", highlightthickness=1)
            frame.pack(fill="both", expand=True)
            scroll = ttk.Scrollbar(frame, orient="vertical")
            listbox = tk.Listbox(frame, selectmode="extended", height=10, width=32,
                                  exportselection=False, font=("Segoe UI", 10),
                                  activestyle="none", relief="flat", borderwidth=0,
                                  selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                  highlightthickness=0, yscrollcommand=scroll.set)
            scroll.config(command=listbox.yview)
            listbox.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
            return listbox

        self.claims_unit_listbox = _claims_filter_listbox_col(claims_filter_box, 0, "Unit (Branch Name) filter")
        self.claims_client_listbox = _claims_filter_listbox_col(claims_filter_box, 1, "Client Name filter")
        self.claims_department_listbox = _claims_filter_listbox_col(claims_filter_box, 2, "Department filter")

        # ---- Section: Load & Match (kept below the Date Filter, so any
        # filter is already set before the user loads the data) - a proper
        # bordered Labelframe matching Data Source / Date Filter above it,
        # rather than a bare unstyled frame that read as a stray, floating
        # button with no visual home of its own. Uses the compact button
        # variant and tight spacing throughout so the card hugs its
        # content instead of reserving a lot of empty space below the
        # button for a progress bar / summary line that are blank most
        # of the time.
        load_box = ttk.Labelframe(canvas_frame, text="  Load & Match  ", padding=(14, 12))
        load_box.pack(fill="x", padx=12, pady=8)

        ttk.Label(load_box, text="Loads both files, applies the Date Filter (if set), and runs coinsurer matching.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))

        claims_load_match_btn_row = ttk.Frame(load_box)
        claims_load_match_btn_row.pack(anchor="w", fill="x")
        # Full-size (not compact) to match the visual weight of "Run
        # Matching" on Policy Check / Claims Check - all three are the
        # same kind of primary action (load + fuzzy-match two files) and
        # should read identically regardless of which page they're on.
        self.claims_match_btn = self._make_action_button(claims_load_match_btn_row, "Load & Match Claims", self.load_and_match_claims,
                                                           kind="accent", compact=False)
        self.claims_match_btn.pack(side="left")

        self.claims_match_progress = perf_utils.ProgressBarPanel(load_box)
        self.claims_match_progress.frame.pack(fill="x", pady=(8, 0))

        self.claims_match_summary_var = tk.StringVar(value="")
        self.claims_match_summary_label = ttk.Label(load_box, textvariable=self.claims_match_summary_var,
                                                      foreground=self.COLOR_SUCCESS, font=("Segoe UI", 10, "bold"))
        self.claims_match_summary_label.pack(anchor="w")
        self.claims_match_summary_label.pack_forget()  # hidden until there's a real summary to show (see on_done below)

        # ---- Section: Log ----
        log_box = ttk.Labelframe(canvas_frame, text="  Activity Log  ", padding=10)
        log_box.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.claims_setup_log = scrolledtext.ScrolledText(log_box, height=12, font=("Consolas", 9),
                                                            bg="#FBFCFD", relief="flat", borderwidth=0)
        self.claims_setup_log.pack(fill="both", expand=True)

    def _build_claims_groups_tab(self):
        f = self.tab_claims_groups

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Generate Excel Reports", style="Accent.TButton", width=24,
                   command=self.generate_claims_register_reports).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Open Output Folder", command=self.open_claims_register_output_folder, width=20).pack(side="left")
        self.claims_results_summary_var = tk.StringVar(value="Run 'Load & Match Claims' on tab 1 to see one row per coinsurer here.")
        ttk.Label(top, textvariable=self.claims_results_summary_var, style="Muted.TLabel").pack(side="right")

        hint = ttk.Label(f, text="One workbook per coinsurer (\"OFC\"-named companies pooled together). Unmatched claims "
                                  "go to a separate \"Unmatched Claims\" file.",
                          style="Muted.TLabel", font=("Segoe UI", 9), wraplength=900, justify="left")
        hint.pack(anchor="w", padx=14, pady=(0, 6))

        merge_row = ttk.Frame(f)
        merge_row.pack(fill="x", padx=12, pady=(0, 8))
        self.claims_merge_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(merge_row, text="Merge files for the same company (combine before generating)",
                         variable=self.claims_merge_var).pack(side="left")
        merge_note = ttk.Label(
            f, text="Combines coinsurers sharing a leading name (e.g. \"Atlas_Insurance_Faisalabad\" / "
                    "\"Atlas_Insurance_ClaimRecovery\") or a known abbreviation (AIC, JGE/JGI, UBL). No data is lost.",
            style="Muted.TLabel", font=("Segoe UI", 9), wraplength=900, justify="left")
        merge_note.pack(anchor="w", padx=14, pady=(0, 6))

        self.claims_reports_progress = perf_utils.ProgressBarPanel(f)
        self.claims_reports_progress.frame.pack(fill="x", padx=12, pady=(0, 8))

        tree_frame = ttk.Frame(f)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("company", "num_claims", "paid_total", "os_total")
        self.claims_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.claims_results_tree.heading("company", text="Coinsurer (output file)")
        self.claims_results_tree.heading("num_claims", text="Claims")
        self.claims_results_tree.heading("paid_total", text="Share of Claim Paid")
        self.claims_results_tree.heading("os_total", text="Share of Claim O/S")
        self.claims_results_tree.column("company", width=380)
        self.claims_results_tree.column("num_claims", width=80, anchor="center")
        self.claims_results_tree.column("paid_total", width=170, anchor="e")
        self.claims_results_tree.column("os_total", width=170, anchor="e")
        self.claims_results_tree.tag_configure("oddrow", background="#F6F8FB")
        self.claims_results_tree.tag_configure("evenrow", background="white")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.claims_results_tree.yview)
        self.claims_results_tree.configure(yscrollcommand=vsb.set)
        self.claims_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_claims_emails_tab(self):
        f = self.tab_claims_emails

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Draft Emails", style="Accent.TButton", command=self.draft_claims_register_emails, width=18).pack(side="left", padx=(0, 12))
        self.claims_dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Dry run (preview only, do not actually send)", variable=self.claims_dry_run_var).pack(side="left", padx=(0, 12))
        ttk.Button(top, text="Send All", style="Send.TButton", command=self.send_all_claims_register_emails, width=14).pack(side="left")

        sig_frame = ttk.Labelframe(f, text="  Your Signature (appears after 'Regards,' in every email - shared across every GIMS module)  ", padding=10)
        sig_frame.pack(fill="x", padx=12, pady=(0, 8))
        sig_frame.columnconfigure(0, weight=1)
        self.claims_signature_text = tk.Text(sig_frame, height=3, font=("Segoe UI", 10), relief="solid", borderwidth=1, wrap="word")
        self.claims_signature_text.grid(row=0, column=0, sticky="ew")
        self.claims_signature_text.insert("1.0", eg.load_signature())

        sig_btn_holder = tk.Frame(sig_frame, bg=self.COLOR_PANEL)
        sig_btn_holder.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        sig_btn = self._make_action_button(sig_btn_holder, "Save Signature", self.save_claims_register_signature,
                                            kind="quiet", compact=True)
        sig_btn.pack(expand=True)

        # ---- Scrollable body ----
        # The coinsurer list + email preview panel (Body textbox, Save/Send
        # buttons, etc.) used to sit in a plain fixed-height frame, so on a
        # window shorter than the panel's natural height everything below
        # "Subject" was silently clipped by the window edge, with no
        # scrollbar to reach it. Wrapped in the same canvas+scrollbar
        # pattern every other scrollable page in the app already uses (see
        # _build_claims_setup_tab) so the Body field and the Save/Send
        # buttons are always reachable regardless of window size.
        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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

        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        # A fixed minsize keeps the list/preview panel at a sensible
        # working height instead of collapsing to a cramped size just
        # because it now lives inside a scrollable canvas rather than a
        # fixed-size parent (a canvas child has nothing to expand
        # against, so weight=1 alone isn't enough here).
        body.rowconfigure(0, weight=1, minsize=560)

        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(left, text="Coinsurer Files", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        columns = ("company", "claims", "email", "status")
        self.claims_email_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=26)
        self.claims_email_tree.heading("company", text="Coinsurer")
        self.claims_email_tree.heading("claims", text="Claims")
        self.claims_email_tree.heading("email", text="Email")
        self.claims_email_tree.heading("status", text="Status")
        self.claims_email_tree.column("company", width=210)
        self.claims_email_tree.column("claims", width=60, anchor="center")
        self.claims_email_tree.column("email", width=170)
        self.claims_email_tree.column("status", width=150)
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.claims_email_tree.yview)
        self.claims_email_tree.configure(yscrollcommand=tree_scroll.set)
        self.claims_email_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.claims_email_tree.bind("<<TreeviewSelect>>", self.on_select_claims_register_email)

        right = ttk.Labelframe(body, text="  Email Preview  ", padding=14)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(8, weight=1)

        ttk.Label(right, text="To (recipient email - separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.claims_email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.claims_email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Cc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.claims_email_cc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.claims_email_cc_var, font=("Segoe UI", 10))\
            .grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Bcc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.claims_email_bcc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.claims_email_bcc_var, font=("Segoe UI", 10))\
            .grid(row=5, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=6, column=0, sticky="w", pady=(0, 2))
        self.claims_subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.claims_subject_var, font=("Segoe UI", 10))\
            .grid(row=7, column=0, sticky="ew")

        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=8, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.claims_body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1, height=14)
        self.claims_body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        btn_row = ttk.Frame(right)
        btn_row.grid(row=9, column=0, sticky="w", pady=(10, 0))

        save_btn = self._make_action_button(btn_row, "Save Edits to Selected Email", self.save_claims_register_email_edit, kind="accent")
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = self._make_action_button(btn_row, "Send This Email", self.send_individual_claims_register_email, kind="success")
        send_one_btn.pack(side="left")

    # ---------- Claims Register actions ----------
    def browse_claims_intimation_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.claims_intimation_file_var.set(path)

    def browse_claims_premium_register_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.claims_premium_register_file_var.set(path)

    def browse_claims_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.claims_output_dir_var.set(path)

    def browse_claims_email_lookup_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if path:
            self.claims_email_lookup_file_var.set(path)

    def log_claims(self, msg):
        self.claims_setup_log.insert("end", str(msg) + "\n")
        self.claims_setup_log.see("end")

    def _parse_claims_filter_date(self, var, label):
        """Parse one of the Date Filter entry fields. Returns a date object,
        or None if the field was left blank. Raises ValueError with a
        message ready to show the user if the text isn't a valid
        YYYY-MM-DD date."""
        text = var.get().strip()
        if not text:
            return None
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError(f"{label} must be in YYYY-MM-DD format (or left blank), got '{text}'.")

    def _load_claims_filter_options(self):
        """
        'Load Filter Options' button - reads the Claim Intimation Register
        currently browsed to (Data Source, above) and populates the Unit
        (Branch Name) / Client Name (Insured Name) / Department (Department
        Name) listboxes with that file's own real distinct values, same
        "load once, populate a multi-select Listbox" interaction as
        Outstanding Premium's Department/Agent/Branch filters on its own
        setup tab (see load_data's on_done there).

        The loaded dataframe, its Policy/PMD # column, and its column map
        are cached (self.claims_intimation_df / _pmd_col / _colmap /
        _loaded_path) so load_and_match_claims() can reuse them instead of
        re-reading the same file from disk - the cache is keyed on the
        file path actually loaded, so browsing to a different file and
        clicking Load & Match Claims without re-clicking this button first
        still works correctly (it just re-reads the file itself, same as
        before this feature existed).
        """
        path = self.claims_intimation_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid Claim Intimation Register file first.")
            return

        self.claims_filter_status_var.set("Loading...")

        def worker(report):
            claims_df, pmd_col = crp.load_claim_intimation_file(path, report=report)
            colmap = crp.get_claim_colmap(claims_df)
            return claims_df, pmd_col, colmap

        def on_done(result):
            claims_df, pmd_col, colmap = result
            self.claims_intimation_df = claims_df
            self.claims_intimation_pmd_col = pmd_col
            self.claims_intimation_colmap = colmap
            self.claims_intimation_loaded_path = path

            def _populate(listbox, canonical_col):
                listbox.delete(0, "end")
                orig_col = colmap.get(canonical_col)
                if not orig_col or orig_col not in claims_df.columns:
                    return
                values = sorted(claims_df[orig_col].dropna().astype(str).str.strip().unique())
                values = [v for v in values if v]
                for i, v in enumerate(values):
                    listbox.insert("end", v)
                    listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

            _populate(self.claims_unit_listbox, "Branch Name")
            _populate(self.claims_client_listbox, "Insured Name")
            _populate(self.claims_department_listbox, "Department Name")

            self.claims_filter_status_var.set(f"Loaded {len(claims_df):,} row(s).")
            self.status_var.set("Claims Intimation: filter options loaded.")

        def on_error(exc):
            self.claims_filter_status_var.set("Load failed - see error.")
            messagebox.showerror("Could not load file", str(exc))

        perf_utils.run_in_background(
            self, worker, on_done=on_done, on_error=on_error,
            busy_widgets=[self.claims_load_filters_btn],
        )

    def load_and_match_claims(self):
        intimation_path = self.claims_intimation_file_var.get().strip()
        register_path = self.claims_premium_register_file_var.get().strip()
        if not intimation_path or not os.path.exists(intimation_path):
            messagebox.showerror("Error", "Please select a valid Claim Intimation Register file.")
            return
        if not register_path or not os.path.exists(register_path):
            messagebox.showerror("Error", "Please select a valid IGI Premium Register file.")
            return

        try:
            date_filter_from = self._parse_claims_filter_date(self.claims_date_from_var, "Date Filter 'From'")
            date_filter_to = self._parse_claims_filter_date(self.claims_date_to_var, "Date Filter 'To'")
        except ValueError as e:
            messagebox.showerror("Invalid date filter", str(e))
            return
        if date_filter_from and date_filter_to and date_filter_from > date_filter_to:
            messagebox.showerror("Invalid date filter", "Date Filter 'From' must not be after 'To'.")
            return
        date_filter_field = self.claims_date_field_var.get()  # "Intimation Date" or "Loss Date"

        # Read the filter listbox selections here, on the MAIN thread -
        # Tkinter widgets must never be queried from a background thread
        # (see the identical pattern in analyze() for Outstanding Premium's
        # own Department/Agent/Branch listboxes, read before
        # run_in_background is ever called). The captured plain Python
        # lists below are what the worker closure actually uses.
        selected_units = [self.claims_unit_listbox.get(i) for i in self.claims_unit_listbox.curselection()]
        selected_clients = [self.claims_client_listbox.get(i) for i in self.claims_client_listbox.curselection()]
        selected_depts = [self.claims_department_listbox.get(i) for i in self.claims_department_listbox.curselection()]

        self.claims_setup_log.delete("1.0", "end")
        self.claims_match_summary_var.set("")
        self.claims_match_summary_label.pack_forget()
        self.status_var.set("Loading and matching claims...")
        self.claims_match_progress.start(total=None, text="Loading files...")

        def worker(report):
            # Reuse the cached load from "Load Filter Options" (above) if
            # it's for this exact same file - skips re-reading a
            # potentially huge file from disk a second time. Falls back to
            # loading fresh here (same as before this feature existed) if
            # the button was never clicked, or was clicked for a different
            # file than what's currently browsed to.
            if self.claims_intimation_df is not None and self.claims_intimation_loaded_path == intimation_path:
                claims_df = self.claims_intimation_df.copy()
                pmd_col = self.claims_intimation_pmd_col
                colmap = self.claims_intimation_colmap
                report(f"Using already-loaded Claim Intimation Register: {os.path.basename(intimation_path)} "
                       f"({len(claims_df):,} row(s)).")
            else:
                report(f"Loading Claim Intimation Register: {os.path.basename(intimation_path)}...")
                claims_df, pmd_col = crp.load_claim_intimation_file(intimation_path, report=report)
                report(f"  {len(claims_df):,} claim row(s) loaded. Policy reference column detected as '{pmd_col}'.")
                colmap = crp.get_claim_colmap(claims_df)

            claim_paid_col = colmap.get("Claim Paid")
            claim_os_col = colmap.get("Claim O/S")
            insured_col = colmap.get("Insured Name")

            # Date Filter - scopes this run (matching, output workbooks,
            # everything downstream) to the selected period, BEFORE any
            # matching happens, even though the source register can cover
            # a much wider date range. The Premium Register is deliberately
            # left unfiltered below - it's only a lookup table for
            # resolving coinsurers, so it still needs its full range.
            if date_filter_from or date_filter_to:
                date_col = colmap.get(date_filter_field)
                if not date_col:
                    report(f"\nWARNING: Date Filter requested on '{date_filter_field}', but that column "
                           f"wasn't found in the Claim Intimation Register - filter skipped, processing the "
                           f"full file instead.")
                else:
                    before_count = len(claims_df)
                    claims_df, num_no_date = crp.filter_claims_by_date(
                        claims_df, date_col, start_date=date_filter_from, end_date=date_filter_to)
                    range_text = f"{date_filter_from or '(any)'} to {date_filter_to or '(any)'}"
                    report(f"\nDate Filter applied on '{date_filter_field}' ({date_col}), range {range_text}: "
                           f"{before_count:,} -> {len(claims_df):,} claim(s) kept"
                           + (f", {num_no_date:,} row(s) excluded for a missing/unparseable date." if num_no_date else "."))

            # Unit / Client Name / Department filters - multi-select
            # listboxes (Ctrl+click for several values, none selected =
            # all), same interaction as Outstanding Premium's own
            # Department/Agent/Branch filters. Applied at the same point
            # in the pipeline as the Date Filter above (before matching),
            # so everything downstream only ever sees claims that passed
            # these filters too. Each is optional and independent; a
            # filter whose target column isn't present in this register is
            # skipped with a warning rather than silently ignored or
            # raising. selected_units/selected_clients/selected_depts were
            # already read from the listboxes on the MAIN thread above,
            # before this worker started - see the comment there for why.
            field_filters = []
            if selected_units:
                branch_col = colmap.get("Branch Name")
                if not branch_col:
                    report("\nWARNING: Unit filter requested, but no 'Branch Name' column was found in the "
                           "Claim Intimation Register - filter skipped.")
                else:
                    field_filters.append((branch_col, selected_units))
            if selected_clients:
                insured_col_for_filter = colmap.get("Insured Name")
                if not insured_col_for_filter:
                    report("\nWARNING: Client Name filter requested, but no 'Insured Name' column was found in "
                           "the Claim Intimation Register - filter skipped.")
                else:
                    field_filters.append((insured_col_for_filter, selected_clients))
            if selected_depts:
                dept_col = colmap.get("Department Name")
                if not dept_col:
                    report("\nWARNING: Department filter requested, but no 'Department Name' column was found "
                           "in the Claim Intimation Register - filter skipped.")
                else:
                    field_filters.append((dept_col, selected_depts))

            if field_filters:
                before_count = len(claims_df)
                claims_df = crp.filter_claims_by_fields(claims_df, field_filters)
                report(f"\nUnit/Client Name/Department filter(s) applied: {before_count:,} -> "
                       f"{len(claims_df):,} claim(s) kept.")

            report(f"Loading IGI Premium Register: {os.path.basename(register_path)}...")
            reg_df, ins_col, doc_col, base_doc_col, leader_col, share_col = crp.load_premium_register_file(register_path, report=report)
            report(f"  {len(reg_df):,} row(s) loaded. Coinsurer column='{ins_col}', Base Doc No column (primary)='{doc_col}'"
                   + (f", Document No column (secondary)='{base_doc_col}'" if base_doc_col else " (no Document No column found - matching against Base Doc No only)."))

            report(f"\nMatching {len(claims_df):,} claim(s) against {len(reg_df):,} register row(s)...")
            report.progress(0, len(claims_df), text=f"Matching 0 / {len(claims_df):,} claim(s)")
            matched_df, unmatched_df = crp.match_claims(
                claims_df, pmd_col, reg_df, ins_col, doc_col, base_doc_col, leader_col, share_col,
                claim_paid_col=claim_paid_col, claim_os_col=claim_os_col, report=report,
            )
            groups, totals, breakdowns = crp.build_coinsurer_buckets(matched_df, insured_col=insured_col)
            # pmd_col is the REAL column name as found in this specific
            # source file (e.g. "Policy /PMD#", "Policy/PMD #", "NUMS",
            # etc. - see CLAIM_ALIASES) - never a guessed/hardcoded
            # literal. It has to be threaded back out here so on_done and
            # send_all_claims_register_emails() can look up the same
            # column on matched_df/claims_matched_df later without
            # re-guessing it (a hardcoded literal here previously caused
            # a KeyError crash against real files whose header text
            # didn't happen to match the guess).
            return claims_df, pmd_col, matched_df, unmatched_df, groups, totals, breakdowns

        def on_done(result):
            claims_df, pmd_col, matched_df, unmatched_df, groups, totals, breakdowns = result
            self.claims_pmd_col = pmd_col
            self.claims_matched_df = matched_df
            self.claims_unmatched_df = unmatched_df
            self.claims_groups = groups
            self.claims_totals = totals
            self.claims_breakdowns = breakdowns

            n_claims = len(claims_df)
            n_matched_pairs = len(matched_df)
            n_unique_matched = matched_df[pmd_col].nunique() if len(matched_df) else 0
            n_unmatched = len(unmatched_df)

            self.log_claims(f"\n{n_unique_matched} of {n_claims} claim(s) matched to a coinsurer "
                             f"({n_matched_pairs} claim-coinsurer pair(s) across {len(groups)} coinsurer(s) - "
                             f"a claim shared by several coinsurers counts once per coinsurer).")
            if n_unmatched:
                self.log_claims(f"{n_unmatched} claim(s) need manual review:")
                reason_counts = unmatched_df["Reason"].value_counts()
                for reason, count in reason_counts.items():
                    self.log_claims(f"   {count}x  {reason}")

            self._refresh_claims_results_tree()

            summary_text = f"Matched {n_unique_matched}/{n_claims} claims \u2192 {len(groups)} coinsurer file(s). {n_unmatched} unmatched."
            self.claims_match_summary_var.set(summary_text)
            self.claims_match_summary_label.pack(anchor="w", pady=(8, 0))
            # The summary label above already shows the final result, so
            # the progress bar's own "Done" card is redundant the moment
            # it appears - collapse it immediately instead of leaving it
            # (and the vertical space it reserves) visible for its normal
            # ~2.6s auto-hide delay, which otherwise leaves a blank gap
            # between the button and the summary text.
            self.claims_match_progress.finish(summary_text)
            self.claims_match_progress.hide()
            self.status_var.set(f"{len(groups)} coinsurer files identified. See tab 2.")

        def on_error(exc):
            self.claims_match_progress.finish("Failed - see the Activity Log for details.")
            self.log_claims(f"\nERROR: {exc}")
            messagebox.showerror("Matching failed", str(exc))
            self.status_var.set("Claims matching failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self.log_claims(msg),
            on_progress_value=self.claims_match_progress.update,
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.claims_match_btn],
        )

    def _refresh_claims_results_tree(self):
        """Repopulate tab 2's coinsurer-file tree from self.claims_groups/
        totals - shared by 'Load & Match Claims' (tab 1) and by the merge
        checkbox (below), so both end up showing the exact same bucket
        list the next 'Generate Excel Reports' click will actually write."""
        self.claims_results_tree.delete(*self.claims_results_tree.get_children())
        for i, (bucket_label, _g) in enumerate(self.claims_groups.items()):
            t = self.claims_totals[bucket_label]
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.claims_results_tree.insert("", "end", values=(
                bucket_label, t["num_claims"], f"{t['total_share_paid']:,.0f}", f"{t['total_share_os']:,.0f}"
            ), tags=(tag,))
        self.claims_results_summary_var.set(f"{len(self.claims_groups)} coinsurer file(s) identified")

    def _run_claims_self_merge(self):
        """
        The 'Merge files for the same company' checkbox's actual work.

        Groups self.claims_groups' own bucket labels (each already the
        exact, unmodified Coinsurer Name a claim was matched under) by
        company_alias_resolver.self_merge_bucket_labels: two labels merge
        whenever they share the same leading name before the first
        underscore/space (e.g. "Atlas_Insurance_Faisalabad" and
        "Atlas_Insurance_ClaimRecovery" both start with "Atlas"), or when
        one of them is a known abbreviation (AIC = Adamjee, JGE/JGI =
        Jubilee, UBL = United Bank). Fully deterministic - no manual
        review, nothing ever left ambiguous. Every merged bucket's claims
        are concatenated (combine_group below), so no data from any of
        the original, separately-named files is ever lost.

        Always returns True (there's nothing that can block generation).
        """
        import company_alias_resolver as car

        if not self.claims_groups:
            return True

        weights = {label: self.claims_totals[label]["num_claims"] for label in self.claims_groups}
        merge_map = car.self_merge_bucket_labels(list(self.claims_groups.keys()), weight_by=weights)

        # Only bother re-bucketing if the merge actually combines something.
        if all(v == k for k, v in merge_map.items()):
            return True

        def combine_group(dfs, new_label):
            import pandas as pd
            combined = pd.concat(dfs, ignore_index=True)
            combined["Coinsurer Name"] = new_label
            return combined

        merged_groups, merged_totals, merged_breakdowns = car.merge_buckets(
            self.claims_groups, self.claims_totals, self.claims_breakdowns, merge_map, combine_group)

        n_before, n_after = len(self.claims_groups), len(merged_groups)
        self.claims_groups, self.claims_totals, self.claims_breakdowns = merged_groups, merged_totals, merged_breakdowns
        self._refresh_claims_results_tree()
        self.log_claims(f"\nMerge checkbox: combined {n_before} coinsurer file(s) into {n_after} by shared company name.")
        return True

    def generate_claims_register_reports(self):
        if not self.claims_groups:
            messagebox.showwarning("Nothing to generate", "Run 'Load & Match Claims' first (tab 1).")
            return

        if self.claims_merge_var.get():
            if not self._run_claims_self_merge():
                return

        base_output_dir = self.claims_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder on tab 1 first.")
            return
        # Fresh dated/timed subfolder under the configured base folder for
        # this generation - tracked separately so repeated generations
        # don't nest, and so Send Emails (tab 3) attaches from this run.
        output_dir = self._new_run_folder(base_output_dir)
        self.claims_last_run_dir = output_dir

        groups = self.claims_groups
        breakdowns = self.claims_breakdowns
        unmatched_df = self.claims_unmatched_df
        # Read on the main thread, same rule as every other Tkinter widget
        # value this worker closure captures below - Tkinter variables
        # must never be queried from a background thread.
        business_type = self.claims_business_type_var.get()
        self.status_var.set(f"Generating {len(groups)} Claims Register workbook(s)...")
        total_files = len(groups) + (1 if unmatched_df is not None and len(unmatched_df) else 0)
        self.claims_reports_progress.start(total=total_files, text=f"0 / {total_files} workbook(s) generated")

        def worker(report):
            # build_coinsurer_workbooks_parallel spreads workbook generation
            # across CPU cores - each coinsurer's file is fully independent
            # (its own DataFrame slice in, its own .xlsx out), and this
            # module's bucket sizes are far more skewed than Receivable's
            # (one real coinsurer can carry tens of thousands of matched
            # claim-coinsurer rows while most carry a handful) - see its
            # own docstring in claims_register_generator.py. The Unmatched
            # Claims workbook (if any) is folded into this SAME batch as
            # one more independent task, so it runs concurrently with
            # every coinsurer's workbook instead of as a separate,
            # strictly-sequential step afterward that would otherwise
            # leave every other CPU core idle while it alone was written.
            # Falls back to the original one-at-a-time loop on a
            # single-core machine, a tiny batch, or if the process pool
            # can't start for some environment-specific reason.
            return crg.build_coinsurer_workbooks_parallel(
                groups, breakdowns, output_dir, report=report, unmatched_df=unmatched_df,
                business_type=business_type)

        def on_done(result):
            generated, filename_lookup, errors = result
            self.claims_filename_lookup = filename_lookup
            self.claims_reports_progress.finish(f"Done: {len(generated)} / {total_files} workbook(s) generated.")
            msg = f"Generated {len(generated)} workbook(s) in:\n{output_dir}"
            if errors:
                msg += f"\n\n{len(errors)} file(s) failed:\n" + "\n".join(errors[:10])
            messagebox.showinfo("Done", msg)
            self.status_var.set(f"Generated {len(generated)} Claims Register workbook(s).")

        def on_error(exc):
            self.claims_reports_progress.finish("Failed - see the Activity Log for details.")
            messagebox.showerror("Report generation failed", str(exc))
            self.status_var.set("Report generation failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.claims_reports_progress.update,
            on_progress=lambda msg: self.log_claims(msg),
            on_done=on_done, on_error=on_error,
        )

    def open_claims_register_output_folder(self):
        output_dir = self.claims_last_run_dir or self.claims_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate reports first.")

    def draft_claims_register_emails(self):
        if not self.claims_groups:
            messagebox.showwarning("Nothing to draft", "Run 'Load & Match Claims' first (tab 1).")
            return

        self.claims_email_lookup, self.claims_poc_lookup = {}, {}
        if self.claims_load_email_var.get():
            lookup_path = self.claims_email_lookup_file_var.get().strip()
            if lookup_path and os.path.exists(lookup_path):
                try:
                    self.claims_email_lookup, self.claims_poc_lookup = ra.load_email_lookup_from_file(lookup_path)
                    self.log_claims(f"\nLoaded {len(self.claims_email_lookup)} coinsurer email mapping(s) "
                                     f"from '{os.path.basename(lookup_path)}'.")
                except Exception as e:
                    self.log_claims(f"\nWARNING: could not load email lookup: {e}")
            else:
                self.log_claims("\nNo email lookup file selected - all drafts will show 'Missing email'.")

        signature = self.claims_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)  # shared signature store across every GIMS module

        period_text = self._format_period_text(self.claims_date_from_var, self.claims_date_to_var)
        self.claims_drafts = cre.build_claims_register_email_drafts(
            self.claims_groups, self.claims_totals, self.claims_email_lookup,
            poc_lookup=self.claims_poc_lookup, signature=signature, period_text=period_text,
            business_type=self.claims_business_type_var.get(),
        )
        self.claims_email_tree.delete(*self.claims_email_tree.get_children())
        for d in self.claims_drafts:
            self.claims_email_tree.insert("", "end", values=(
                d["company_name"], d["num_claims"], d["email"] or "(missing)", d["status"]
            ))
        self.status_var.set(f"Drafted {len(self.claims_drafts)} Claims Register emails.")

    def save_claims_register_signature(self):
        signature = self.claims_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)
        self.status_var.set("Signature saved (shared across every GIMS module). Click Draft Emails again to refresh existing drafts.")

    def on_select_claims_register_email(self, event):
        sel = self.claims_email_tree.selection()
        if not sel:
            return
        idx = self.claims_email_tree.index(sel[0])
        d = self.claims_drafts[idx]
        self.claims_email_to_var.set(d.get("email") or "")
        self.claims_email_cc_var.set(d.get("cc") or "")
        self.claims_email_bcc_var.set(d.get("bcc") or "")
        self.claims_subject_var.set(d["subject"])
        self.claims_body_text.delete("1.0", "end")
        self.claims_body_text.insert("1.0", d["body"])

    def save_claims_register_email_edit(self):
        sel = self.claims_email_tree.selection()
        if not sel:
            messagebox.showinfo("No coinsurer selected", "Select a coinsurer from the list on the left first.")
            return
        idx = self.claims_email_tree.index(sel[0])
        new_email = self.claims_email_to_var.get().strip()
        self.claims_drafts[idx]["email"] = new_email or None
        self.claims_drafts[idx]["cc"] = self.claims_email_cc_var.get().strip()
        self.claims_drafts[idx]["bcc"] = self.claims_email_bcc_var.get().strip()
        self.claims_drafts[idx]["subject"] = self.claims_subject_var.get()
        self.claims_drafts[idx]["body"] = self.claims_body_text.get("1.0", "end").strip()

        status = "Ready" if new_email else "Missing email - needs manual entry"
        row_id = self.claims_email_tree.get_children()[idx]
        self.claims_email_tree.set(row_id, "email", new_email or "(missing)")
        self.claims_email_tree.set(row_id, "status", status)

        self.status_var.set(f"Updated draft for {self.claims_drafts[idx]['company_name']}.")

    def send_all_claims_register_emails(self):
        if not self.claims_drafts:
            messagebox.showwarning("Nothing to send", "Draft emails first.")
            return

        filename_lookup = self.claims_filename_lookup
        missing = [d["company_name"] for d in self.claims_drafts
                   if d["company_name"] not in filename_lookup]
        if missing:
            proceed = messagebox.askyesno(
                "Reports not generated",
                f"{len(missing)} of {len(self.claims_drafts)} file(s) have no generated Excel workbook yet, "
                f"so those emails would go out with no attachment:\n\n"
                + "\n".join(missing[:10]) + (f"\n...and {len(missing) - 10} more" if len(missing) > 10 else "")
                + "\n\nGo to tab 2 and click 'Generate Excel Reports' first, "
                  "or continue anyway without attachments?\n\nContinue anyway?"
            )
            if not proceed:
                return

        dry_run = self.claims_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send {len(self.claims_drafts)} emails via Outlook now? This cannot be undone."):
                return

        output_dir = self.claims_last_run_dir or self.claims_output_dir_var.get().strip()

        def worker():
            results = cre.send_claims_register_emails(
                self.claims_drafts, output_dir, filename_lookup, dry_run=dry_run
            )
            sent = sum(1 for r in results if r["sent"])
            failed = len(results) - sent
            missing_email = sum(1 for d in self.claims_drafts if not d.get("email"))
            # pmd_col: the real policy-reference column name found in this
            # run's own source file (set in load_and_match_claims' on_done),
            # never a hardcoded guess like the literal "Policy/PMD #" used
            # here previously - that guess doesn't match the confirmed real
            # header text ("Policy /PMD#", space before the slash) and
            # would raise a KeyError here, crashing Send All after emails
            # had potentially already started sending.
            pmd_col = getattr(self, "claims_pmd_col", None)
            crlog.log_run(
                num_claims_loaded=(len(self.claims_matched_df[pmd_col].unique()) + len(self.claims_unmatched_df)) if self.claims_matched_df is not None and pmd_col else 0,
                num_matched_claims=self.claims_matched_df[pmd_col].nunique() if self.claims_matched_df is not None and pmd_col else 0,
                num_unmatched_claims=len(self.claims_unmatched_df) if self.claims_unmatched_df is not None else 0,
                num_groups=len(self.claims_groups), num_files_generated=len(filename_lookup),
                num_sent=sent, num_failed=failed, num_missing_email=missing_email, dry_run=dry_run,
            )
            for i, r in enumerate(results):
                self.claims_email_tree.set(self.claims_email_tree.get_children()[i], "status",
                                            "Sent" if r["sent"] else f"Failed: {r['send_message']}")
            messagebox.showinfo("Done", f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")
            self.status_var.set(f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")

        threading.Thread(target=worker, daemon=True).start()

    def send_individual_claims_register_email(self):
        sel = self.claims_email_tree.selection()
        if not sel:
            messagebox.showinfo("No coinsurer selected", "Click a coinsurer in the list on the left first.")
            return
        idx = self.claims_email_tree.index(sel[0])

        self.save_claims_register_email_edit()
        d = self.claims_drafts[idx]

        if not d.get("email"):
            messagebox.showerror("Missing email", "Enter a recipient email address in the 'To' field before sending.")
            return

        output_dir = self.claims_last_run_dir or self.claims_output_dir_var.get().strip()
        fname = self.claims_filename_lookup.get(d["company_name"])
        attachment_path = os.path.join(output_dir, fname) if fname and os.path.exists(os.path.join(output_dir, fname)) else None
        if attachment_path is None:
            if not messagebox.askyesno(
                "No report generated",
                f"No Excel workbook has been generated yet for {d['company_name']}. "
                f"Send this email without an attachment?\n\n"
                f"(Go to tab 2 and click 'Generate Excel Reports' first if you'd rather attach one.)"
            ):
                return

        dry_run = self.claims_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send this email to {d['email']} ({d['company_name']}) now?"):
                return

        row_id = self.claims_email_tree.get_children()[idx]

        def worker():
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run,
                                                 cc=d.get("cc"), bcc=d.get("bcc"))
            self.claims_email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
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

        self.analyze_btn = ttk.Button(filter_box, text="Analyze / Apply Filters", style="Accent.TButton", width=24,
                   command=self.analyze)
        self.analyze_btn.grid(row=3, column=0, sticky="w", padx=6, pady=(12, 4))

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

        # Progress bar for "Generate Excel Reports" - stays packed in place
        # (never appears/disappears, just resets to empty between runs) and
        # is driven purely off the background worker thread's own progress
        # reports, so it never touches or slows down the generation itself.
        self.reports_progress = perf_utils.ProgressBarPanel(f)
        self.reports_progress.frame.pack(fill="x", padx=12, pady=(0, 8))

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
        self.results_tree.tag_configure("oddrow", background="#F6F8FB")
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
        sig_btn = self._make_action_button(sig_btn_holder, "Save Signature", self.save_signature,
                                            kind="quiet", compact=True)
        sig_btn.pack(expand=True)

        # ---- Scrollable body ----
        # Same fix as the Coinsurance modules' own email preview panels
        # (see _build_claims_emails_tab): wrapped in a canvas+scrollbar so
        # the Body field and the Save/Send buttons stay reachable on a
        # window shorter than the panel's natural height, instead of being
        # silently clipped by the window edge.
        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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

        # Give the list a fixed proportion and the preview panel the rest,
        # using grid so the split stays consistent as the window resizes.
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1, minsize=560)

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
        right.rowconfigure(8, weight=1)  # body text area gets the extra vertical space

        ttk.Label(right, text="To (recipient email)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Cc", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.email_cc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.email_cc_var, font=("Segoe UI", 10))\
            .grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Bcc", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.email_bcc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.email_bcc_var, font=("Segoe UI", 10))\
            .grid(row=5, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=6, column=0, sticky="w", pady=(0, 2))
        self.subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.subject_var, font=("Segoe UI", 10))\
            .grid(row=7, column=0, sticky="ew")

        # Body label + text area share the remaining vertical space
        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=8, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1, height=14)
        self.body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        # Plain tk.Button (not ttk) for this action: on some Windows ttk
        # themes, ttk.Button silently ignores the 'width' option for bold
        # fonts and clips the label. A classic tk.Button always sizes
        # itself exactly to its own text, so the full label is guaranteed
        # to render.
        btn_row = ttk.Frame(right)
        btn_row.grid(row=9, column=0, sticky="w", pady=(10, 0))

        save_btn = self._make_action_button(btn_row, "Save Edits to Selected Email", self.save_email_edit, kind="accent")
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = self._make_action_button(btn_row, "Send This Email", self.send_individual_email, kind="success")
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

        self.setup_log.delete("1.0", "end")
        self.status_var.set("Loading data...")

        email_path = self.email_file_var.get().strip()
        sheet_name = self.sheet_name_var.get()
        date_col = self.date_col_var.get()

        def worker(report):
            report(f"Reading {os.path.basename(path)}...")
            raw_df = dp.load_raw_data(path, sheet_name=sheet_name, report=report)
            report(f"Loaded {len(raw_df):,} rows. Validating...")
            clean_df, issues_df, dropped_columns = dp.validate_data(raw_df, date_col_for_year=date_col)

            email_lookup, poc_lookup, email_warning = {}, {}, None
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
                    email_lookup = dict(zip(edf[name_col].astype(str).str.strip(), edf[email_col].astype(str).str.strip()))

                    poc_col = next((c for c in edf.columns
                                     if c != email_col and c != name_col
                                     and any(k in c.lower() for k in ["poc", "contact", "attn", "attention", "recipient name", "contact person"])),
                                    None)
                    if poc_col:
                        for client_raw, poc_raw in zip(edf[name_col], edf[poc_col]):
                            if pd.isna(client_raw) or pd.isna(poc_raw):
                                continue
                            client_name = str(client_raw).strip()
                            poc_name = str(poc_raw).strip()
                            if client_name and poc_name:
                                poc_lookup[client_name] = poc_name
                except Exception as e:
                    email_warning = str(e)

            return raw_df, clean_df, issues_df, dropped_columns, email_lookup, poc_lookup, email_warning

        def on_done(result):
            self.raw_df, self.clean_df, self.issues_df, dropped_columns, self.email_lookup, self.poc_lookup, email_warning = result

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

            if self.email_lookup:
                self.log(self.setup_log, f"Loaded {len(self.email_lookup)} client-email mappings from lookup file.")
                if self.poc_lookup:
                    self.log(self.setup_log, f"{len(self.poc_lookup)} client(s) will be greeted by contact name.")
            if email_warning:
                self.log(self.setup_log, f"Warning: could not load email lookup file: {email_warning}")

            self.status_var.set("Data loaded. Set threshold/filters and click Analyze.")

        def on_error(exc):
            messagebox.showerror("Error loading file", str(exc))
            self.status_var.set("Load failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self.log(self.setup_log, msg),
            on_done=on_done, on_error=on_error,
        )

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

        clean_df = self.clean_df
        self.status_var.set("Analyzing...")

        def worker(report):
            filtered_df, grand_totals = dp.build_pivot(
                clean_df, company_filter=company_filter,
                department_filter=selected_depts or None,
                agent_filter=selected_agents or None,
                branch_filter=selected_branches or None,
            )
            qualifying = dp.qualifying_clients(filtered_df, grand_totals, threshold)
            # Single bulk pass over filtered_df for every qualifying
            # client's agent(s)/branch(es), instead of calling
            # client_agents()/client_branches() once per client in the
            # treeview-population loop below (each of which used to
            # rescan the whole filtered dataset) - see
            # data_processor.client_agents_branches_bulk().
            agents_by_client, branches_by_client = dp.client_agents_branches_bulk(filtered_df)
            return filtered_df, grand_totals, qualifying, agents_by_client, branches_by_client

        def on_done(result):
            self.filtered_df, self.grand_totals, self.qualifying, agents_by_client, branches_by_client = result

            self.results_tree.delete(*self.results_tree.get_children())
            for i, (client, total) in enumerate(self.qualifying.items()):
                agents = ", ".join(agents_by_client.get(client, []))
                branches = ", ".join(branches_by_client.get(client, []))
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                self.results_tree.insert("", "end", values=(client, f"{total:,.0f}", agents, branches), tags=(tag,))
            self.results_summary_var.set(f"{len(self.qualifying)} client(s) meet/exceed threshold {threshold:,.0f}")

            self.log(self.setup_log, f"\nAnalysis complete: {len(self.qualifying)} clients meet/exceed threshold {threshold:,.0f}.")
            self.status_var.set(f"{len(self.qualifying)} qualifying clients. See tab 2.")

        def on_error(exc):
            messagebox.showerror("Analysis failed", str(exc))
            self.status_var.set("Analysis failed.")

        perf_utils.run_in_background(
            self, worker, on_done=on_done, on_error=on_error,
            busy_widgets=[self.analyze_btn],
        )

    def generate_reports(self):
        if self.qualifying is None or len(self.qualifying) == 0:
            messagebox.showwarning("Nothing to generate", "Run Analyze first (tab 1).")
            return

        clients = list(self.qualifying.index)
        filtered_df = self.filtered_df
        # Fresh dated/timed subfolder under the module's own base folder for
        # this generation - self.output_dir is updated to point at it so
        # Draft/Send Emails (tab 3) attach from this run's own files.
        output_dir = self._new_run_folder(self.output_dir_base)
        self.output_dir = output_dir
        self.status_var.set(f"Generating {len(clients)} report(s)...")
        self.reports_progress.start(total=len(clients), text=f"0 / {len(clients)} report(s) generated")

        def worker(report):
            # win32com requires COM to be initialized on whichever thread
            # actually calls it - the main thread gets this for free from
            # Tkinter's own message loop, a plain background thread does
            # not, so it's done explicitly here and torn down in finally.
            try:
                import pythoncom
                pythoncom.CoInitialize()
                com_initialized = True
            except ImportError:
                com_initialized = False

            # Split filtered_df into one sub-dataframe per client in a
            # SINGLE pass up front, instead of the loop below calling
            # get_client_raw_rows()/client_pivot_table() once per client
            # (each of which used to rescan/refilter the whole - possibly
            # huge - filtered_df) - see data_processor.py for the
            # before/after benchmark in the docstrings.
            report("Grouping raw rows by client...")
            rows_by_client = dp.group_raw_rows_by_client(filtered_df)

            use_native = pivot_native.EXCEL_AVAILABLE
            excel_app = None
            generated = []
            native_failures = []
            try:
                if use_native:
                    try:
                        excel_app = pivot_native.start_excel()
                    except Exception as e:
                        use_native = False
                        report(f"Could not start Excel for native PivotTables ({e}) - using static formatted tables instead.")

                if use_native:
                    # The native pivot path drives a single shared Excel COM
                    # instance across every client - COM automation isn't
                    # safe to parallelize the same way plain openpyxl
                    # writing is (see report_generator.write_client_reports_parallel's
                    # own docstring), so this stays a sequential loop, just
                    # over a COM instance that now also has ScreenUpdating
                    # and automatic recalculation turned off for the
                    # duration (see pivot_native.start_excel) - a standard
                    # speedup for batch COM automation.
                    for i, client in enumerate(clients, start=1):
                        raw_rows = rows_by_client.get(client)
                        if raw_rows is None:  # shouldn't normally happen - safety net only
                            raw_rows = dp.get_client_raw_rows(filtered_df, client)
                        report(f"[{i}/{len(clients)}] {client} ({len(raw_rows):,} row(s))")
                        report.progress(i - 1, len(clients), text=f"Generating: {client} ({i}/{len(clients)})")
                        fp = None
                        try:
                            fp = rg.write_raw_only_workbook(client, raw_rows, output_dir, report=report)
                            pivot_native.add_native_pivot_table(fp, client_name=client, excel_app=excel_app)
                        except Exception:
                            native_failures.append(client)
                            pivot_df = dp.client_pivot_table_from_rows(raw_rows)
                            fp = rg.write_client_report(client, pivot_df, output_dir, raw_df=raw_rows, report=report)
                        generated.append(fp)
                        report.progress(i, len(clients), text=f"{i} / {len(clients)} report(s) generated")
                else:
                    # No Excel available - every client's report is a
                    # plain openpyxl write, fully independent of every
                    # other, so this path is safely parallelized across
                    # CPU cores (falls back to sequential on its own for a
                    # single-core machine or a small batch - see
                    # write_client_reports_parallel's own docstring).
                    generated, gen_errors = rg.write_client_reports_parallel(
                        clients, rows_by_client, output_dir, report=report)
                    if gen_errors:
                        for err in gen_errors:
                            report(f"  ERROR: {err}")
            finally:
                pivot_native.stop_excel(excel_app)
                if com_initialized:
                    pythoncom.CoUninitialize()

            return generated, use_native, native_failures

        def on_done(result):
            generated, use_native, native_failures = result
            self.reports_progress.finish(f"Done: {len(generated)} / {len(clients)} report(s) generated.")
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

        def on_error(exc):
            self.reports_progress.finish("Failed - see the Activity Log for details.")
            messagebox.showerror("Report generation failed", str(exc))
            self.status_var.set("Report generation failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self.log(self.setup_log, msg),
            on_progress_value=self.reports_progress.update,
            on_done=on_done, on_error=on_error,
        )

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
        self.email_cc_var.set(d.get("cc") or "")
        self.email_bcc_var.set(d.get("bcc") or "")
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
        self.drafts[idx]["cc"] = self.email_cc_var.get().strip()
        self.drafts[idx]["bcc"] = self.email_bcc_var.get().strip()
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
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run,
                                                 cc=d.get("cc"), bcc=d.get("bcc"))
            self.email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
            prefix = "[DRY RUN] " if dry_run else ""
            messagebox.showinfo("Done", f"{prefix}{message}")
            self.status_var.set(f"{prefix}{message}")

        threading.Thread(target=worker, daemon=True).start()


    # ---------- Coinsurance > Premium > Premium Register (company split) module ----------
    def _build_premium_register_split_page(self, parent):
        """Premium Register (Coinsurance company split) - splits IGI's own
        raw Premium Register export by coinsurer and drafts/sends a
        statement email per company file. Same three-tab structure as
        Claims Intimation / Premium-Claims Email Automation."""
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        self.tab_premregsplit_setup = ttk.Frame(notebook)
        self.tab_premregsplit_groups = ttk.Frame(notebook)
        self.tab_premregsplit_emails = ttk.Frame(notebook)
        notebook.add(self.tab_premregsplit_setup, text="1. Upload & Build Groups")
        notebook.add(self.tab_premregsplit_groups, text="2. Company Files")
        notebook.add(self.tab_premregsplit_emails, text="3. Review & Send Emails")

        self._build_premregsplit_setup_tab()
        self._build_premregsplit_groups_tab()
        self._build_premregsplit_emails_tab()

    def _build_premregsplit_setup_tab(self):
        f = self.tab_premregsplit_setup
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

        ttk.Label(source_box, text="Premium Register File (.xlsx/.xls)", style="Panel.TLabel").grid(row=0, column=0, sticky="w", **pad)
        self.premregsplit_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.premregsplit_file_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_premregsplit_file).grid(row=0, column=2, **pad)

        premregsplit_note = ttk.Label(
            source_box, text="Only Leader = \"No\" rows (a coinsurer's own share) are included; IGI's own rows are excluded.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premregsplit_note.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(2, 2))
        self._bind_dynamic_wraplength(premregsplit_note, source_box)

        ttk.Label(source_box, text="Output folder (workbooks written here)", style="Panel.TLabel").grid(row=2, column=0, sticky="w", **pad)
        self.premregsplit_output_dir_var = tk.StringVar(value=self.premregsplit_output_dir)
        ttk.Entry(source_box, textvariable=self.premregsplit_output_dir_var).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_premregsplit_output_dir).grid(row=2, column=2, **pad)

        self.premregsplit_load_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(source_box, text="Look up recipient emails from a lookup file",
                         variable=self.premregsplit_load_email_var).grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))

        ttk.Label(source_box, text="Email Lookup File (.xlsx/.csv, optional)", style="Panel.TLabel").grid(row=4, column=0, sticky="w", **pad)
        self.premregsplit_email_lookup_file_var = tk.StringVar()
        ttk.Entry(source_box, textvariable=self.premregsplit_email_lookup_file_var).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Button(source_box, text="Browse...", command=self.browse_premregsplit_email_lookup_file).grid(row=4, column=2, **pad)
        premregsplit_email_lookup_note = ttk.Label(
            source_box, text="Company Name, Email, and optional POC columns, matched by exact company name "
                              "(pooled OFC file uses the key 'OFC'). Left blank, drafts are flagged \"missing email\".",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premregsplit_email_lookup_note.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 2))
        self._bind_dynamic_wraplength(premregsplit_email_lookup_note, source_box)

        # ---- Section: Business Type ----
        # Conventional (default) or Takaful - determines whether "Takaful"
        # is worked into every generated company filename, the email
        # subject, and the email body for this run (see
        # premium_register_company_generator.build_output_filename and
        # premium_register_company_email.build_premium_register_email_drafts
        # for the exact wording). A Conventional run is worded exactly as
        # before - nothing is added.
        premregsplit_business_type_box = ttk.Labelframe(canvas_frame, text="  Business Type  ", padding=14)
        premregsplit_business_type_box.pack(fill="x", padx=12, pady=8)
        ttk.Label(premregsplit_business_type_box,
                  text="Applies to the company file names, the email subject, and the email body "
                       "generated for this run.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        self.premregsplit_business_type_var = tk.StringVar(value="conventional")
        premregsplit_business_type_row = ttk.Frame(premregsplit_business_type_box)
        premregsplit_business_type_row.pack(fill="x")
        ttk.Radiobutton(premregsplit_business_type_row, text="Conventional", variable=self.premregsplit_business_type_var,
                         value="conventional").pack(side="left", padx=(0, 24))
        ttk.Radiobutton(premregsplit_business_type_row, text="Takaful", variable=self.premregsplit_business_type_var,
                         value="takaful").pack(side="left")

        # ---- Section: Date Filter ----
        date_filter_box = ttk.Labelframe(canvas_frame, text="  Date Filter (optional)  ", padding=14)
        date_filter_box.pack(fill="x", padx=12, pady=8)

        ttk.Label(date_filter_box,
                  text="Leave blank to process the whole register. When set, only rows in this date range are included.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9),
                  wraplength=900, justify="left").pack(anchor="w", pady=(0, 10))

        date_filter_row = ttk.Frame(date_filter_box)
        date_filter_row.pack(fill="x")

        ttk.Label(date_filter_row, text="Date field", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.premregsplit_date_field_var = tk.StringVar(value="Issue Date")
        ttk.Combobox(date_filter_row, textvariable=self.premregsplit_date_field_var,
                     values=["Issue Date", "Comm Date"], state="readonly", width=14).pack(side="left", padx=(0, 20))

        ttk.Label(date_filter_row, text="From (YYYY-MM-DD)", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.premregsplit_date_from_var = tk.StringVar()
        ttk.Entry(date_filter_row, textvariable=self.premregsplit_date_from_var, width=14).pack(side="left", padx=(0, 20))

        ttk.Label(date_filter_row, text="To (YYYY-MM-DD)", style="Panel.TLabel").pack(side="left", padx=(0, 6))
        self.premregsplit_date_to_var = tk.StringVar()
        ttk.Entry(date_filter_row, textvariable=self.premregsplit_date_to_var, width=14).pack(side="left")

        # ---- Section: Filters (Business Class / Insured) ----
        premregsplit_filter_box = ttk.Labelframe(canvas_frame, text="  Filters (optional)  ", padding=14)
        premregsplit_filter_box.pack(fill="x", padx=12, pady=8)
        premregsplit_filter_box.columnconfigure(0, weight=1)
        premregsplit_filter_box.columnconfigure(1, weight=1)

        premregsplit_filter_note = ttk.Label(
            premregsplit_filter_box, text="Click \"Load Filter Options\" above to populate these lists from the file's own values.",
            style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9), justify="left")
        premregsplit_filter_note.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 8))
        self._bind_dynamic_wraplength(premregsplit_filter_note, premregsplit_filter_box, side_padding=10)

        load_filters_row = ttk.Frame(premregsplit_filter_box, style="Panel.TFrame")
        load_filters_row.grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 10))
        self.premregsplit_load_filters_btn = ttk.Button(load_filters_row, text="Load Filter Options",
                                                          command=self._load_premregsplit_filter_options)
        self.premregsplit_load_filters_btn.pack(side="left")
        self.premregsplit_filter_status_var = tk.StringVar(value="")
        ttk.Label(load_filters_row, textvariable=self.premregsplit_filter_status_var, style="Muted.TLabel",
                  background=self.COLOR_PANEL).pack(side="left", padx=(10, 0))

        def _premregsplit_filter_listbox_col(parent, col_idx, title):
            col = ttk.Frame(parent, style="Panel.TFrame")
            col.grid(row=2, column=col_idx, sticky="nsew", padx=6, pady=4)
            ttk.Label(col, text=title, style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
            ttk.Label(col, text="Ctrl+click to select multiple \u00b7 none = all", style="Muted.TLabel",
                      background=self.COLOR_PANEL).pack(anchor="w", pady=(0, 4))
            frame = tk.Frame(col, highlightbackground="#D6DCE1", highlightthickness=1)
            frame.pack(fill="both", expand=True)
            scroll = ttk.Scrollbar(frame, orient="vertical")
            listbox = tk.Listbox(frame, selectmode="extended", height=10, width=32,
                                  exportselection=False, font=("Segoe UI", 10),
                                  activestyle="none", relief="flat", borderwidth=0,
                                  selectbackground=self.COLOR_ACCENT, selectforeground="white",
                                  highlightthickness=0, yscrollcommand=scroll.set)
            scroll.config(command=listbox.yview)
            listbox.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
            return listbox

        self.premregsplit_class_listbox = _premregsplit_filter_listbox_col(premregsplit_filter_box, 0, "Business Class filter")
        self.premregsplit_client_listbox = _premregsplit_filter_listbox_col(premregsplit_filter_box, 1, "Insured Name filter")

        # ---- Section: Load & Build Groups ----
        load_box = ttk.Labelframe(canvas_frame, text="  Load & Build Groups  ", padding=(14, 12))
        load_box.pack(fill="x", padx=12, pady=8)

        ttk.Label(load_box, text="Loads the register, applies filters, and builds the per-company groups on tab 2.",
                  style="Panel.TLabel", foreground=self.COLOR_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))

        premregsplit_load_btn_row = ttk.Frame(load_box)
        premregsplit_load_btn_row.pack(anchor="w", fill="x")
        self.premregsplit_build_btn = self._make_action_button(
            premregsplit_load_btn_row, "Load & Build Groups", self.load_and_build_premregsplit_groups, kind="accent")
        self.premregsplit_build_btn.pack(side="left")

        self.premregsplit_build_progress = perf_utils.ProgressBarPanel(load_box)
        self.premregsplit_build_progress.frame.pack(fill="x", pady=(8, 0))

        self.premregsplit_build_summary_var = tk.StringVar(value="")
        self.premregsplit_build_summary_label = ttk.Label(load_box, textvariable=self.premregsplit_build_summary_var,
                                                            foreground=self.COLOR_SUCCESS, font=("Segoe UI", 10, "bold"))
        self.premregsplit_build_summary_label.pack(anchor="w")
        self.premregsplit_build_summary_label.pack_forget()

        # ---- Section: Log ----
        log_box = ttk.Labelframe(canvas_frame, text="  Activity Log  ", padding=10)
        log_box.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.premregsplit_setup_log = scrolledtext.ScrolledText(log_box, height=12, font=("Consolas", 9),
                                                                  bg="#FBFCFD", relief="flat", borderwidth=0)
        self.premregsplit_setup_log.pack(fill="both", expand=True)

    def _build_premregsplit_groups_tab(self):
        f = self.tab_premregsplit_groups

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Generate Excel Reports", style="Accent.TButton", width=24,
                   command=self.generate_premregsplit_reports).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Open Output Folder", command=self.open_premregsplit_output_folder, width=20).pack(side="left")
        self.premregsplit_results_summary_var = tk.StringVar(value="Run 'Load & Build Groups' on tab 1 to see one row per company here.")
        ttk.Label(top, textvariable=self.premregsplit_results_summary_var, style="Muted.TLabel").pack(side="right")

        hint = ttk.Label(f, text="One workbook per coinsurer (\"OFC\"-named companies pooled together).",
                          style="Muted.TLabel", font=("Segoe UI", 9), wraplength=900, justify="left")
        hint.pack(anchor="w", padx=14, pady=(0, 6))

        merge_row = ttk.Frame(f)
        merge_row.pack(fill="x", padx=12, pady=(0, 8))
        self.premregsplit_merge_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(merge_row, text="Merge files for the same company (combine before generating)",
                         variable=self.premregsplit_merge_var).pack(side="left")
        merge_note = ttk.Label(
            f, text="Recommended: combines branch-name variants of the same company (e.g. \"EFU H.O., Karachi\" "
                    "/ \"EFU Multan Branch\") using the same rule as other modules. No data is lost.",
            style="Muted.TLabel", font=("Segoe UI", 9), wraplength=900, justify="left")
        merge_note.pack(anchor="w", padx=14, pady=(0, 6))

        self.premregsplit_reports_progress = perf_utils.ProgressBarPanel(f)
        self.premregsplit_reports_progress.frame.pack(fill="x", padx=12, pady=(0, 8))

        tree_frame = ttk.Frame(f)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("company", "num_policies", "gross_total", "net_total")
        self.premregsplit_results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.premregsplit_results_tree.heading("company", text="Coinsurer (output file)")
        self.premregsplit_results_tree.heading("num_policies", text="Policies")
        self.premregsplit_results_tree.heading("gross_total", text="Gross Premium")
        self.premregsplit_results_tree.heading("net_total", text="Net Premium")
        self.premregsplit_results_tree.column("company", width=380)
        self.premregsplit_results_tree.column("num_policies", width=80, anchor="center")
        self.premregsplit_results_tree.column("gross_total", width=170, anchor="e")
        self.premregsplit_results_tree.column("net_total", width=170, anchor="e")
        self.premregsplit_results_tree.tag_configure("oddrow", background="#F6F8FB")
        self.premregsplit_results_tree.tag_configure("evenrow", background="white")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.premregsplit_results_tree.yview)
        self.premregsplit_results_tree.configure(yscrollcommand=vsb.set)
        self.premregsplit_results_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_premregsplit_emails_tab(self):
        f = self.tab_premregsplit_emails

        top = ttk.Frame(f)
        top.pack(fill="x", padx=12, pady=(12, 8))
        ttk.Button(top, text="Draft Emails", style="Accent.TButton", command=self.draft_premregsplit_emails, width=18).pack(side="left", padx=(0, 12))
        self.premregsplit_dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Dry run (preview only, do not actually send)", variable=self.premregsplit_dry_run_var).pack(side="left", padx=(0, 12))
        ttk.Button(top, text="Send All", style="Send.TButton", command=self.send_all_premregsplit_emails, width=14).pack(side="left")

        sig_frame = ttk.Labelframe(f, text="  Your Signature (appears after 'Regards,' in every email - shared across every GIMS module)  ", padding=10)
        sig_frame.pack(fill="x", padx=12, pady=(0, 8))
        sig_frame.columnconfigure(0, weight=1)
        self.premregsplit_signature_text = tk.Text(sig_frame, height=3, font=("Segoe UI", 10), relief="solid", borderwidth=1, wrap="word")
        self.premregsplit_signature_text.grid(row=0, column=0, sticky="ew")
        self.premregsplit_signature_text.insert("1.0", eg.load_signature())

        sig_btn_holder = tk.Frame(sig_frame, bg=self.COLOR_PANEL)
        sig_btn_holder.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        sig_btn = self._make_action_button(sig_btn_holder, "Save Signature", self.save_premregsplit_signature,
                                            kind="quiet", compact=True)
        sig_btn.pack(expand=True)

        outer = ttk.Frame(f)
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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

        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1, minsize=560)

        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        ttk.Label(left, text="Coinsurer Files", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        columns = ("company", "policies", "email", "status")
        self.premregsplit_email_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=26)
        self.premregsplit_email_tree.heading("company", text="Coinsurer")
        self.premregsplit_email_tree.heading("policies", text="Policies")
        self.premregsplit_email_tree.heading("email", text="Email")
        self.premregsplit_email_tree.heading("status", text="Status")
        self.premregsplit_email_tree.column("company", width=210)
        self.premregsplit_email_tree.column("policies", width=60, anchor="center")
        self.premregsplit_email_tree.column("email", width=170)
        self.premregsplit_email_tree.column("status", width=150)
        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.premregsplit_email_tree.yview)
        self.premregsplit_email_tree.configure(yscrollcommand=tree_scroll.set)
        self.premregsplit_email_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.premregsplit_email_tree.bind("<<TreeviewSelect>>", self.on_select_premregsplit_email)

        right = ttk.Labelframe(body, text="  Email Preview  ", padding=14)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(8, weight=1)

        ttk.Label(right, text="To (recipient email - separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.premregsplit_email_to_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.premregsplit_email_to_var, font=("Segoe UI", 10))\
            .grid(row=1, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Cc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.premregsplit_email_cc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.premregsplit_email_cc_var, font=("Segoe UI", 10))\
            .grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(right, text="Bcc (separate multiple with a semicolon)", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.premregsplit_email_bcc_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.premregsplit_email_bcc_var, font=("Segoe UI", 10))\
            .grid(row=5, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(right, text="Subject", style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=6, column=0, sticky="w", pady=(0, 2))
        self.premregsplit_subject_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.premregsplit_subject_var, font=("Segoe UI", 10))\
            .grid(row=7, column=0, sticky="ew")

        body_wrap = ttk.Frame(right)
        body_wrap.grid(row=8, column=0, sticky="nsew", pady=(14, 10))
        body_wrap.columnconfigure(0, weight=1)
        body_wrap.rowconfigure(1, weight=1)
        ttk.Label(body_wrap, text="Body", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.premregsplit_body_text = scrolledtext.ScrolledText(body_wrap, font=("Segoe UI", 10), relief="solid", borderwidth=1, height=14)
        self.premregsplit_body_text.grid(row=1, column=0, sticky="nsew", pady=(2, 0))

        btn_row = ttk.Frame(right)
        btn_row.grid(row=9, column=0, sticky="w", pady=(10, 0))

        save_btn = self._make_action_button(btn_row, "Save Edits to Selected Email", self.save_premregsplit_email_edit, kind="accent")
        save_btn.pack(side="left", padx=(0, 8))

        send_one_btn = self._make_action_button(btn_row, "Send This Email", self.send_individual_premregsplit_email, kind="success")
        send_one_btn.pack(side="left")

    # ---------- Premium Register (company split) actions ----------
    def browse_premregsplit_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.premregsplit_file_var.set(path)

    def browse_premregsplit_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.premregsplit_output_dir_var.set(path)

    def browse_premregsplit_email_lookup_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if path:
            self.premregsplit_email_lookup_file_var.set(path)

    def log_premregsplit(self, msg):
        self.premregsplit_setup_log.insert("end", str(msg) + "\n")
        self.premregsplit_setup_log.see("end")

    def _parse_premregsplit_filter_date(self, var, label):
        text = var.get().strip()
        if not text:
            return None
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError(f"{label} must be in YYYY-MM-DD format (or left blank), got '{text}'.")

    def _load_premregsplit_filter_options(self):
        path = self.premregsplit_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid Premium Register file first.")
            return

        self.premregsplit_filter_status_var.set("Loading...")

        def worker(report):
            df, colmap = prcp.load_premium_register(path, report=report)
            return df, colmap

        def on_done(result):
            df, colmap = result
            self.premregsplit_df = df
            self.premregsplit_colmap = colmap
            self.premregsplit_loaded_path = path

            def _populate(listbox, canonical_col):
                listbox.delete(0, "end")
                orig_col = colmap.get(canonical_col)
                if not orig_col or orig_col not in df.columns:
                    return
                values = sorted(df[orig_col].dropna().astype(str).str.strip().unique())
                values = [v for v in values if v]
                for i, v in enumerate(values):
                    listbox.insert("end", v)
                    listbox.itemconfig(i, bg="#F4F7FA" if i % 2 else "white")

            _populate(self.premregsplit_class_listbox, "Business Class")
            _populate(self.premregsplit_client_listbox, "Insured")

            self.premregsplit_filter_status_var.set(f"Loaded {len(df):,} row(s).")
            self.status_var.set("Premium Register: filter options loaded.")

        def on_error(exc):
            self.premregsplit_filter_status_var.set("Load failed - see error.")
            messagebox.showerror("Could not load file", str(exc))

        perf_utils.run_in_background(
            self, worker, on_done=on_done, on_error=on_error,
            busy_widgets=[self.premregsplit_load_filters_btn],
        )

    def load_and_build_premregsplit_groups(self):
        path = self.premregsplit_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid Premium Register file.")
            return

        try:
            date_filter_from = self._parse_premregsplit_filter_date(self.premregsplit_date_from_var, "Date Filter 'From'")
            date_filter_to = self._parse_premregsplit_filter_date(self.premregsplit_date_to_var, "Date Filter 'To'")
        except ValueError as e:
            messagebox.showerror("Invalid date filter", str(e))
            return
        if date_filter_from and date_filter_to and date_filter_from > date_filter_to:
            messagebox.showerror("Invalid date filter", "Date Filter 'From' must not be after 'To'.")
            return
        date_filter_field = self.premregsplit_date_field_var.get()

        selected_classes = [self.premregsplit_class_listbox.get(i) for i in self.premregsplit_class_listbox.curselection()]
        selected_clients = [self.premregsplit_client_listbox.get(i) for i in self.premregsplit_client_listbox.curselection()]

        self.premregsplit_setup_log.delete("1.0", "end")
        self.premregsplit_build_summary_var.set("")
        self.premregsplit_build_summary_label.pack_forget()
        self.status_var.set("Loading and building company groups...")
        self.premregsplit_build_progress.start(total=None, text="Loading file...")

        def worker(report):
            if self.premregsplit_df is not None and self.premregsplit_loaded_path == path:
                df = self.premregsplit_df.copy()
                colmap = self.premregsplit_colmap
                report(f"Using already-loaded file: {os.path.basename(path)} ({len(df):,} row(s)).")
            else:
                report(f"Loading Premium Register: {os.path.basename(path)}...")
                df, colmap = prcp.load_premium_register(path, report=report)
                report(f"  {len(df):,} row(s) loaded.")

            num_rows_loaded = len(df)

            coins = prcp.filter_to_coinsurer_rows(df, colmap)
            report(f"\n{len(coins):,} of {num_rows_loaded:,} row(s) are coinsurer (Leader = \"No\") rows "
                   f"- IGI's own rows are excluded.")

            if date_filter_from or date_filter_to:
                before_count = len(coins)
                coins, num_no_date = prcp.filter_by_date(
                    coins, colmap, date_filter_field, start_date=date_filter_from, end_date=date_filter_to)
                range_text = f"{date_filter_from or '(any)'} to {date_filter_to or '(any)'}"
                report(f"Date Filter applied on '{date_filter_field}', range {range_text}: "
                       f"{before_count:,} -> {len(coins):,} row(s) kept"
                       + (f", {num_no_date:,} row(s) excluded for a missing/unparseable date." if num_no_date else "."))

            field_filters = []
            if selected_classes:
                class_col = colmap.get("Business Class")
                if class_col:
                    field_filters.append((class_col, selected_classes))
            if selected_clients:
                insured_col = colmap.get("Insured")
                if insured_col:
                    field_filters.append((insured_col, selected_clients))
            if field_filters:
                before_count = len(coins)
                coins = prcp.filter_by_fields(coins, field_filters)
                report(f"Business Class/Insured filter(s) applied: {before_count:,} -> {len(coins):,} row(s) kept.")

            groups, totals, breakdowns = prcp.build_company_buckets(coins, colmap)
            return df, colmap, coins, groups, totals, breakdowns

        def on_done(result):
            df, colmap, coins, groups, totals, breakdowns = result
            self.premregsplit_df = df
            self.premregsplit_colmap = colmap
            self.premregsplit_loaded_path = path
            self.premregsplit_groups = groups
            self.premregsplit_totals = totals
            self.premregsplit_breakdowns = breakdowns

            self._refresh_premregsplit_results_tree()

            summary_text = f"{len(coins):,} coinsurer row(s) \u2192 {len(groups)} company file(s)."
            self.premregsplit_build_summary_var.set(summary_text)
            self.premregsplit_build_summary_label.pack(anchor="w", pady=(8, 0))
            self.premregsplit_build_progress.finish(summary_text)
            self.premregsplit_build_progress.hide()
            self.status_var.set(f"{len(groups)} company files identified. See tab 2.")

        def on_error(exc):
            self.premregsplit_build_progress.finish("Failed - see the Activity Log for details.")
            self.log_premregsplit(f"\nERROR: {exc}")
            messagebox.showerror("Load failed", str(exc))
            self.status_var.set("Premium Register load failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress=lambda msg: self.log_premregsplit(msg),
            on_progress_value=self.premregsplit_build_progress.update,
            on_done=on_done, on_error=on_error,
            busy_widgets=[self.premregsplit_build_btn],
        )

    def _refresh_premregsplit_results_tree(self):
        self.premregsplit_results_tree.delete(*self.premregsplit_results_tree.get_children())
        for i, (bucket_label, _g) in enumerate(self.premregsplit_groups.items()):
            t = self.premregsplit_totals[bucket_label]
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.premregsplit_results_tree.insert("", "end", values=(
                bucket_label, t["num_policies"], f"{t['total_gross_premium']:,.0f}", f"{t['total_net_premium']:,.0f}"
            ), tags=(tag,))
        self.premregsplit_results_summary_var.set(f"{len(self.premregsplit_groups)} company file(s) identified")

    def _run_premregsplit_self_merge(self):
        import company_alias_resolver as car

        if not self.premregsplit_groups:
            return True

        weights = {label: self.premregsplit_totals[label]["num_policies"] for label in self.premregsplit_groups}
        merge_map = car.self_merge_bucket_labels(list(self.premregsplit_groups.keys()), weight_by=weights)

        if all(v == k for k, v in merge_map.items()):
            return True

        ins_col = self.premregsplit_colmap["Ins"]

        def combine_group(dfs, new_label):
            import pandas as pd
            combined = pd.concat(dfs, ignore_index=True)
            if ins_col in combined.columns:
                combined[ins_col] = new_label
            return combined

        merged_groups, merged_totals, merged_breakdowns = car.merge_buckets(
            self.premregsplit_groups, self.premregsplit_totals, self.premregsplit_breakdowns, merge_map, combine_group)

        n_before, n_after = len(self.premregsplit_groups), len(merged_groups)
        self.premregsplit_groups, self.premregsplit_totals, self.premregsplit_breakdowns = merged_groups, merged_totals, merged_breakdowns
        self._refresh_premregsplit_results_tree()
        self.log_premregsplit(f"\nMerge checkbox: combined {n_before} company file(s) into {n_after} by shared company name.")
        return True

    def generate_premregsplit_reports(self):
        if not self.premregsplit_groups:
            messagebox.showwarning("Nothing to generate", "Run 'Load & Build Groups' first (tab 1).")
            return

        if self.premregsplit_merge_var.get():
            if not self._run_premregsplit_self_merge():
                return

        base_output_dir = self.premregsplit_output_dir_var.get().strip()
        if not base_output_dir:
            messagebox.showerror("No output folder", "Set an output folder on tab 1 first.")
            return
        output_dir = self._new_run_folder(base_output_dir)
        self.premregsplit_last_run_dir = output_dir

        groups = self.premregsplit_groups
        breakdowns = self.premregsplit_breakdowns
        # Read on the main thread, same rule as every other Tkinter widget
        # value this worker closure captures below.
        business_type = self.premregsplit_business_type_var.get()
        self.status_var.set(f"Generating {len(groups)} Premium Register workbook(s)...")
        self.premregsplit_reports_progress.start(total=len(groups), text=f"0 / {len(groups)} workbook(s) generated")

        def worker(report):
            return prcg2.build_company_workbooks_parallel(groups, breakdowns, output_dir, report=report,
                                                            business_type=business_type)

        def on_done(result):
            generated, filename_lookup, errors = result
            self.premregsplit_filename_lookup = filename_lookup
            self.premregsplit_reports_progress.finish(f"Done: {len(generated)} / {len(groups)} workbook(s) generated.")
            msg = f"Generated {len(generated)} workbook(s) in:\n{output_dir}"
            if errors:
                msg += f"\n\n{len(errors)} file(s) failed:\n" + "\n".join(errors[:10])
            messagebox.showinfo("Done", msg)
            self.status_var.set(f"Generated {len(generated)} Premium Register workbook(s).")

        def on_error(exc):
            self.premregsplit_reports_progress.finish("Failed - see the Activity Log for details.")
            messagebox.showerror("Report generation failed", str(exc))
            self.status_var.set("Report generation failed.")

        perf_utils.run_in_background(
            self, worker,
            on_progress_value=self.premregsplit_reports_progress.update,
            on_progress=lambda msg: self.log_premregsplit(msg),
            on_done=on_done, on_error=on_error,
        )

    def open_premregsplit_output_folder(self):
        output_dir = self.premregsplit_last_run_dir or self.premregsplit_output_dir_var.get().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir) if hasattr(os, "startfile") else os.system(f'open "{output_dir}"')
        else:
            messagebox.showinfo("Not found", "No output folder yet - generate reports first.")

    def draft_premregsplit_emails(self):
        if not self.premregsplit_groups:
            messagebox.showwarning("Nothing to draft", "Run 'Load & Build Groups' first (tab 1).")
            return

        self.premregsplit_email_lookup, self.premregsplit_poc_lookup = {}, {}
        if self.premregsplit_load_email_var.get():
            lookup_path = self.premregsplit_email_lookup_file_var.get().strip()
            if lookup_path and os.path.exists(lookup_path):
                try:
                    self.premregsplit_email_lookup, self.premregsplit_poc_lookup = ra.load_email_lookup_from_file(lookup_path)
                    self.log_premregsplit(f"\nLoaded {len(self.premregsplit_email_lookup)} company email mapping(s) "
                                           f"from '{os.path.basename(lookup_path)}'.")
                except Exception as e:
                    self.log_premregsplit(f"\nWARNING: could not load email lookup: {e}")
            else:
                self.log_premregsplit("\nNo email lookup file selected - all drafts will show 'Missing email'.")

        signature = self.premregsplit_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)

        period_text = self._format_period_text(self.premregsplit_date_from_var, self.premregsplit_date_to_var)
        self.premregsplit_drafts = prce2.build_premium_register_email_drafts(
            self.premregsplit_groups, self.premregsplit_totals, self.premregsplit_email_lookup,
            poc_lookup=self.premregsplit_poc_lookup, signature=signature, period_text=period_text,
            business_type=self.premregsplit_business_type_var.get(),
        )
        self.premregsplit_email_tree.delete(*self.premregsplit_email_tree.get_children())
        for d in self.premregsplit_drafts:
            self.premregsplit_email_tree.insert("", "end", values=(
                d["company_name"], d["num_policies"], d["email"] or "(missing)", d["status"]
            ))
        self.status_var.set(f"Drafted {len(self.premregsplit_drafts)} Premium Register emails.")

    def save_premregsplit_signature(self):
        signature = self.premregsplit_signature_text.get("1.0", "end").strip()
        eg.save_signature(signature)
        self.status_var.set("Signature saved (shared across every GIMS module). Click Draft Emails again to refresh existing drafts.")

    def on_select_premregsplit_email(self, event):
        sel = self.premregsplit_email_tree.selection()
        if not sel:
            return
        idx = self.premregsplit_email_tree.index(sel[0])
        d = self.premregsplit_drafts[idx]
        self.premregsplit_email_to_var.set(d.get("email") or "")
        self.premregsplit_email_cc_var.set(d.get("cc") or "")
        self.premregsplit_email_bcc_var.set(d.get("bcc") or "")
        self.premregsplit_subject_var.set(d["subject"])
        self.premregsplit_body_text.delete("1.0", "end")
        self.premregsplit_body_text.insert("1.0", d["body"])

    def save_premregsplit_email_edit(self):
        sel = self.premregsplit_email_tree.selection()
        if not sel:
            messagebox.showinfo("No coinsurer selected", "Select a coinsurer from the list on the left first.")
            return
        idx = self.premregsplit_email_tree.index(sel[0])
        new_email = self.premregsplit_email_to_var.get().strip()
        self.premregsplit_drafts[idx]["email"] = new_email or None
        self.premregsplit_drafts[idx]["cc"] = self.premregsplit_email_cc_var.get().strip()
        self.premregsplit_drafts[idx]["bcc"] = self.premregsplit_email_bcc_var.get().strip()
        self.premregsplit_drafts[idx]["subject"] = self.premregsplit_subject_var.get()
        self.premregsplit_drafts[idx]["body"] = self.premregsplit_body_text.get("1.0", "end").strip()

        status = "Ready" if new_email else "Missing email - needs manual entry"
        row_id = self.premregsplit_email_tree.get_children()[idx]
        self.premregsplit_email_tree.set(row_id, "email", new_email or "(missing)")
        self.premregsplit_email_tree.set(row_id, "status", status)

        self.status_var.set(f"Updated draft for {self.premregsplit_drafts[idx]['company_name']}.")

    def send_all_premregsplit_emails(self):
        if not self.premregsplit_drafts:
            messagebox.showwarning("Nothing to send", "Draft emails first.")
            return

        filename_lookup = self.premregsplit_filename_lookup
        missing = [d["company_name"] for d in self.premregsplit_drafts
                   if d["company_name"] not in filename_lookup]
        if missing:
            proceed = messagebox.askyesno(
                "Reports not generated",
                f"{len(missing)} of {len(self.premregsplit_drafts)} file(s) have no generated Excel workbook yet, "
                f"so those emails would go out with no attachment:\n\n"
                + "\n".join(missing[:10]) + (f"\n...and {len(missing) - 10} more" if len(missing) > 10 else "")
                + "\n\nGo to tab 2 and click 'Generate Excel Reports' first, "
                  "or continue anyway without attachments?\n\nContinue anyway?"
            )
            if not proceed:
                return

        dry_run = self.premregsplit_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send {len(self.premregsplit_drafts)} emails via Outlook now? This cannot be undone."):
                return

        output_dir = self.premregsplit_last_run_dir or self.premregsplit_output_dir_var.get().strip()

        def worker():
            results = prce2.send_premium_register_emails(
                self.premregsplit_drafts, output_dir, filename_lookup, dry_run=dry_run
            )
            sent = sum(1 for r in results if r["sent"])
            failed = len(results) - sent
            missing_email = sum(1 for d in self.premregsplit_drafts if not d.get("email"))
            num_rows_loaded = len(self.premregsplit_df) if self.premregsplit_df is not None else 0
            num_coinsurer_rows = sum(t["num_policies"] for t in self.premregsplit_totals.values()) if self.premregsplit_totals else 0
            prclog2.log_run(
                num_rows_loaded=num_rows_loaded,
                num_coinsurer_rows=num_coinsurer_rows,
                num_groups=len(self.premregsplit_groups),
                num_files_generated=len(filename_lookup),
                num_sent=sent, num_failed=failed, num_missing_email=missing_email, dry_run=dry_run,
            )
            for i, r in enumerate(results):
                self.premregsplit_email_tree.set(self.premregsplit_email_tree.get_children()[i], "status",
                                                  "Sent" if r["sent"] else f"Failed: {r['send_message']}")
            messagebox.showinfo("Done", f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")
            self.status_var.set(f"{'[DRY RUN] ' if dry_run else ''}Sent: {sent}, Failed: {failed}")

        threading.Thread(target=worker, daemon=True).start()

    def send_individual_premregsplit_email(self):
        sel = self.premregsplit_email_tree.selection()
        if not sel:
            messagebox.showinfo("No coinsurer selected", "Click a coinsurer in the list on the left first.")
            return
        idx = self.premregsplit_email_tree.index(sel[0])

        self.save_premregsplit_email_edit()
        d = self.premregsplit_drafts[idx]

        if not d.get("email"):
            messagebox.showerror("Missing email", "Enter a recipient email address in the 'To' field before sending.")
            return

        output_dir = self.premregsplit_last_run_dir or self.premregsplit_output_dir_var.get().strip()
        fname = self.premregsplit_filename_lookup.get(d["company_name"])
        attachment_path = os.path.join(output_dir, fname) if fname and os.path.exists(os.path.join(output_dir, fname)) else None
        if attachment_path is None:
            if not messagebox.askyesno(
                "No report generated",
                f"No Excel workbook has been generated yet for {d['company_name']}. "
                f"Send this email without an attachment?\n\n"
                f"(Go to tab 2 and click 'Generate Excel Reports' first if you'd rather attach one.)"
            ):
                return

        dry_run = self.premregsplit_dry_run_var.get()
        if not dry_run:
            if not messagebox.askyesno("Confirm", f"Send this email to {d['email']} ({d['company_name']}) now?"):
                return

        row_id = self.premregsplit_email_tree.get_children()[idx]

        def worker():
            success, message = osend.send_email(d["email"], d["subject"], d["body"], attachment_path, dry_run,
                                                 cc=d.get("cc"), bcc=d.get("bcc"))
            self.premregsplit_email_tree.set(row_id, "status", "Sent" if success else f"Failed: {message}")
            prefix = "[DRY RUN] " if dry_run else ""
            messagebox.showinfo("Done", f"{prefix}{message}")
            self.status_var.set(f"{prefix}{message}")

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    # REQUIRED for a PyInstaller-frozen build that uses
    # concurrent.futures.ProcessPoolExecutor (receivable_generator.py,
    # claims_register_generator.py, report_generator.py all use it for
    # parallel workbook generation). On Windows, spawning a worker
    # process re-imports and re-runs this whole script from scratch. In
    # a normal "python app.py" run, multiprocessing's own bookkeeping
    # correctly recognizes a spawned worker and skips straight to running
    # its target function, never reaching the BalanceApp()/mainloop()
    # lines below. In a FROZEN exe, that recognition depends entirely on
    # freeze_support() being called here, as literally the first thing
    # inside this guard, before anything else. Without it, every worker
    # process falls straight through into "app = BalanceApp();
    # app.mainloop()" - i.e. it opens a brand new, fully separate copy of
    # the whole GUI. That's the exact bug: extra windows appearing right
    # before/during "Generate Excel Reports" (when the batch is large
    # enough to trigger the parallel path), and closing the original
    # window kills every one of them at once, since PyInstaller ties
    # spawned children to the parent through a Windows Job Object.
    import multiprocessing
    multiprocessing.freeze_support()

    app = BalanceApp()
    app.mainloop()
