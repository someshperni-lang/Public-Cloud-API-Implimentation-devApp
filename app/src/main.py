"""
SAP bulk uploads - DEV (Windows + Android app)
==============================================

The installable twin of sap_dev_dashboard.py. Same login, banner, tabs,
templates, preview, posting, results, metadata explorer and history.

It does not re-implement any payload logic: it imports the dashboard's own
scripts from ./scripts (copied unchanged from Sap_Upload_Dashboard_DEV) and
calls their build_documents() and sap_common.post_documents(), so a preview
or post here gives the same result as the dashboard.

Results of every run are saved as CSV in the app's data folder on the device
(the dashboard's results/ folder), and the History tab reads them.
"""
from __future__ import annotations

import asyncio
import hashlib
import importlib
import json
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

import flet as ft
import pandas as pd

APP_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = APP_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import sap_common as sc  # noqa: E402

ENVIRONMENT = "DEV"
# Same tabs, order and dropdown as sap_dev_dashboard.py
TABS = [
    {"label": "Material PO texts", "icon": "📝",
     "scripts": {"ptext": "bulk_product_purchase_text_upload"}},
    {"label": "Purchase requisitions", "icon": "📋",
     "scripts": {"pr": "bulk_purchase_requisition_upload"}},
    {"label": "Purchase orders", "icon": "🛒", "picker": "Purchase order type",
     "scripts": {"po": "bulk_purchase_order_upload",
                 "subcon": "bulk_subcontracting_po_upload"}},
    {"label": "Goods movements", "icon": "📦", "scripts": {"gr": "bulk_material_document_upload"}},
    {"label": "Supplier invoices", "icon": "🧾",
     "scripts": {"invoice": "bulk_supplier_invoice_upload"}},
]
EXTRA_TABS = [("metadata", "🔎", "Metadata"), ("history", "📊", "History")]
RESULT_COLUMNS = sc.RESULT_COLUMNS

# ----------------------------------------------------------------------------
# Look (values from the dashboard's CSS and .streamlit/config.toml)
# ----------------------------------------------------------------------------
PRIMARY = "#2B5F8A"
BG = "#FBFCFD"
BG2 = "#EEF2F5"
TEXT = "#1C2A36"
MUTED = "#6B7785"
LINE = "#E1E6EB"
SIDEBAR = ["#1E3A8A", "#5B21B6", "#9D174D"]
HERO = ["#2563EB", "#7C3AED", "#DB2777", "#EA580C"]
BADGE = "#FBBF24"
CARD = {"green": ("#059669", "#10B981"), "red": ("#DC2626", "#F43F5E"),
        "blue": ("#2563EB", "#06B6D4"), "purple": ("#7C3AED", "#C026D3"),
        "amber": ("#D97706", "#F59E0B")}
NOTE = {"info": ("#E8F1FB", "#1E5A9E"), "warning": ("#FEF6E0", "#8A5A00"),
        "error": ("#FDECEC", "#9B1C1C"), "success": ("#E6F6EC", "#1E6B43")}
WIDE = 900
WHITE_SOFT = ft.Colors.with_opacity(0.14, ft.Colors.WHITE)
WHITE_LINE = ft.Colors.with_opacity(0.35, ft.Colors.WHITE)


def status_style(val):
    """(text colour, background) for a Status cell - same as the dashboard."""
    val = str(val)
    if val in ("Success", "Ready"):
        return "#065F46", "#D1FAE5"
    if val.startswith("Failed") or val == "Error":
        return "#9F1239", "#FFE4E6"
    if val == "Invalid":
        return "#92400E", "#FEF3C7"
    return None


def status_counts(df):
    s = df["Status"].astype(str)
    return {"success": int((s == "Success").sum()),
            "failed": int(s.str.startswith("Failed").sum() + (s == "Error").sum()),
            "invalid": int((s == "Invalid").sum()),
            "total": len(df)}


GENERIC_SAP_MESSAGES = {"Exception raised without specific error",
                        "Purchase order still contains faulty items",
                        "PO header data still faulty"}


def main_failure_message(detail):
    parts = [p.strip() for p in str(detail).split(" | ") if p.strip()]
    for p in parts:
        if p not in GENERIC_SAP_MESSAGES:
            return p
    return parts[0] if parts else "(no message)"


def data_dir() -> Path:
    d = Path(os.getenv("FLET_APP_STORAGE_DATA") or Path.home() / ".sap_dev_uploads")
    d.mkdir(parents=True, exist_ok=True)
    return d


RESULTS_DIR = data_dir() / "results"


def save_results(slug, rows) -> Path:
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    path = RESULTS_DIR / f"{slug}_{datetime.now():%Y%m%d_%H%M%S}.csv"
    pd.DataFrame(rows, columns=RESULT_COLUMNS).to_csv(path, index=False)
    return path


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Same clean-up the dashboard applies after its data editor."""
    df = df.fillna("").astype(str)
    df = df.map(sc.clean) if hasattr(df, "map") else df.applymap(sc.clean)
    return df[(df != "").any(axis=1)].reset_index(drop=True)


# ----------------------------------------------------------------------------
# Small UI building blocks
# ----------------------------------------------------------------------------
def txt(value, size=15, color=TEXT, weight=None, selectable=False, **kw):
    return ft.Text(str(value), size=size, color=color, weight=weight, selectable=selectable, **kw)


def caption(value):
    return txt(value, 14, MUTED)


def h2(value):
    return txt(value, 30, TEXT, ft.FontWeight.W_700)


def h3(value):
    return txt(value, 22, TEXT, ft.FontWeight.W_700)


def note(value, kind="info"):
    bg, fg = NOTE[kind]
    return ft.Row([ft.Container(txt(value, 15, fg, selectable=True), bgcolor=bg,
                                border_radius=10, expand=True,
                                padding=ft.Padding.symmetric(horizontal=18, vertical=14))])


def code_chip(value):
    return ft.Container(txt(value, 13, "#1F7A3D", font_family="monospace"), bgcolor=BG2,
                        border_radius=6, padding=ft.Padding.symmetric(horizontal=8, vertical=3))


def stat_cards(items):
    """Gradient cards like the dashboard's .card (value + label)."""
    cards = []
    for label, value, colour in items:
        a, b = CARD[colour]
        cards.append(ft.Container(
            ft.Column([txt(value, 30, "white", ft.FontWeight.W_800),
                       txt(label, 14, "white")], spacing=2, tight=True),
            col={"xs": 6, "md": 12 / max(1, len(items))},
            gradient=ft.LinearGradient(begin=ft.Alignment.TOP_LEFT, end=ft.Alignment.BOTTOM_RIGHT,
                                       colors=[a, b]),
            border_radius=14, padding=ft.Padding.symmetric(horizontal=18, vertical=14),
            shadow=ft.BoxShadow(blur_radius=18, spread_radius=-8, offset=ft.Offset(0, 8),
                                color=ft.Colors.with_opacity(0.35, "black"))))
    return ft.ResponsiveRow(cards, spacing=12, run_spacing=12)


def outline_button(label, on_click, icon=None, expand=False):
    return ft.OutlinedButton(label, icon=icon, on_click=on_click, expand=expand,
                             style=ft.ButtonStyle(color=TEXT, shape=ft.RoundedRectangleBorder(
                                 radius=8), padding=ft.Padding.symmetric(horizontal=18, vertical=16),
                                 side=ft.BorderSide(1, "#D5DBE1")))


def primary_button(label, on_click, disabled=False, icon=None):
    return ft.FilledButton(label, icon=icon, on_click=on_click, disabled=disabled,
                           style=ft.ButtonStyle(
                               bgcolor={ft.ControlState.DEFAULT: PRIMARY,
                                        ft.ControlState.DISABLED: "#C9D3DC"},
                               color={ft.ControlState.DEFAULT: "white"},
                               shape=ft.RoundedRectangleBorder(radius=8),
                               padding=ft.Padding.symmetric(horizontal=18, vertical=16)))


def table(df: pd.DataFrame, status_col="Status", max_rows=500, on_cell_tap=None):
    """Read-only table (horizontal scroll). Status cells coloured like the dashboard.
    on_cell_tap(row_index, column) makes cells tappable (raw-file editing)."""
    cols = list(df.columns)
    rows = []
    for i, rec in enumerate(df.head(max_rows).itertuples(index=False)):
        cells = []
        for c, v in zip(cols, rec):
            v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
            st = status_style(v) if c == status_col else None
            if st:
                content = ft.Container(txt(v, 14, st[0], ft.FontWeight.W_700), bgcolor=st[1],
                                       padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                                       border_radius=4)
            else:
                content = txt(v if len(v) <= 160 else v[:157] + "...", 14, TEXT, selectable=True,
                              max_lines=3, overflow=ft.TextOverflow.ELLIPSIS)
            cell = ft.DataCell(content)
            if on_cell_tap:
                cell.on_tap = lambda e, i=i, c=c: on_cell_tap(i, c)
                cell.show_edit_icon = False
            cells.append(cell)
        rows.append(ft.DataRow(cells=cells))
    dt = ft.DataTable(
        columns=[ft.DataColumn(txt(c, 14, MUTED, ft.FontWeight.W_600)) for c in cols],
        rows=rows, border=ft.Border.all(1, LINE), border_radius=10,
        horizontal_lines=ft.BorderSide(1, LINE), heading_row_color=BG,
        heading_row_height=44, data_row_min_height=40, data_row_max_height=72,
        column_spacing=22)
    out = [ft.Row([dt], scroll=ft.ScrollMode.AUTO)]
    if len(df) > max_rows:
        out.append(caption(f"Showing the first {max_rows} of {len(df)} rows."))
    return ft.Column(out, spacing=4)


def expander(title, controls, expanded=False):
    return ft.Container(
        ft.ExpansionTile(title=txt(title, 15), controls=controls, expanded=expanded,
                         controls_padding=ft.Padding.only(left=12, right=12, bottom=12),
                         shape=ft.RoundedRectangleBorder(radius=10),
                         collapsed_shape=ft.RoundedRectangleBorder(radius=10)),
        border=ft.Border.all(1, "#D5DBE1"), border_radius=10)


# ----------------------------------------------------------------------------
# App
# ----------------------------------------------------------------------------
class App:
    def __init__(self, page: ft.Page):
        self.page = page
        self.prefs = ft.SharedPreferences()
        self.picker = ft.FilePicker()
        page.services.extend([self.prefs, self.picker])

        self.user = ""
        self.password = ""
        self.remember = False
        self.conn_test = None                 # (ok, msg)
        self.tab = 0                          # index into TABS + EXTRA_TABS
        self.slug_for_tab = {i: next(iter(t["scripts"])) for i, t in enumerate(TABS)}
        self.modules = {slug: importlib.import_module(name)
                        for t in TABS for slug, name in t["scripts"].items()}
        self.state: dict[str, dict] = {slug: {} for slug in self.modules}
        self.metas = None
        self.meta_service = None
        self.meta_set = None
        self.meta_filter = ""
        self.meta_creatable = True
        self.hist_module = "All"
        self.hist_outcome = "All"
        self.hist_text = ""
        self._wide = None

        page.title = f"SAP bulk uploads - {ENVIRONMENT}"
        page.bgcolor = BG
        page.padding = 0
        page.theme = ft.Theme(color_scheme_seed=PRIMARY)
        page.theme_mode = ft.ThemeMode.LIGHT
        page.on_resize = self.on_resize

    # ---- start / layout -------------------------------------------------
    async def start(self):
        try:
            self.remember = bool(await self.prefs.get("sapdev.remember"))
            self.user = (await self.prefs.get("sapdev.user")) or ""
            if self.remember:
                self.password = (await self.prefs.get("sapdev.password")) or ""
        except Exception:
            pass
        self.layout(force=True)

    def on_resize(self, e):
        wide = (self.page.width or 1200) >= WIDE
        if wide != self._wide:
            self.layout()

    def layout(self, force=False):
        wide = (self.page.width or 1200) >= WIDE
        if not force and wide == self._wide:
            return
        self._wide = wide
        self.body = ft.Container(expand=True)
        self.page.controls.clear()
        if wide:
            self.page.drawer = None
            self.sidebar_box = ft.Container(
                width=300, padding=ft.Padding.symmetric(horizontal=22, vertical=28),
                gradient=ft.LinearGradient(begin=ft.Alignment.TOP_CENTER,
                                           end=ft.Alignment.BOTTOM_CENTER,
                                           colors=SIDEBAR, stops=[0, 0.55, 1]))
            root = ft.Row([self.sidebar_box, self.body], spacing=0, expand=True,
                          vertical_alignment=ft.CrossAxisAlignment.STRETCH)
        else:
            self.sidebar_box = ft.Container(
                padding=ft.Padding.symmetric(horizontal=20, vertical=24),
                height=max(700, (self.page.height or 800)),
                gradient=ft.LinearGradient(begin=ft.Alignment.TOP_CENTER,
                                           end=ft.Alignment.BOTTOM_CENTER,
                                           colors=SIDEBAR, stops=[0, 0.55, 1]))
            self.page.drawer = ft.NavigationDrawer(controls=[self.sidebar_box],
                                                   bgcolor=SIDEBAR[2],
                                                   on_dismiss=lambda e: self.render())
            root = self.body
        self.page.add(ft.SafeArea(root, expand=True))
        self.render_sidebar()
        self.render()

    # ---- sidebar (connection) --------------------------------------------
    def render_sidebar(self):
        self.f_user = ft.TextField(value=self.user, bgcolor="white", color=TEXT,
                                   border_radius=8, border_color="transparent",
                                   on_change=self.on_user, dense=True)
        self.f_pw = ft.TextField(value=self.password, password=True, can_reveal_password=True,
                                 bgcolor="white", color=TEXT, border_radius=8,
                                 border_color="transparent", on_change=self.on_pw, dense=True)
        self.f_remember = ft.Checkbox(
            label="Remember on this device", value=self.remember, on_change=self.on_remember,
            check_color=SIDEBAR[1], fill_color="white",
            label_style=ft.TextStyle(color="white", size=14))
        status = []
        if self.conn_test:
            ok, msg = self.conn_test
            status = [ft.Container(txt(msg, 17, "white" if ok else "#FFE4E6"),
                                   bgcolor=ft.Colors.with_opacity(0.18, "#1E0B2E") if ok
                                   else ft.Colors.with_opacity(0.35, "#7F1D1D"),
                                   border_radius=10,
                                   padding=ft.Padding.symmetric(horizontal=18, vertical=20))]
        logo = ft.Container(ft.Image(src="logo.png", fit=ft.BoxFit.CONTAIN), bgcolor="#EFEFEE",
                            border_radius=12, padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                            shadow=ft.BoxShadow(blur_radius=18, spread_radius=-8,
                                                offset=ft.Offset(0, 6),
                                                color=ft.Colors.with_opacity(0.45, "black")))
        keep = ("Saved on this device. Turn off on shared phones." if self.remember
                else "Held in memory for this session only.")
        self.sidebar_box.content = ft.Column([
            logo,
            ft.Container(height=8),
            txt("Connection", 24, "white", ft.FontWeight.W_700),
            txt("Communication user", 15, "white"), self.f_user,
            txt("Password", 15, "white"), self.f_pw,
            self.f_remember,
            txt(keep, 14, ft.Colors.with_opacity(0.75, "white")),
            ft.OutlinedButton("Test connection", on_click=self.on_test, width=1000,
                              style=ft.ButtonStyle(
                                  color="white", bgcolor=WHITE_SOFT,
                                  side=ft.BorderSide(1, WHITE_LINE),
                                  shape=ft.RoundedRectangleBorder(radius=8),
                                  padding=ft.Padding.symmetric(vertical=16))),
            *status,
            ft.Divider(color=ft.Colors.with_opacity(0.25, "white")),
            txt(f"System: {ENVIRONMENT}", 16, "white", ft.FontWeight.W_700),
            txt(sc.DEV_HOST, 13, ft.Colors.with_opacity(0.85, "white"), selectable=True),
            ft.Divider(color=ft.Colors.with_opacity(0.25, "white")),
            txt("Flow: material PO texts, requisition, order, receive, invoice. Each tab builds payloads "
                "with the script's own logic and checks them against the DEV metadata, so a "
                "preview here matches --dry-run exactly.", 13,
                ft.Colors.with_opacity(0.75, "white")),
        ], spacing=10, scroll=ft.ScrollMode.AUTO)

    def on_user(self, e):
        self.user = e.control.value or ""
        self.page.run_task(self._save_login)

    def on_pw(self, e):
        self.password = e.control.value or ""
        if self.remember:
            self.page.run_task(self._save_login)

    async def on_remember(self, e):
        self.remember = bool(e.control.value)
        await self._save_login()
        self.render_sidebar()
        self.page.update()

    async def _save_login(self):
        try:
            await self.prefs.set("sapdev.user", self.user)
            await self.prefs.set("sapdev.remember", self.remember)
            if self.remember:
                await self.prefs.set("sapdev.password", self.password)
            else:
                await self.prefs.remove("sapdev.password")
        except Exception:
            pass

    async def on_test(self, e):
        if not (self.user.strip() and self.password):
            self.conn_test = (False, "Enter the communication user and password first.")
        else:
            client = sc.SAPClient(sc.service_url("API_PURCHASEORDER_PROCESS_SRV"),
                                  self.user.strip(), self.password, timeout=30)
            self.conn_test = (False, "Connecting to DEV...")
            self.render_sidebar()
            self.page.update()
            self.conn_test = await asyncio.to_thread(client.test)
        self.render_sidebar()
        self.render()

    async def open_drawer(self, e=None):
        if self.page.drawer:
            await self.page.show_drawer()

    # ---- main area --------------------------------------------------------
    def hero(self):
        wide = self._wide
        badge = ft.Container(txt(f"{ENVIRONMENT} · my440263", 18 if wide else 14, TEXT,
                                 ft.FontWeight.W_800),
                             bgcolor=BADGE, border_radius=999,
                             padding=ft.Padding.symmetric(horizontal=18, vertical=8),
                             shadow=ft.BoxShadow(blur_radius=14, spread_radius=-6,
                                                 offset=ft.Offset(0, 4),
                                                 color=ft.Colors.with_opacity(0.4, "black")))
        title = [txt("SAP bulk uploads", 34 if wide else 26, "white", ft.FontWeight.W_800),
                 txt("Preview, post and track your S/4HANA uploads from one place.",
                     17 if wide else 14, ft.Colors.with_opacity(0.92, "white"))]
        top = [ft.Column(title, spacing=4, expand=True), badge]
        extra = []
        if not wide:
            who = self.user or "Sign in"
            ok = bool(self.conn_test and self.conn_test[0])
            extra = [ft.Container(
                ft.Row([ft.Icon(ft.Icons.MENU, color="white", size=20),
                        txt(f"Connection: {who}" + ("  ✓" if ok else ""), 14, "white",
                            ft.FontWeight.W_600, expand=True, no_wrap=True,
                            overflow=ft.TextOverflow.ELLIPSIS)], spacing=8),
                bgcolor=WHITE_SOFT, border=ft.Border.all(1, WHITE_LINE), border_radius=10,
                padding=ft.Padding.symmetric(horizontal=12, vertical=10),
                on_click=self.open_drawer, ink=True)]
        if wide:
            head = ft.Row(top, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16)
        else:
            head = ft.Column([ft.Column(title, spacing=4), ft.Row([badge])], spacing=12)
        return ft.Container(
            ft.Column([head, *extra], spacing=14),
            gradient=ft.LinearGradient(begin=ft.Alignment.CENTER_LEFT,
                                       end=ft.Alignment.CENTER_RIGHT,
                                       colors=HERO, stops=[0, 0.35, 0.7, 1]),
            border_radius=18,
            padding=ft.Padding.symmetric(horizontal=32 if wide else 20, vertical=28 if wide else 20),
            shadow=ft.BoxShadow(blur_radius=30, spread_radius=-12, offset=ft.Offset(0, 10),
                                color=ft.Colors.with_opacity(0.55, "#7C3AED")))

    def tab_bar(self):
        labels = [f"{t['icon']} {t['label']}" for t in TABS] + \
                 [f"{i} {l}" for _, i, l in EXTRA_TABS]
        items = []
        for i, label in enumerate(labels):
            sel = i == self.tab
            items.append(ft.Container(
                txt(label, 16, PRIMARY if sel else TEXT),
                padding=ft.Padding.only(left=6, right=6, top=10, bottom=10),
                border=ft.Border.only(bottom=ft.BorderSide(3, PRIMARY if sel else "transparent")),
                on_click=lambda e, i=i: self.on_tab(i), ink=True))
        return ft.Container(ft.Row(items, spacing=22, scroll=ft.ScrollMode.AUTO),
                            border=ft.Border.only(bottom=ft.BorderSide(2, LINE)))

    def on_tab(self, i):
        self.tab = i
        self.render()

    def render(self):
        n = len(TABS)
        if self.tab < n:
            content = self.upload_view(self.tab)
        elif EXTRA_TABS[self.tab - n][0] == "metadata":
            content = self.metadata_view()
        else:
            content = self.history_view()
        footer = ft.Container(
            txt("Prepared by: Somesh Perni  |  somesh.perni@incresol.com", 14, "#475569",
                ft.FontWeight.W_500, text_align=ft.TextAlign.CENTER),
            border=ft.Border.only(top=ft.BorderSide(1, "#E5E7EB")),
            padding=ft.Padding.only(top=14, bottom=20), alignment=ft.Alignment.CENTER,
            margin=ft.Margin.only(top=36))
        pad = 40 if self._wide else 14
        column = ft.Column([self.hero(), self.tab_bar(), content, footer], spacing=22)
        self.body.content = ft.ListView(
            [ft.Container(ft.Container(column, width=1220 if self._wide else None),
                          alignment=ft.Alignment.TOP_CENTER,
                          padding=ft.Padding.symmetric(horizontal=pad, vertical=pad - 6))],
            expand=True)
        self.page.update()

    def toast(self, msg, error=False):
        self.page.show_dialog(ft.SnackBar(txt(msg, 14, "white"),
                                          bgcolor="#9B1C1C" if error else TEXT))

    async def save_bytes(self, name: str, data: bytes):
        try:
            path = await self.picker.save_file(dialog_title="Save file", file_name=name,
                                               src_bytes=data)
        except Exception as ex:
            self.toast(f"Couldn't save {name}: {ex}", error=True)
            return
        if not path:
            return
        if not self.page.platform.is_mobile() and not Path(path).exists():
            Path(path).write_bytes(data)
        self.toast(f"Saved {Path(path).name}")

    # =====================================================================
    # Upload tab (mirrors upload_tab() in the dashboard)
    # =====================================================================
    def upload_view(self, tab_index):
        cfg = TABS[tab_index]
        slug = self.slug_for_tab[tab_index]
        module = self.modules[slug]
        s = self.state[slug]
        s.setdefault("url", module.BASE_URL)
        out = []

        if len(cfg["scripts"]) > 1:
            out.append(ft.Dropdown(
                label=cfg["picker"], value=slug, expand=True, border_radius=8,
                options=[ft.DropdownOption(key=k, text=f"{self.modules[k].ICON} "
                                                       f"{self.modules[k].TITLE}")
                         for k in cfg["scripts"]],
                on_select=lambda e, t=tab_index: self.on_pick_script(t, e.control.value)))

        out += [
            h2(f"{module.ICON} {module.TITLE}"),
            ft.Row([code_chip(f"scripts/{module.__name__}.py"),
                    txt(f"· {module.TCODE} · {module.SERVICE}", 16)], wrap=True, spacing=6),
            caption(module.DESCRIPTION),
        ]

        url_field = ft.TextField(label="Service URL", value=s["url"], bgcolor=BG2,
                                 border_color="transparent", border_radius=8, dense=True,
                                 tooltip="Defaults to BASE_URL from the script (DEV). Changing "
                                         "it here affects this session only.",
                                 on_change=lambda e: s.__setitem__("url", e.control.value or ""))
        multi_sheet = hasattr(module, "template_xlsx_bytes")
        t1 = outline_button("Excel template (4 sheets)" if multi_sheet else "Excel template",
                            lambda e, m=module: self.page.run_task(
                                self.save_bytes, f"{m.SLUG}_template.xlsx",
                                sc.module_template_xlsx(m)), expand=True)
        t2 = outline_button("CSV template",
                            lambda e, m=module: self.page.run_task(
                                self.save_bytes, f"{m.SLUG}_template.csv",
                                sc.template_csv(m.COLUMNS, m.SAMPLE_ROWS)), expand=True)
        out.append(ft.ResponsiveRow([
            ft.Container(url_field, col={"xs": 12, "md": 7}),
            ft.Container(ft.Row([t1, t2], spacing=10), col={"xs": 12, "md": 5}),
        ], vertical_alignment=ft.CrossAxisAlignment.END, run_spacing=10))

        guide = pd.DataFrame([{"Column": c, "What to enter": module.COLUMN_NOTES.get(c, "")}
                              for c in module.COLUMNS])
        out.append(expander("Column guide", [table(guide, status_col=None)]))

        out.append(txt(f"Upload CSV or Excel (rows grouped by {module.GROUP_LABEL})", 15))
        file_label = s.get("file_name") or "CSV, XLSX, XLSM"
        out.append(ft.Container(
            ft.Row([ft.OutlinedButton("Upload", icon=ft.Icons.UPLOAD,
                                      on_click=lambda e, sl=slug: self.page.run_task(
                                          self.on_upload, sl),
                                      style=ft.ButtonStyle(
                                          color=TEXT, bgcolor="white",
                                          shape=ft.RoundedRectangleBorder(radius=8),
                                          side=ft.BorderSide(1, "#D5DBE1"),
                                          padding=ft.Padding.symmetric(horizontal=18,
                                                                       vertical=16))),
                    txt(file_label, 15, MUTED if not s.get("file_name") else TEXT,
                        expand=True, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS),
                    *([ft.IconButton(ft.Icons.CLOSE, icon_color=MUTED, tooltip="Remove file",
                                     on_click=lambda e, sl=slug: self.clear_file(sl))]
                      if s.get("file_name") else [])], spacing=16),
            bgcolor=BG2, border_radius=10, padding=ft.Padding.symmetric(horizontal=18,
                                                                        vertical=16)))

        if s.get("df_in") is None:
            out.append(note("Upload a filled template to preview the payloads before posting."))
            return ft.Column(out, spacing=16)

        out += self.preview_controls(slug, module, s)
        return ft.Column(out, spacing=16)

    def on_pick_script(self, tab_index, slug):
        self.slug_for_tab[tab_index] = slug
        self.render()

    def clear_file(self, slug):
        url = self.state[slug].get("url")
        self.state[slug] = {"url": url}
        self.render()

    async def on_upload(self, slug):
        module = self.modules[slug]
        files = await self.picker.pick_files(
            dialog_title="Choose a filled template", file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["csv", "xlsx", "xlsm"], with_data=True)
        if not files:
            return
        f = files[0]
        raw = f.bytes
        if raw is None and f.path:
            raw = Path(f.path).read_bytes()
        if raw is None:
            self.toast("Couldn't read that file. Save it to the device first.", error=True)
            return
        s = self.state[slug]
        url = s.get("url", module.BASE_URL)
        self.state[slug] = s = {"url": url, "file_name": f.name,
                                "digest": hashlib.sha1(raw).hexdigest()}
        try:
            df_in = sc.module_read_upload(module, raw, f.name)
        except Exception as e:  # noqa: BLE001
            s["read_error"] = f"{f.name} could not be read: {e}"
            s["df_in"] = pd.DataFrame()
            self.render()
            return
        s["unknown"] = [c for c in df_in.columns if c not in module.COLUMNS]
        for c in module.COLUMNS:
            if c not in df_in.columns:
                df_in[c] = ""
        s["df_in"] = df_in
        s["edited"] = clean_frame(df_in[module.COLUMNS])
        self.rebuild(slug)
        self.render()

    def rebuild(self, slug):
        s, module = self.state[slug], self.modules[slug]
        s.pop("build_error", None)
        s["confirmed"] = False
        s["docs"] = []
        if s["edited"].empty:
            return
        try:
            s["docs"] = module.build_documents(s["edited"])
        except Exception as e:  # noqa: BLE001
            s["build_error"] = f"The file could not be turned into payloads: " \
                               f"{type(e).__name__}: {e}"
        if s["docs"] and s.get("chosen") not in [d["ref"] for d in s["docs"]]:
            s["chosen"] = s["docs"][0]["ref"]

    def preview_controls(self, slug, module, s):
        out = []
        if s.get("read_error"):
            return [note(s["read_error"], "error")]
        if s.get("unknown"):
            out.append(note("These columns are not in the template and are ignored: "
                            + ", ".join(s["unknown"]), "warning"))
        flat_note = (" - workbook shown as one row per component"
                     if hasattr(module, "read_upload") and s["file_name"].lower().endswith(
                         (".xlsx", ".xlsm")) else "")
        out.append(expander(
            f"Raw file ({s['file_name']}){flat_note} - tap a cell to correct it before posting",
            [table(s["edited"], status_col=None, max_rows=300,
                   on_cell_tap=lambda i, c, sl=slug: self.edit_cell(sl, i, c))]))
        if s["edited"].empty:
            out.append(note("The file has no data rows.", "warning"))
            return out
        if s.get("build_error"):
            out.append(note(s["build_error"], "error"))
            return out

        docs = s["docs"]
        valid = [d for d in docs if not d["errors"]]
        out.append(stat_cards([
            ("Documents ready", len(valid), "purple"),
            (module.ITEMS_LABEL, sum(module.count_items(d["payload"]) for d in valid), "blue"),
            ("Invalid", len(docs) - len(valid), "amber" if len(valid) < len(docs) else "green")]))
        summary = pd.DataFrame([{module.GROUP_LABEL: d["ref"], "Rows": d["lines"],
                                 module.ITEMS_LABEL: module.count_items(d["payload"]),
                                 "Status": "Invalid" if d["errors"] else "Ready",
                                 "Problems": " | ".join(d["errors"])} for d in docs])
        out.append(table(summary))

        chosen = s.get("chosen")
        payload = next((d["payload"] for d in docs if d["ref"] == chosen), {})
        pick = ft.Dropdown(label="Inspect JSON payload", value=chosen, border_radius=8,
                           options=[ft.DropdownOption(key=d["ref"], text=d["ref"]) for d in docs],
                           on_select=lambda e, sl=slug: self.on_choose(sl, e.control.value),
                           expand=True)
        dl = outline_button("Download all payloads", lambda e, sl=slug: self.page.run_task(
            self.save_bytes, f"{sl}_payloads_{datetime.now():%Y%m%d_%H%M%S}.json",
            json.dumps({d["ref"]: d["payload"] for d in docs}, indent=2,
                       default=str).encode("utf-8")))
        out.append(ft.ResponsiveRow([ft.Container(pick, col={"xs": 12, "md": 9}),
                                     ft.Container(dl, col={"xs": 12, "md": 3})],
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                    run_spacing=10))
        out.append(expander(f"JSON payload: {chosen}", [ft.Container(
            txt(json.dumps(payload, indent=2, default=str), 13, TEXT, selectable=True,
                font_family="monospace"),
            bgcolor=BG2, border_radius=8, padding=14)]))

        # --- Post (same wording and order as the dashboard) -------------------
        out.append(ft.Divider(color=LINE))
        url = (s.get("url") or module.BASE_URL).rstrip("/")
        out.append(ft.Row([txt(f"Post to SAP {ENVIRONMENT} at", 16, TEXT, ft.FontWeight.W_700),
                           code_chip(url)], wrap=True, spacing=8))
        self.confirm_box = ft.Checkbox(
            label=f"I've checked these {len(valid)} payload(s) and want to create them in "
                  f"{ENVIRONMENT}", value=bool(s.get("confirmed")),
            on_change=lambda e, sl=slug: self.on_confirm(sl, e.control.value),
            disabled=bool(s.get("posting")))
        stop_box = ft.Checkbox(label="Stop at the first failure", value=bool(s.get("stop")),
                               on_change=lambda e, sl=slug: self.state[sl].__setitem__(
                                   "stop", bool(e.control.value)))
        out.append(ft.ResponsiveRow([ft.Container(self.confirm_box, col={"xs": 12, "md": 7}),
                                     ft.Container(stop_box, col={"xs": 12, "md": 5})]))
        if len(valid) < len(docs):
            out.append(caption(f"{len(docs) - len(valid)} invalid document(s) will be skipped "
                               f"and listed in the results."))
        self.post_btn = primary_button(
            f"Create {len(valid)} document(s) in SAP {ENVIRONMENT}",
            lambda e, sl=slug: self.page.run_task(self.on_post, sl),
            disabled=not (s.get("confirmed") and valid) or bool(s.get("posting")))
        out.append(ft.Row([self.post_btn]))
        self.post_msg = ft.Column([], spacing=8)
        if s.get("post_error"):
            self.post_msg.controls.append(note(s["post_error"], "error"))
        out.append(self.post_msg)
        self.progress_text = txt(s.get("progress_text", ""), 15, TEXT,
                                 visible=bool(s.get("progress_text")))
        self.progress = ft.ProgressBar(value=s.get("progress", 0), color=PRIMARY, bgcolor=BG2,
                                       bar_height=8, border_radius=4,
                                       visible=bool(s.get("progress_text")))
        out += [self.progress_text, self.progress]
        if s.get("stopped_note"):
            out.append(note(s["stopped_note"], "warning"))
        if s.get("saved_to"):
            out.append(caption(f"Saved to {s['saved_to']}"))

        res = s.get("results")
        if res is not None:
            out.append(h3("Results of this run"))
            out += self.results_controls(res, slug)
        return out

    def results_controls(self, df, key):
        c = status_counts(df)
        return [stat_cards([("Posted", c["success"], "green"), ("Failed", c["failed"], "red"),
                            ("Invalid (not sent)", c["invalid"], "amber"),
                            ("Documents in run", c["total"], "blue")]),
                table(df),
                ft.Row([outline_button("Download results CSV", lambda e: self.page.run_task(
                    self.save_bytes, f"{key}_results_{datetime.now():%Y%m%d_%H%M%S}.csv",
                    df.to_csv(index=False).encode("utf-8")))])]

    def on_choose(self, slug, ref):
        self.state[slug]["chosen"] = ref
        self.render()

    def on_confirm(self, slug, value):
        s = self.state[slug]
        s["confirmed"] = bool(value)
        valid = [d for d in s.get("docs", []) if not d["errors"]]
        self.post_btn.disabled = not (s["confirmed"] and valid) or bool(s.get("posting"))
        self.page.update()

    def edit_cell(self, slug, row_index, column):
        s = self.state[slug]
        field = ft.TextField(value=str(s["edited"].at[row_index, column]), multiline=True,
                             min_lines=1, max_lines=6, autofocus=True, width=460)

        def save(e):
            s["edited"].at[row_index, column] = field.value or ""
            s["edited"] = clean_frame(s["edited"])
            self.page.pop_dialog()
            self.rebuild(slug)
            self.render()

        self.page.show_dialog(ft.AlertDialog(
            title=txt(f"Line {row_index + 2}: {column}", 18, TEXT, ft.FontWeight.W_600),
            content=field,
            actions=[ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()),
                     ft.TextButton("Save", on_click=save)]))

    async def on_post(self, slug):
        s, module = self.state[slug], self.modules[slug]
        s.pop("post_error", None)
        s.pop("stopped_note", None)
        s.pop("saved_to", None)
        if not self.user.strip() or not self.password:
            s["post_error"] = ("Enter the communication user and password in the "
                               + ("sidebar" if self._wide else "Connection menu")
                               + ", then try again.")
            self.render()
            return
        base_url = (s.get("url") or module.BASE_URL).rstrip("/")
        docs = s["docs"]
        s["posting"] = True
        s["progress"], s["progress_text"] = 0.0, "Starting"
        self.render()

        client = sc.SAPClient(base_url, self.user.strip(), self.password)
        ok, msg = await asyncio.to_thread(client.test)
        if not ok:
            s["posting"] = False
            s["progress_text"] = ""
            s["post_error"] = f"Could not fetch a CSRF token, so nothing was posted: {msg}"
            self.render()
            return

        live = {"i": 0, "n": len(docs), "text": "Starting", "rows": None, "error": None}

        def progress(i, n, r):
            live.update(i=i, n=n, text=f"Posting {r['Group']} ({i} of {n}): {r['Status']}")

        def work():
            try:
                live["rows"] = sc.post_documents(client, module, docs, bool(s.get("stop")),
                                                 progress)
            except Exception as e:  # noqa: BLE001
                live["error"] = f"{type(e).__name__}: {e}"

        th = threading.Thread(target=work, daemon=True)
        th.start()
        while th.is_alive():
            await asyncio.sleep(0.25)
            self.progress.value = live["i"] / max(1, live["n"])
            self.progress_text.value = live["text"]
            self.page.update()

        s["posting"] = False
        if live["error"]:
            s["post_error"] = live["error"]
            s["progress_text"] = ""
            self.render()
            return
        rows = live["rows"] or []
        s["progress"], s["progress_text"] = 1.0, "Finished"
        if s.get("stop") and len(rows) < len(docs):
            s["stopped_note"] = (f"Stopped after a failure. {len(docs) - len(rows)} document(s) "
                                 "were not attempted.")
        try:
            s["saved_to"] = str(save_results(slug, rows))
        except Exception as e:  # noqa: BLE001
            s["post_error"] = f"Posted, but the results file could not be saved: {e}"
        s["results"] = pd.DataFrame(rows, columns=RESULT_COLUMNS)
        s["confirmed"] = False
        self.render()

    # =====================================================================
    # Metadata tab
    # =====================================================================
    def metadata_view(self):
        if self.metas is None:
            self.metas = sc.all_metadata()
        out = [h2("🔎 DEV metadata")]
        if not self.metas:
            out.append(note(f"No metadata JSON files found in {sc.METADATA_DIR}", "error"))
            return ft.Column(out, spacing=16)
        services = list(self.metas)
        if self.meta_service not in services:
            self.meta_service = services[0]
        md = self.metas[self.meta_service]
        sets = list(md.entity_sets)
        if self.meta_set not in sets:
            self.meta_set = sets[0]
        s = md.entity_sets[self.meta_set]
        et_name = s["EntityTypeName"]
        et = md.entity_types[et_name]

        def yes(k):
            return "Yes" if s.get(k, "true") != "false" else "No"

        out.append(ft.ResponsiveRow([
            ft.Container(ft.Dropdown(label="Service", value=self.meta_service, border_radius=8,
                                     options=[ft.DropdownOption(key=x, text=x) for x in services],
                                     on_select=self.on_meta_service, expand=True),
                         col={"xs": 12, "md": 6}),
            ft.Container(ft.Dropdown(label="Entity set", value=self.meta_set, border_radius=8,
                                     options=[ft.DropdownOption(key=x, text=x) for x in sets],
                                     on_select=self.on_meta_set, expand=True),
                         col={"xs": 12, "md": 6}),
        ], run_spacing=10))
        out.append(stat_cards([("Fields", len(et["props"]), "purple"),
                               ("Create", yes("creatable"), "blue"),
                               ("Update", yes("updatable"), "blue"),
                               ("Delete", yes("deletable"), "blue")]))
        out.append(caption(f"{sc.service_url(self.meta_service)}/{self.meta_set} · entity type "
                           f"{et_name} · keys: {', '.join(et['keys'])}"))
        if et["navs"]:
            rows = []
            for n in et["navs"]:
                try:
                    target, many = md.nav_target(et_name, n)
                except KeyError:
                    target, many = "?", False
                rows.append({"Navigation": n, "Target": target,
                             "Cardinality": "many" if many else "one"})
            out.append(expander(f"Navigation properties ({len(rows)}) - used for deep insert",
                                [table(pd.DataFrame(rows), status_col=None)]))
        out.append(ft.ResponsiveRow([
            ft.Container(ft.TextField(label="Filter fields", value=self.meta_filter,
                                      hint_text="name, label or description", bgcolor=BG2,
                                      border_color="transparent", border_radius=8,
                                      on_submit=self.on_meta_filter, on_blur=self.on_meta_filter),
                         col={"xs": 9, "md": 9}),
            ft.Container(ft.Checkbox(label="Creatable only", value=self.meta_creatable,
                                     on_change=self.on_meta_creatable), col={"xs": 3, "md": 3}),
        ], vertical_alignment=ft.CrossAxisAlignment.CENTER))
        df = pd.DataFrame(md.fields_table(et_name))
        if self.meta_creatable:
            df = df[(df["Creatable"] == "X") | (df["Key"] == "X")]
        q = self.meta_filter.strip().lower()
        if q:
            df = df[df.apply(lambda r: q in " ".join(map(str, r.values)).lower(), axis=1)]
        out.append(table(df, status_col=None, max_rows=400))
        out.append(ft.Row([outline_button("Download field list (CSV)", lambda e, d=df:
                                          self.page.run_task(
                                              self.save_bytes,
                                              f"{self.meta_service}_{self.meta_set}_fields.csv",
                                              d.to_csv(index=False).encode("utf-8-sig")))]))
        return ft.Column(out, spacing=16)

    def on_meta_service(self, e):
        self.meta_service, self.meta_set = e.control.value, None
        self.render()

    def on_meta_set(self, e):
        self.meta_set = e.control.value
        self.render()

    def on_meta_filter(self, e):
        if (e.control.value or "") != self.meta_filter:
            self.meta_filter = e.control.value or ""
            self.render()

    def on_meta_creatable(self, e):
        self.meta_creatable = bool(e.control.value)
        self.render()

    # =====================================================================
    # History tab (same figures as the dashboard)
    # =====================================================================
    def history_view(self):
        files = sorted(RESULTS_DIR.glob("*.csv"), reverse=True) if RESULTS_DIR.exists() else []
        if not files:
            return ft.Column([note("No runs yet. Results appear here after your first post to "
                                   "SAP DEV.")])
        frames = []
        for f in files:
            try:
                df = pd.read_csv(f, dtype=str).fillna("")
                df["Run file"] = f.name
                frames.append(df)
            except Exception:  # noqa: BLE001
                continue
        if not frames:
            return ft.Column([note("Result files exist but none could be read.", "warning")])
        hist = pd.concat(frames, ignore_index=True)
        for col in RESULT_COLUMNS:
            if col not in hist:
                hist[col] = ""

        c = status_counts(hist)
        sent = c["total"] - c["invalid"]
        rate = f"{(c['success'] / sent * 100):.0f}%" if sent else "-"
        out = [stat_cards([("Runs", len(files), "purple"),
                           ("Documents posted", c["success"], "green"),
                           ("Failures", c["failed"], "red"), ("Success rate", rate, "blue")])]

        hist["Outcome"] = hist["Status"].map(
            lambda v: "Posted" if v == "Success" else ("Invalid" if v == "Invalid" else "Failed"))
        out.append(h3("By document type"))
        out.append(self.bar_chart(hist))

        out.append(h3("All posted and failed documents"))
        mods = ["All"] + sorted(hist["Module"].unique())
        if self.hist_module not in mods:
            self.hist_module = "All"
        out.append(ft.ResponsiveRow([
            ft.Container(ft.Dropdown(label="Document type", value=self.hist_module,
                                     border_radius=8, expand=True,
                                     options=[ft.DropdownOption(key=m, text=m) for m in mods],
                                     on_select=self.on_hist_module), col={"xs": 12, "md": 4}),
            ft.Container(ft.Dropdown(label="Outcome", value=self.hist_outcome, border_radius=8,
                                     expand=True,
                                     options=[ft.DropdownOption(key=o, text=o) for o in
                                              ("All", "Posted", "Failed", "Invalid")],
                                     on_select=self.on_hist_outcome), col={"xs": 12, "md": 3}),
            ft.Container(ft.TextField(label="Search group, document number or message",
                                      value=self.hist_text, bgcolor=BG2,
                                      border_color="transparent", border_radius=8,
                                      on_submit=self.on_hist_text, on_blur=self.on_hist_text),
                         col={"xs": 12, "md": 5}),
        ], run_spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER))
        view = hist
        if self.hist_module != "All":
            view = view[view["Module"] == self.hist_module]
        if self.hist_outcome != "All":
            view = view[view["Outcome"] == self.hist_outcome]
        if self.hist_text:
            mask = view[["Group", "Document", "Detail"]].apply(
                lambda col: col.str.contains(self.hist_text, case=False, regex=False)).any(axis=1)
            view = view[mask]
        out.append(table(view[RESULT_COLUMNS + ["Run file"]]))

        failures = view[view["Outcome"] == "Failed"]
        if self.hist_outcome in ("All", "Failed") and not failures.empty:
            out.append(h3("Most common failure messages"))
            top = failures["Detail"].map(main_failure_message).value_counts().head(10)
            out.append(table(top.rename_axis("Message").reset_index(name="Count"),
                             status_col=None))
        return ft.Column(out, spacing=16)

    def bar_chart(self, hist):
        """Stacked bars per document type: Posted / Failed / Invalid (dashboard colours)."""
        colours = {"Posted": "#1E7B4F", "Failed": "#B4232C", "Invalid": "#D97706"}
        by = hist.groupby(["Module", "Outcome"]).size().unstack(fill_value=0)
        for col in colours:
            if col not in by:
                by[col] = 0
        biggest = max(1, int(by[list(colours)].sum(axis=1).max()))
        rows = []
        for module, r in by.iterrows():
            total = int(sum(r[k] for k in colours))
            segs = [ft.Container(expand=int(r[k]), bgcolor=colours[k], height=26,
                                 tooltip=f"{k}: {int(r[k])}")
                    for k in colours if int(r[k]) > 0]
            if total < biggest:
                segs.append(ft.Container(expand=biggest - total, height=26))
            rows.append(ft.Column([
                ft.Row([txt(module, 14, TEXT, expand=True), txt(total, 14, MUTED)]),
                ft.Container(ft.Row(segs, spacing=0), border_radius=4,
                             clip_behavior=ft.ClipBehavior.HARD_EDGE)], spacing=4))
        legend = ft.Row([ft.Row([ft.Container(width=14, height=14, bgcolor=v), txt(k, 14, MUTED)],
                                spacing=6) for k, v in colours.items()], spacing=18)
        return ft.Column(rows + [legend], spacing=14)

    def on_hist_module(self, e):
        self.hist_module = e.control.value
        self.render()

    def on_hist_outcome(self, e):
        self.hist_outcome = e.control.value
        self.render()

    def on_hist_text(self, e):
        if (e.control.value or "") != self.hist_text:
            self.hist_text = e.control.value or ""
            self.render()


async def main(page: ft.Page):
    app = App(page)
    await app.start()


if __name__ == "__main__":
    ft.run(main, assets_dir="assets")
