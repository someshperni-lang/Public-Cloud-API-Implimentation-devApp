"""
SAP DEV Uploads - Windows and Android app (Flet).

Pick an upload type, choose a filled CSV/Excel template, check what will be
sent, post to the DEV tenant and save the results. Runs the same code on a
Windows PC and an Android phone.
"""
from __future__ import annotations

import asyncio
import json
import os
from datetime import datetime
from pathlib import Path

import flet as ft

from core import sap_common as sc
from core.uploads import BY_SLUG, UPLOADS

# ----------------------------------------------------------------------------
# Look
# ----------------------------------------------------------------------------
INK = "#1F2A33"          # text, headers
MUTED = "#5B6770"        # secondary text
SURFACE = "#F4F6F8"      # app background
CARD = "#FFFFFF"
LINE = "#DCE1E5"
DEV = "#F2A900"          # the DEV marker: the one loud colour in the app
STATUS = {               # (text colour, background)
    "READY": ("#1E4E8C", "#E3ECF8"),
    "SUCCESS": ("#1F6B43", "#DDF1E4"),
    "FAILED": ("#9B1C1C", "#F8DEDE"),
    "ERROR": ("#9B1C1C", "#F8DEDE"),
    "INVALID": ("#7A4F00", "#FBEBC8"),
}
WIDE = 820               # wider than this: side rail; narrower: bottom bar

PREF_KEYS = ("host", "client", "user", "password", "remember")


def data_dir() -> Path:
    d = Path(os.getenv("FLET_APP_STORAGE_DATA") or Path.home() / ".sap_dev_uploads")
    d.mkdir(parents=True, exist_ok=True)
    return d


def pill(status: str) -> ft.Container:
    fg, bg = STATUS.get(status, (INK, LINE))
    return ft.Container(ft.Text(status.capitalize(), size=12, weight=ft.FontWeight.W_600, color=fg),
                        bgcolor=bg, border_radius=10,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=3))


def card(content, padding=16) -> ft.Container:
    return ft.Container(content, bgcolor=CARD, border_radius=12,
                        border=ft.Border.all(1, LINE), padding=padding)


def heading(text: str) -> ft.Text:
    return ft.Text(text, size=18, weight=ft.FontWeight.W_600, color=INK)


def stat(label: str, value, color=INK) -> ft.Container:
    return ft.Container(
        ft.Column([ft.Text(str(value), size=22, weight=ft.FontWeight.W_700, color=color),
                   ft.Text(label, size=12, color=MUTED)], spacing=0, tight=True),
        padding=ft.Padding.symmetric(horizontal=14, vertical=10),
        bgcolor=SURFACE, border_radius=10, expand=True)


# ----------------------------------------------------------------------------
# App
# ----------------------------------------------------------------------------
class App:
    def __init__(self, page: ft.Page):
        self.page = page
        self.prefs = ft.SharedPreferences()
        self.picker = ft.FilePicker()
        page.services.extend([self.prefs, self.picker])

        self.conn = {"host": sc.DEFAULT_HOST, "client": "", "user": "", "password": "",
                     "remember": False}
        self.section = 0
        self.module = UPLOADS[0]
        self.table: sc.Table | None = None
        self.file_name = ""
        self.docs: list[dict] = []
        self.results: list[dict] | None = None
        self.posting = False
        self.history_file = data_dir() / "history.json"

        page.title = "SAP DEV Uploads"
        page.bgcolor = SURFACE
        page.padding = 0
        page.theme = ft.Theme(color_scheme_seed=INK)
        page.theme_mode = ft.ThemeMode.LIGHT
        page.on_resize = self.on_resize
        self._wide = None

    # ---- startup --------------------------------------------------------
    async def start(self):
        for k in PREF_KEYS:
            try:
                v = await self.prefs.get(f"sapdev.{k}")
            except Exception:
                v = None
            if v not in (None, ""):
                self.conn[k] = v
        self.conn["remember"] = bool(self.conn.get("remember"))
        if not (self.conn["user"] and self.conn["password"]):
            self.section = 2                        # first run: go to Connection
        self.layout(force=True)

    # ---- layout ---------------------------------------------------------
    def on_resize(self, e):
        wide = (self.page.width or 1000) >= WIDE
        if wide != self._wide:
            self.layout()
        elif getattr(self, "box", None) is not None:
            self.box.width = self.box_width()
            self.page.update()

    def layout(self, force=False):
        wide = (self.page.width or 1000) >= WIDE
        if not force and wide == self._wide:
            return
        self._wide = wide
        self.body = ft.Container(expand=True)
        dests = [("Upload", ft.Icons.UPLOAD_FILE_OUTLINED, ft.Icons.UPLOAD_FILE),
                 ("History", ft.Icons.HISTORY, ft.Icons.HISTORY),
                 ("Connection", ft.Icons.LINK, ft.Icons.LINK)]
        self.page.controls.clear()
        self.page.navigation_bar = None
        if wide:
            rail = ft.NavigationRail(
                selected_index=self.section, label_type=ft.NavigationRailLabelType.ALL,
                bgcolor=CARD, min_width=96, on_change=self.on_nav,
                destinations=[ft.NavigationRailDestination(icon=i, selected_icon=s, label=l)
                              for l, i, s in dests])
            main = ft.Row([rail, ft.VerticalDivider(width=1, color=LINE), self.body],
                          expand=True, spacing=0)
        else:
            self.page.navigation_bar = ft.NavigationBar(
                selected_index=self.section, on_change=self.on_nav, bgcolor=CARD,
                destinations=[ft.NavigationBarDestination(icon=i, selected_icon=s, label=l)
                              for l, i, s in dests])
            main = self.body
        self.dev_bar = ft.Container()
        self.page.add(ft.SafeArea(ft.Column([self.dev_bar, main], spacing=0, expand=True),
                                  expand=True))
        self.render()

    def on_nav(self, e):
        self.section = int(e.control.selected_index)
        self.render()

    def render(self):
        self.dev_bar.content = self.build_dev_bar()
        self.dev_bar.bgcolor = DEV
        self.dev_bar.padding = ft.Padding.symmetric(horizontal=16, vertical=8)
        view = [self.build_upload, self.build_history, self.build_connection][self.section]()
        self.body.content = ft.ListView([ft.Container(view, padding=ft.Padding.all(16),
                                                      alignment=ft.Alignment.TOP_CENTER)],
                                        expand=True)
        self.page.update()

    def build_dev_bar(self):
        who = self.conn["user"] or "not signed in"
        return ft.Row([
            ft.Container(ft.Text("DEV", size=13, weight=ft.FontWeight.W_800, color=DEV),
                         bgcolor=INK, border_radius=4,
                         padding=ft.Padding.symmetric(horizontal=8, vertical=2)),
            ft.Text(f"{self.conn['host']}", size=13, weight=ft.FontWeight.W_600, color=INK,
                    expand=True, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS),
            ft.Text(who, size=13, color=INK, no_wrap=True),
        ], spacing=10)

    def toast(self, msg: str, error=False):
        self.page.show_dialog(ft.SnackBar(ft.Text(msg, color="white"),
                                          bgcolor=STATUS["FAILED"][0] if error else INK))

    def width_box(self, controls) -> ft.Container:
        """Content column, capped at 900 px on wide screens so lines stay readable."""
        self.box = ft.Container(ft.Column(controls, spacing=16), width=self.box_width())
        return self.box

    def box_width(self):
        w = self.page.width or 0
        return min(900, w - 140) if w >= WIDE else None

    # =====================================================================
    # Upload
    # =====================================================================
    def build_upload(self):
        m = self.module
        md_missing = sc.load_metadata(m.SERVICE) is None

        type_dd = ft.Dropdown(
            label="What are you uploading?", value=m.SLUG, on_select=self.on_type,
            options=[ft.DropdownOption(key=u.SLUG, text=u.TITLE) for u in UPLOADS],
            expand=True)

        guide = ft.ExpansionTile(
            title=ft.Text("Column guide", size=14, color=INK),
            tile_padding=ft.Padding.symmetric(horizontal=4),
            controls=[ft.Container(ft.Column([
                ft.Row([ft.Text(c, size=13, weight=ft.FontWeight.W_600, color=INK, width=210,
                                selectable=True),
                        ft.Text(m.COLUMN_NOTES.get(c, ""), size=13, color=MUTED, expand=True)],
                       vertical_alignment=ft.CrossAxisAlignment.START)
                for c in m.COLUMNS], spacing=6), padding=ft.Padding.only(left=4, bottom=8))])

        top = [ft.Row([type_dd]),
               ft.Text(f"{m.DESCRIPTION}  Sends to {m.SERVICE}.", size=13, color=MUTED)]
        if md_missing:
            top.append(ft.Container(ft.Text(
                f"Field checks are off: metadata for {m.SERVICE} isn't bundled in this build. "
                "Dates and yes/no values are still converted. Add the JSON file to the "
                "metadata folder and rebuild to turn checks on.", size=12, color=STATUS["INVALID"][0]),
                bgcolor=STATUS["INVALID"][1], border_radius=8, padding=10))
        top += [
            ft.Row([
                ft.FilledButton("Choose file", icon=ft.Icons.FOLDER_OPEN, on_click=self.on_pick,
                                disabled=self.posting),
                ft.OutlinedButton("Excel template", icon=ft.Icons.DOWNLOAD,
                                  on_click=self.on_template_xlsx),
                ft.OutlinedButton("CSV template", icon=ft.Icons.DOWNLOAD,
                                  on_click=self.on_template_csv),
            ], wrap=True, spacing=8, run_spacing=8),
            guide,
        ]
        controls = [card(ft.Column(top, spacing=12))]
        if self.table is not None:
            controls.append(self.build_check())
            if self.results is not None:
                controls.append(self.build_results())
        return self.width_box(controls)

    def build_check(self):
        valid = [d for d in self.docs if not d["errors"]]
        invalid = [d for d in self.docs if d["errors"]]
        tiles = []
        for d in self.docs:
            status = "INVALID" if d["errors"] else "READY"
            body = [ft.Text(e, size=13, color=STATUS["INVALID"][0], selectable=True)
                    for e in d["errors"]]
            body.append(ft.TextButton("View what will be sent", icon=ft.Icons.DATA_OBJECT,
                                      on_click=lambda e, d=d: self.show_payload(d)))
            tiles.append(ft.ExpansionTile(
                title=ft.Text(d["ref"], size=14, weight=ft.FontWeight.W_600, color=INK),
                subtitle=ft.Text(f"Sheet lines {d['lines']}", size=12, color=MUTED),
                leading=pill(status),
                controls=[ft.Container(ft.Column(body, spacing=4),
                                       padding=ft.Padding.only(left=16, right=16, bottom=8))]))

        self.confirm = ft.Checkbox(
            label=f"Post {len(valid)} document(s) to {self.conn['host']}. "
                  "Invalid ones are skipped.", value=False, on_change=self.on_confirm,
            disabled=self.posting or not valid)
        self.post_btn = ft.FilledButton(
            f"Post {len(valid)} to DEV", icon=ft.Icons.SEND, on_click=self.on_post,
            disabled=True, style=ft.ButtonStyle(bgcolor={ft.ControlState.DEFAULT: INK,
                                                         ft.ControlState.DISABLED: LINE},
                                                color={ft.ControlState.DEFAULT: "white"}))
        self.progress = ft.ProgressBar(value=0, visible=False, color=DEV, bgcolor=LINE)
        self.progress_text = ft.Text("", size=12, color=MUTED, visible=False)

        signed_in = bool(self.conn["user"] and self.conn["password"])
        post_area = [self.confirm, ft.Row([self.post_btn]), self.progress, self.progress_text]
        if not signed_in:
            post_area.insert(0, ft.Text("Add your user and password under Connection to post.",
                                        size=13, color=STATUS["INVALID"][0]))

        return card(ft.Column([
            heading(self.file_name),
            ft.Row([stat("Rows", len(self.table)), stat("Documents", len(self.docs)),
                    stat("Ready", len(valid), STATUS["READY"][0]),
                    stat("Invalid", len(invalid), STATUS["INVALID"][0])], spacing=8),
            ft.Column(tiles, spacing=0),
            ft.Divider(color=LINE),
            *post_area,
        ], spacing=12))

    def build_results(self):
        res = self.results or []
        ok = sum(r["Status"] == "SUCCESS" for r in res)
        rows = []
        for r in res:
            rows.append(ft.Container(ft.Column([
                ft.Row([pill(r["Status"]),
                        ft.Text(r["Reference"], size=14, weight=ft.FontWeight.W_600, color=INK,
                                expand=True),
                        ft.Text(r["SAP Document"], size=14, weight=ft.FontWeight.W_600,
                                color=STATUS["SUCCESS"][0], selectable=True)], spacing=10),
                ft.Text(r["Message"], size=13, color=MUTED, selectable=True),
            ], spacing=4), padding=ft.Padding.symmetric(vertical=8),
                border=ft.Border.only(bottom=ft.BorderSide(1, LINE))))
        return card(ft.Column([
            heading(f"{ok} of {len(res)} created in DEV"),
            ft.Column(rows, spacing=0),
            ft.Row([
                ft.OutlinedButton("Save results (CSV)", icon=ft.Icons.SAVE_ALT,
                                  on_click=self.on_save_results),
                ft.OutlinedButton("Save sheet with status (Excel)", icon=ft.Icons.SAVE_ALT,
                                  on_click=self.on_save_status),
            ], wrap=True, spacing=8, run_spacing=8),
        ], spacing=12))

    # ---- upload actions --------------------------------------------------
    def on_type(self, e):
        self.module = BY_SLUG[e.control.value]
        self.table, self.docs, self.results, self.file_name = None, [], None, ""
        self.render()

    async def save_bytes(self, name: str, data: bytes):
        try:
            path = await self.picker.save_file(dialog_title="Save file", file_name=name,
                                               src_bytes=data)
        except Exception as ex:
            self.toast(f"Couldn't save {name}: {ex}", error=True)
            return
        if path is None:
            return
        if self.page.platform in (ft.PagePlatform.WINDOWS, ft.PagePlatform.LINUX,
                                  ft.PagePlatform.MACOS) and not Path(path).exists():
            Path(path).write_bytes(data)           # older desktop clients don't write bytes
        self.toast(f"Saved {Path(path).name}")

    async def on_template_xlsx(self, e):
        m = self.module
        await self.save_bytes(f"{m.SLUG}_template.xlsx",
                              sc.template_xlsx(m.COLUMNS, m.SAMPLE_ROWS, m.COLUMN_NOTES, m.TITLE))

    async def on_template_csv(self, e):
        m = self.module
        await self.save_bytes(f"{m.SLUG}_template.csv", sc.template_csv(m.COLUMNS, m.SAMPLE_ROWS))

    async def on_pick(self, e):
        files = await self.picker.pick_files(
            dialog_title="Choose a filled template", file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm", "csv"], with_data=True)
        if not files:
            return
        f = files[0]
        data = f.bytes
        if data is None and f.path:
            data = Path(f.path).read_bytes()
        if data is None:
            self.toast("Couldn't read that file. Try saving it to the device first.", error=True)
            return
        try:
            table = sc.read_table(data, f.name)
            docs = self.module.build_documents(table)
        except Exception as ex:
            self.toast(f"Couldn't read {f.name}: {ex}", error=True)
            return
        unknown = [c for c in table.columns if c not in self.module.COLUMNS]
        self.table, self.docs, self.results, self.file_name = table, docs, None, f.name
        self.render()
        if unknown:
            self.toast("Ignored columns not in the template: " + ", ".join(unknown[:6]))

    def show_payload(self, d):
        text = json.dumps(d["payload"], indent=2)
        dlg = ft.AlertDialog(
            title=ft.Text(d["ref"]),
            content=ft.Container(ft.Column([ft.Text(text, size=12, font_family="monospace",
                                                    selectable=True)], scroll=ft.ScrollMode.AUTO),
                                 width=560, height=440),
            actions=[ft.TextButton("Copy", on_click=lambda e: self.copy(text)),
                     ft.TextButton("Close", on_click=lambda e: self.page.pop_dialog())])
        self.page.show_dialog(dlg)

    async def copy(self, text):
        try:
            await ft.Clipboard().set(text)
            self.toast("Copied")
        except Exception:
            self.toast("Copy isn't available here", error=True)

    def on_confirm(self, e):
        self.post_btn.disabled = not (e.control.value and self.conn["user"]
                                      and self.conn["password"] and not self.posting)
        self.page.update()

    async def on_post(self, e):
        m = self.module
        self.posting = True
        self.post_btn.disabled = True
        self.confirm.disabled = True
        self.progress.visible = self.progress_text.visible = True
        self.page.update()

        client = sc.SAPClient(self.conn["host"], m.SERVICE, self.conn["user"],
                              self.conn["password"], self.conn["client"])
        results = []
        n = len(self.docs)
        for i, d in enumerate(self.docs, 1):
            r = await asyncio.to_thread(sc.post_one, client, m.ENTITY_SET, d, m.RESULT_KEYS)
            results.append(r)
            self.progress.value = i / n
            self.progress_text.value = f"{i} of {n}: {r['Reference']} {r['Status'].lower()}"
            self.page.update()

        self.posting = False
        self.results = results
        self.add_history(m, results)
        self.render()

    async def on_save_results(self, e):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        await self.save_bytes(f"results_{self.module.SLUG}_{stamp}.csv",
                              sc.results_csv(self.results or []))

    async def on_save_status(self, e):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        await self.save_bytes(f"{self.module.SLUG}_with_status_{stamp}.xlsx",
                              sc.upload_with_status_xlsx(self.table, self.results or []))

    # =====================================================================
    # History
    # =====================================================================
    def load_history(self) -> list[dict]:
        try:
            return json.loads(self.history_file.read_text(encoding="utf-8"))
        except Exception:
            return []

    def add_history(self, m, results):
        hist = self.load_history()
        hist.insert(0, {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M"), "upload": m.TITLE,
            "file": self.file_name, "user": self.conn["user"], "host": self.conn["host"],
            "success": sum(r["Status"] == "SUCCESS" for r in results),
            "failed": sum(r["Status"] in ("FAILED", "ERROR") for r in results),
            "invalid": sum(r["Status"] == "INVALID" for r in results),
            "results": results})
        try:
            self.history_file.write_text(json.dumps(hist[:200], indent=1), encoding="utf-8")
        except Exception as ex:
            self.toast(f"Couldn't save history: {ex}", error=True)

    def build_history(self):
        hist = self.load_history()
        if not hist:
            return self.width_box([card(ft.Column([
                heading("No posts yet"),
                ft.Text("Every post to DEV is listed here with the SAP document numbers, "
                        "so you can find them later.", size=14, color=MUTED)]))])
        tiles = []
        for h in hist:
            res_rows = [ft.Row([pill(r["Status"]),
                                ft.Text(r["Reference"], size=13, color=INK, expand=True),
                                ft.Text(r["SAP Document"] or r["Message"][:60], size=13,
                                        color=MUTED, selectable=True)], spacing=8)
                        for r in h["results"]]
            tiles.append(ft.ExpansionTile(
                title=ft.Text(f"{h['upload']}, {h['file']}", size=14,
                              weight=ft.FontWeight.W_600, color=INK),
                subtitle=ft.Text(f"{h['time']} by {h['user']}: {h['success']} created, "
                                 f"{h['failed']} failed, {h['invalid']} invalid",
                                 size=12, color=MUTED),
                controls=[ft.Container(ft.Column(res_rows, spacing=6),
                                       padding=ft.Padding.only(left=16, right=16, bottom=10))]))
        return self.width_box([
            card(ft.Column([
                ft.Row([heading("Posts to DEV"),
                        ft.TextButton("Clear history", icon=ft.Icons.DELETE_OUTLINE,
                                      on_click=self.on_clear_history)],
                       alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Column(tiles, spacing=0)], spacing=8))])

    def on_clear_history(self, e):
        def do_clear(ev):
            try:
                self.history_file.unlink(missing_ok=True)
            except Exception:
                pass
            self.page.pop_dialog()
            self.render()
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("Clear history?"),
            content=ft.Text("This removes the list of past posts from this device. "
                            "Documents already created in SAP are not affected."),
            actions=[ft.TextButton("Cancel", on_click=lambda ev: self.page.pop_dialog()),
                     ft.TextButton("Clear history", on_click=do_clear)]))

    # =====================================================================
    # Connection
    # =====================================================================
    def build_connection(self):
        c = self.conn
        self.f_host = ft.TextField(label="Host", value=c["host"],
                                   helper="DEV API host, e.g. my440263-api.s4hana.cloud.sap",
                                   keyboard_type=ft.KeyboardType.URL)
        self.f_client = ft.TextField(label="Client (optional)", value=c["client"],
                                     helper="Leave blank for S/4HANA Cloud",
                                     keyboard_type=ft.KeyboardType.NUMBER)
        self.f_user = ft.TextField(label="Communication user", value=c["user"])
        self.f_pw = ft.TextField(label="Password", value=c["password"], password=True,
                                 can_reveal_password=True)
        self.f_remember = ft.Checkbox(label="Remember password on this device",
                                      value=bool(c["remember"]))
        self.conn_status = ft.Text("", size=13)
        return self.width_box([card(ft.Column([
            heading("Connection"),
            ft.Text("The app posts straight from this device to your DEV tenant.",
                    size=14, color=MUTED),
            self.f_host, self.f_client, self.f_user, self.f_pw, self.f_remember,
            ft.Text("A remembered password is stored unencrypted in the app's data. "
                    "Leave this off on shared devices.", size=12, color=MUTED),
            ft.Row([ft.FilledButton("Save", icon=ft.Icons.CHECK, on_click=self.on_save_conn),
                    ft.OutlinedButton("Test connection", icon=ft.Icons.WIFI_TETHERING,
                                      on_click=self.on_test_conn)], spacing=8, wrap=True),
            self.conn_status,
        ], spacing=12))])

    def read_conn_fields(self):
        self.conn.update(host=(self.f_host.value or "").strip() or sc.DEFAULT_HOST,
                         client=(self.f_client.value or "").strip(),
                         user=(self.f_user.value or "").strip(),
                         password=self.f_pw.value or "",
                         remember=bool(self.f_remember.value))

    async def on_save_conn(self, e):
        self.read_conn_fields()
        c = self.conn
        try:
            for k in ("host", "client", "user"):
                await self.prefs.set(f"sapdev.{k}", c[k])
            await self.prefs.set("sapdev.remember", c["remember"])
            if c["remember"]:
                await self.prefs.set("sapdev.password", c["password"])
            else:
                await self.prefs.remove("sapdev.password")
        except Exception as ex:
            self.toast(f"Couldn't save settings: {ex}", error=True)
            return
        self.dev_bar.content = self.build_dev_bar()
        self.toast("Connection saved")
        self.page.update()

    async def on_test_conn(self, e):
        self.read_conn_fields()
        c = self.conn
        if not (c["user"] and c["password"]):
            self.conn_status.value, self.conn_status.color = \
                "Enter the user and password first.", STATUS["INVALID"][0]
            self.page.update()
            return
        self.conn_status.value, self.conn_status.color = "Contacting DEV...", MUTED
        self.page.update()
        client = sc.SAPClient(c["host"], "API_PRODUCT_SRV", c["user"], c["password"],
                              c["client"], timeout=30)
        ok, msg = await asyncio.to_thread(client.test)
        self.conn_status.value = f"{msg} to {c['host']}" if ok else msg
        self.conn_status.color = STATUS["SUCCESS"][0] if ok else STATUS["FAILED"][0]
        self.dev_bar.content = self.build_dev_bar()
        self.page.update()


async def main(page: ft.Page):
    app = App(page)
    await app.start()


if __name__ == "__main__":
    ft.run(main)
