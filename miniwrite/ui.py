"""MiniWrite GUI — mini Word-like editor."""

from __future__ import annotations

import ctypes
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox

from . import __version__
from . import document as doc
from .instance import WINDOW_TITLE

BG = "#E6E6E6"          # серое поле вокруг листа
CHROME = "#FFFFFF"      # лента
CHROME2 = "#F3F3F3"
FG = "#242424"
MUTED = "#605E5C"
LINE = "#C8C6C4"
DOC_BG = "#FFFFFF"
DOC_FG = "#000000"
WORD_BLUE = "#2B579A"
WORD_BLUE_H = "#1E3F6F"
ACTIVE = "#C5D9F1"
BTN = "#FFFFFF"
BTN_HOVER = "#E1DFDD"

SIZES = (10, 11, 12, 14, 16, 18, 20, 24, 28, 36)
FAMILIES = ("Calibri", "Segoe UI", "Times New Roman", "Arial", "Consolas")
STYLE_TAGS = ("font_br", "font_bi", "font_nr", "font_ni", "underline", "left", "center", "right")


class MiniWriteApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("960x680")
        self.minsize(720, 480)
        self.configure(bg=BG)

        self._path: Path | None = None
        self._dirty = False
        self._family = tk.StringVar(value="Calibri")
        self._size = tk.IntVar(value=11)
        self._fmt_btns: dict[str, tk.Label] = {}
        self._type_bold = False
        self._type_italic = False
        self._type_underline = False

        self._build()
        self._bind_keys()
        self.after(20, self._place)
        self.after(100, self._body.focus_set)
        self._set_dirty(False)
        self._update_counts()

    def _build(self) -> None:
        # синяя шапка как у Word
        titlebar = tk.Frame(self, bg=WORD_BLUE, height=36)
        titlebar.pack(fill="x")
        titlebar.pack_propagate(False)
        tk.Label(
            titlebar,
            text="W  MiniWrite",
            bg=WORD_BLUE,
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 11),
            padx=12,
        ).pack(side="left", pady=8)
        self._title_doc = tk.Label(
            titlebar,
            text="Документ 1",
            bg=WORD_BLUE,
            fg="#D6E4F5",
            font=("Segoe UI", 10),
        )
        self._title_doc.pack(side="left", padx=(8, 0))
        tk.Label(
            titlebar,
            text="день 42 · Текст",
            bg=WORD_BLUE,
            fg="#A9C4E0",
            font=("Segoe UI", 9),
            padx=12,
        ).pack(side="right")

        # вкладка «Главная»
        tabs = tk.Frame(self, bg=CHROME2)
        tabs.pack(fill="x")
        tab_home = tk.Label(
            tabs,
            text="Главная",
            bg=CHROME,
            fg=WORD_BLUE,
            font=("Segoe UI Semibold", 10),
            padx=16,
            pady=8,
        )
        tab_home.pack(side="left")
        tk.Frame(tabs, bg=WORD_BLUE, height=2).place(x=0, rely=1.0, relwidth=0.12, anchor="sw")

        # лента
        ribbon = tk.Frame(self, bg=CHROME, highlightbackground=LINE, highlightthickness=1)
        ribbon.pack(fill="x")

        # группа Файл
        g_file = self._ribbon_group(ribbon, "Файл")
        for text, cmd in (
            ("Создать", self.file_new),
            ("Открыть", self.file_open),
            ("Сохранить", self.file_save),
            ("Как…", self.file_save_as),
        ):
            self._tool_btn(g_file, text, cmd).pack(side="left", padx=3, pady=(6, 2))

        self._ribbon_sep(ribbon)

        # группа Шрифт
        g_font = self._ribbon_group(ribbon, "Шрифт")
        fam = tk.OptionMenu(g_font, self._family, *FAMILIES, command=lambda _v: self.apply_font())
        fam.config(
            bg=BTN,
            fg=FG,
            activebackground=BTN_HOVER,
            highlightthickness=1,
            highlightbackground=LINE,
            bd=0,
            width=14,
        )
        fam["menu"].config(bg=BTN, fg=FG)
        fam.pack(side="left", padx=(4, 4), pady=(8, 2))

        size = tk.OptionMenu(g_font, self._size, *SIZES, command=lambda _v: self.apply_font())
        size.config(
            bg=BTN,
            fg=FG,
            activebackground=BTN_HOVER,
            highlightthickness=1,
            highlightbackground=LINE,
            bd=0,
            width=3,
        )
        size["menu"].config(bg=BTN, fg=FG)
        size.pack(side="left", padx=(0, 8), pady=(8, 2))

        for key, label in (("bold", "Ж"), ("italic", "К"), ("underline", "Ч")):
            btn = self._tool_btn(g_font, label, lambda k=key: self.toggle_style(k), square=True)
            btn.pack(side="left", padx=2, pady=(8, 2))
            self._fmt_btns[key] = btn

        self._ribbon_sep(ribbon)

        # группа Абзац
        g_par = self._ribbon_group(ribbon, "Абзац")
        for how, label in (("left", "≡ слева"), ("center", "≡ центр"), ("right", "≡ справа")):
            self._tool_btn(g_par, label, lambda a=how: self.align(a)).pack(
                side="left", padx=2, pady=(8, 2)
            )
        self._tool_btn(g_par, "• Список", self.bullet).pack(side="left", padx=(8, 4), pady=(8, 2))

        # рабочая область — серый стол, белый лист
        workspace = tk.Frame(self, bg=BG)
        workspace.pack(fill="both", expand=True)

        # прокрутка, если окно меньше листа
        canvas = tk.Canvas(workspace, bg=BG, highlightthickness=0, bd=0)
        vscroll = tk.Scrollbar(workspace, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        page_host = tk.Frame(canvas, bg=BG)
        canvas_window = canvas.create_window((0, 0), window=page_host, anchor="n")

        def _center_page(_event=None) -> None:
            canvas.update_idletasks()
            cw = canvas.winfo_width()
            page_host.update_idletasks()
            pw = page_host.winfo_reqwidth()
            x = max((cw - pw) // 2, 8)
            canvas.coords(canvas_window, x, 12)
            canvas.configure(scrollregion=canvas.bbox("all"))

        page_host.bind("<Configure>", _center_page)
        canvas.bind("<Configure>", _center_page)

        shadow = tk.Frame(page_host, bg="#C0C0C0")
        shadow.pack(padx=8, pady=8)
        page = tk.Frame(shadow, bg=DOC_BG)
        page.pack(padx=(0, 2), pady=(0, 2))

        self._body = tk.Text(
            page,
            bg=DOC_BG,
            fg=DOC_FG,
            insertbackground=DOC_FG,
            relief="flat",
            font=("Calibri", 11),
            wrap="word",
            width=84,
            height=32,
            padx=56,
            pady=48,
            undo=True,
            maxundo=0,
            spacing1=3,
            spacing3=6,
            highlightthickness=0,
            bd=0,
        )
        self._body.pack()

        self._fonts: dict[str, tkfont.Font] = {}
        self._rebuild_fonts()
        self._body.tag_configure("underline", underline=1)
        self._body.tag_configure("left", justify="left")
        self._body.tag_configure("center", justify="center")
        self._body.tag_configure("right", justify="right")

        self._body.bind("<<Modified>>", self._on_modified)
        self._body.bind("<Key>", self._on_key)
        self._body.bind("<KeyRelease>", lambda _e: self._sync_format_buttons())
        self._body.bind("<ButtonRelease-1>", lambda _e: self._sync_format_buttons())

        # статусная строка Word
        status = tk.Frame(self, bg=CHROME2, highlightbackground=LINE, highlightthickness=1)
        status.pack(fill="x", side="bottom")
        self._status = tk.Label(
            status,
            text="",
            bg=CHROME2,
            fg=FG,
            font=("Segoe UI", 9),
            anchor="w",
            padx=12,
            pady=4,
        )
        self._status.pack(side="left")
        tk.Label(
            status,
            text=f"v{__version__}  ·  darkshade",
            bg=CHROME2,
            fg=MUTED,
            font=("Segoe UI", 9),
            padx=12,
        ).pack(side="right")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _ribbon_group(self, parent: tk.Frame, title: str) -> tk.Frame:
        box = tk.Frame(parent, bg=CHROME)
        box.pack(side="left", fill="y", padx=6, pady=4)
        inner = tk.Frame(box, bg=CHROME)
        inner.pack(side="top", fill="x")
        tk.Label(
            box,
            text=title,
            bg=CHROME,
            fg=MUTED,
            font=("Segoe UI", 8),
        ).pack(side="bottom", pady=(0, 2))
        return inner

    def _ribbon_sep(self, parent: tk.Frame) -> None:
        tk.Frame(parent, bg=LINE, width=1).pack(side="left", fill="y", padx=6, pady=10)

    def _tool_btn(self, parent, text, cmd, square: bool = False) -> tk.Label:
        lbl = tk.Label(
            parent,
            text=text,
            bg=BTN,
            fg=FG,
            font=("Segoe UI Semibold" if square else "Segoe UI", 10 if square else 9),
            padx=10 if not square else 8,
            pady=5,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=LINE,
        )
        lbl.bind("<Button-1>", lambda _e: cmd())
        lbl.bind("<Enter>", lambda _e: lbl.configure(bg=BTN_HOVER))
        lbl.bind("<Leave>", lambda _e: self._btn_leave(lbl))
        return lbl

    def _btn_leave(self, lbl: tk.Label) -> None:
        for btn in self._fmt_btns.values():
            if btn is lbl:
                # leave active color if selected — resync
                self._sync_format_buttons()
                return
        lbl.configure(bg=BTN)

    def _font_key(self, bold: bool, italic: bool) -> str:
        return f"{'b' if bold else 'n'}{'i' if italic else 'r'}"

    def _rebuild_fonts(self) -> None:
        family = self._family.get()
        size = int(self._size.get())
        for bold in (False, True):
            for italic in (False, True):
                key = self._font_key(bold, italic)
                name = f"mw_{key}_{family.replace(' ', '_')}_{size}"
                try:
                    f = tkfont.Font(self, name=name, exists=True)
                except tk.TclError:
                    f = tkfont.Font(self, name=name, exists=False)
                f.configure(
                    family=family,
                    size=size,
                    weight="bold" if bold else "normal",
                    slant="italic" if italic else "roman",
                )
                self._fonts[key] = f
                self._body.tag_configure(f"font_{key}", font=f)
        self._body.configure(font=self._fonts[self._font_key(False, False)])

    def _place(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width() or 960, self.winfo_height() or 680
        sw = ctypes.windll.user32.GetSystemMetrics(0)
        sh = ctypes.windll.user32.GetSystemMetrics(1)
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 5}")

    def _bind_keys(self) -> None:
        self.bind("<Control-n>", lambda e: (self.file_new(), "break")[1])
        self.bind("<Control-o>", lambda e: (self.file_open(), "break")[1])
        self.bind("<Control-s>", lambda e: (self.file_save(), "break")[1])
        self.bind("<Control-b>", lambda e: (self.toggle_style("bold"), "break")[1])
        self.bind("<Control-i>", lambda e: (self.toggle_style("italic"), "break")[1])
        self.bind("<Control-u>", lambda e: (self.toggle_style("underline"), "break")[1])

    def _on_modified(self, _event=None) -> None:
        if self._body.edit_modified():
            self._set_dirty(True)
            self._body.edit_modified(False)
            self._update_counts()

    def _on_key(self, event: tk.Event):
        if event.keysym in (
            "Shift_L",
            "Shift_R",
            "Control_L",
            "Control_R",
            "Alt_L",
            "Alt_R",
            "Caps_Lock",
            "Left",
            "Right",
            "Up",
            "Down",
            "Home",
            "End",
            "Next",
            "Prior",
            "BackSpace",
            "Delete",
            "Return",
            "Tab",
            "Escape",
        ):
            return
        if event.state & 0x4:  # Ctrl shortcuts
            return
        ch = event.char
        if not ch or ch < " ":
            return
        # insert styled char ourselves when typing style active
        if not (self._type_bold or self._type_italic or self._type_underline):
            return
        try:
            self._body.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        tag = f"font_{self._font_key(self._type_bold, self._type_italic)}"
        self._body.insert("insert", ch, (tag, "underline") if self._type_underline else (tag,))
        self._set_dirty(True)
        return "break"

    def _set_dirty(self, dirty: bool) -> None:
        self._dirty = dirty
        name = self._path.name if self._path else "Документ 1"
        self.title(f"MiniWrite - {name}{' *' if dirty else ''}")
        if hasattr(self, "_title_doc"):
            self._title_doc.configure(text=f"{name}{' *' if dirty else ''}")

    def _update_counts(self) -> None:
        text = self._body.get("1.0", "end-1c")
        words = len(text.split()) if text.strip() else 0
        lines = text.count("\n") + (1 if text else 0)
        self._status.configure(
            text=f"Стр. 1 из 1    │    {words} слов    │    {len(text)} знаков    │    {lines} строк"
        )

    def _sel(self) -> tuple[str, str] | None:
        try:
            return self._body.index("sel.first"), self._body.index("sel.last")
        except tk.TclError:
            return None

    def _sync_format_buttons(self) -> None:
        sel = self._sel()
        if sel:
            tags = set(self._body.tag_names(sel[0]))
            bold = "font_bi" in tags or "font_br" in tags
            italic = "font_bi" in tags or "font_ni" in tags
            under = "underline" in tags
        else:
            bold, italic, under = self._type_bold, self._type_italic, self._type_underline
        for key, on in (("bold", bold), ("italic", italic), ("underline", under)):
            btn = self._fmt_btns[key]
            if on:
                btn.configure(bg=ACTIVE, highlightbackground=WORD_BLUE)
            else:
                btn.configure(bg=BTN, highlightbackground=LINE)

    def apply_font(self) -> None:
        self._rebuild_fonts()
        sel = self._sel()
        if not sel:
            return
        a, b = sel
        tags = set(self._body.tag_names(a))
        bold = "font_bi" in tags or "font_br" in tags
        italic = "font_bi" in tags or "font_ni" in tags
        for t in ("font_br", "font_bi", "font_nr", "font_ni"):
            self._body.tag_remove(t, a, b)
        self._body.tag_add(f"font_{self._font_key(bold, italic)}", a, b)
        self._set_dirty(True)

    def toggle_style(self, kind: str) -> None:
        sel = self._sel()
        if kind == "underline":
            if sel:
                a, b = sel
                if "underline" in self._body.tag_names(a):
                    self._body.tag_remove("underline", a, b)
                else:
                    self._body.tag_add("underline", a, b)
            else:
                self._type_underline = not self._type_underline
            self._sync_format_buttons()
            self._set_dirty(True)
            return

        if sel:
            a, b = sel
            tags = set(self._body.tag_names(a))
            bold = "font_bi" in tags or "font_br" in tags
            italic = "font_bi" in tags or "font_ni" in tags
            if kind == "bold":
                bold = not bold
            else:
                italic = not italic
            for t in ("font_br", "font_bi", "font_nr", "font_ni"):
                self._body.tag_remove(t, a, b)
            self._body.tag_add(f"font_{self._font_key(bold, italic)}", a, b)
        else:
            if kind == "bold":
                self._type_bold = not self._type_bold
            else:
                self._type_italic = not self._type_italic
        self._sync_format_buttons()
        self._set_dirty(True)

    def align(self, how: str) -> None:
        sel = self._sel()
        if sel:
            a, b = sel
            start = self._body.index(f"{a} linestart")
            end = self._body.index(f"{b} lineend")
        else:
            start = self._body.index("insert linestart")
            end = self._body.index("insert lineend")
        for t in ("left", "center", "right"):
            self._body.tag_remove(t, start, end)
        self._body.tag_add(how, start, end)
        self._set_dirty(True)

    def bullet(self) -> None:
        self._body.insert("insert linestart", "• ")
        self._set_dirty(True)

    def _export_tags(self) -> list[dict]:
        out: list[dict] = []
        for tag in STYLE_TAGS:
            ranges = self._body.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                out.append(
                    {
                        "tag": tag,
                        "start": str(ranges[i]),
                        "end": str(ranges[i + 1]),
                    }
                )
        return out

    def _import_tags(self, tags: list[dict]) -> None:
        for item in tags:
            tag = item.get("tag")
            start = item.get("start")
            end = item.get("end")
            if tag in STYLE_TAGS and start and end:
                try:
                    self._body.tag_add(tag, start, end)
                except tk.TclError:
                    pass

    def _confirm_discard(self) -> bool:
        if not self._dirty:
            return True
        ans = messagebox.askyesnocancel("MiniWrite", "Сохранить изменения?", parent=self)
        if ans is None:
            return False
        if ans:
            return self.file_save()
        return True

    def file_new(self) -> None:
        if not self._confirm_discard():
            return
        self._body.delete("1.0", "end")
        self._path = None
        self._type_bold = self._type_italic = self._type_underline = False
        self._body.edit_reset()
        self._set_dirty(False)
        self._update_counts()
        self._sync_format_buttons()

    def file_open(self) -> None:
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("MiniWrite", "*.mw"), ("Текст", "*.txt"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        p = Path(path)
        try:
            self._body.delete("1.0", "end")
            if p.suffix.lower() == ".mw":
                text, tags = doc.load_mw(p)
                self._body.insert("1.0", text)
                self._import_tags(tags)
            else:
                self._body.insert("1.0", doc.load_txt(p))
            self._path = p
            self._body.edit_modified(False)
            self._set_dirty(False)
            self._update_counts()
        except Exception as exc:
            messagebox.showerror("MiniWrite", f"Не открыть файл:\n{exc}", parent=self)

    def file_save(self) -> bool:
        if self._path is None:
            return self.file_save_as()
        return self._write(self._path)

    def file_save_as(self) -> bool:
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".mw",
            filetypes=[("MiniWrite", "*.mw"), ("Текст", "*.txt")],
        )
        if not path:
            return False
        return self._write(Path(path))

    def _write(self, path: Path) -> bool:
        try:
            plain = self._body.get("1.0", "end-1c")
            if path.suffix.lower() == ".txt":
                doc.save_txt(path, plain)
            else:
                if path.suffix.lower() != ".mw":
                    path = path.with_suffix(".mw")
                doc.save_mw(path, plain, self._export_tags())
            self._path = path
            self._set_dirty(False)
            self._update_counts()
            return True
        except Exception as exc:
            messagebox.showerror("MiniWrite", f"Не сохранить:\n{exc}", parent=self)
            return False

    def on_close(self) -> None:
        if self._confirm_discard():
            self.destroy()


def run() -> None:
    MiniWriteApp().mainloop()
