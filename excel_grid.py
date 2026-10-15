# -*- coding: utf-8 -*-
"""
ExcelGrid — grilla editable estilo Excel para Tkinter/CustomTkinter.

Widget genérico y reutilizable: no sabe nada de AGP Glass, de mallas.* ni de
ninguna tabla en particular. Solo pinta filas/columnas en un Canvas y avisa
al que lo usa cuándo el usuario cambió algo, vía callbacks. Así se puede
reusar en cualquier otra pestaña futura que necesite una tabla editable.

Qué trae:
  - Seleccionar una celda o arrastrar para seleccionar un rango (como Excel).
  - Ctrl+C copia el rango seleccionado (tab-separated, pegable en Excel real).
  - Ctrl+V pega un bloque tab-separated a partir de la celda activa.
  - Tab / Shift+Tab / Enter / flechas para moverse sin soltar el teclado.
  - Doble clic o F2 para editar una celda de texto; clic simple en una
    columna tipo "bool" prende/apaga el check al toque.
  - Suprimir / Backspace borra el contenido de las celdas seleccionadas.
  - Solo dibuja las filas visibles (virtualizado) — no importa si la tabla
    tiene 50 filas o 50.000, el scroll siempre se siente liviano.

Cómo se usa (ver tab_logos_laser.py para un ejemplo real):

    columnas = [
        {"key": "id",      "header": "ID",      "width": 90,  "type": "text", "editable": False},
        {"key": "cliente", "header": "Cliente",  "width": 160, "type": "text"},
        {"key": "listo",   "header": "Listo",    "width": 80,  "type": "bool", "align": "center"},
    ]

    def on_edit(row, key, valor_nuevo):
        # guardar en BD; devolver True si quedó bien, False si no
        return guardar_en_bd(row["id"], key, valor_nuevo)

    def on_new_row():
        # crear el registro en BD (con su ID autogenerado) y devolver la fila
        return crear_en_bd()

    grid = ExcelGrid(parent, columnas, on_edit=on_edit, on_new_row=on_new_row)
    grid.pack(fill="both", expand=True)
    grid.set_rows(filas_desde_bd)
"""

import tkinter as tk
import tkinter.font as tkfont

try:
    import customtkinter as ctk
    _BASE = ctk.CTkFrame
except ImportError:
    _BASE = tk.Frame

# Paleta neutra por defecto — si el que llama pasa `theme=` propio, se usa esa.
_TEMA_DEFAULT = {
    "bg":          "#0F1117",
    "header_bg":   "#1C2333",
    "header_fg":   "#94A3B8",
    "row_bg":      "#161B27",
    "row_bg_alt":  "#1A2032",
    "border":      "#2E3A55",
    "text":        "#F1F5F9",
    "text_dim":    "#475569",
    "accent":      "#3B82F6",
    "select_bg":   "#1E3A66",
    "input_bg":    "#1C2333",
    "red":         "#EF4444",
    "green":       "#10B981",
}

# Valores de texto que al pegar se interpretan como "casilla marcada"
_VERDADEROS = {"1", "true", "si", "sí", "x", "yes", "on", "✓"}


class ExcelGrid(_BASE):
    def __init__(self, master, columns, on_edit=None, on_new_row=None,
                 status_cb=None, theme=None, row_height=32, header_height=34,
                 **kw):
        kw.setdefault("fg_color", "transparent") if _BASE is not tk.Frame else None
        super().__init__(master, **kw)
        self._columns   = columns
        self._on_edit   = on_edit
        self._on_new_row = on_new_row
        self._status_cb = status_cb
        self._theme     = {**_TEMA_DEFAULT, **(theme or {})}
        self.row_height    = row_height
        self.header_height = header_height

        self._rows      = []          # lista de dicts (una por fila)
        self._anchor    = (0, 0)      # esquina de inicio de la selección
        self._active    = (0, 0)      # celda activa (donde termina la selección / se edita)
        self._editing   = None        # (r, c) si hay una celda en edición
        self._stretch_idx = next(
            (i for i, c in enumerate(columns) if c.get("stretch")), None)
        self._base_widths = [c["width"] for c in columns]

        self._cell_font   = tkfont.Font(family="Segoe UI", size=10)
        self._header_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")

        self._recompute_layout()
        self._build_ui()
        self._draw_header()
        self._redraw()

    # ── construcción de widgets ─────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._canvas_header = tk.Canvas(
            self, height=self.header_height, highlightthickness=0,
            bg=self._theme["header_bg"])
        self._canvas_header.grid(row=0, column=0, sticky="ew")

        body_wrap = tk.Frame(self, bg=self._theme["bg"])
        body_wrap.grid(row=1, column=0, sticky="nsew")
        body_wrap.grid_columnconfigure(0, weight=1)
        body_wrap.grid_rowconfigure(0, weight=1)

        self._canvas_body = tk.Canvas(
            body_wrap, highlightthickness=0, bg=self._theme["row_bg"],
            yscrollincrement=self.row_height)
        self._canvas_body.grid(row=0, column=0, sticky="nsew")

        self._vbar = tk.Scrollbar(body_wrap, orient="vertical",
                                   command=self._canvas_body.yview)
        self._vbar.grid(row=0, column=1, sticky="ns")
        self._hbar = tk.Scrollbar(self, orient="horizontal",
                                   command=self._on_hbar)
        self._hbar.grid(row=2, column=0, sticky="ew")

        self._canvas_body.configure(yscrollcommand=self._on_body_yscroll,
                                     xscrollcommand=self._on_body_xscroll)

        # Entry flotante único, reutilizado para editar cualquier celda
        self._entry = tk.Entry(
            self._canvas_body, relief="flat", bd=0,
            bg=self._theme["input_bg"], fg=self._theme["text"],
            insertbackground=self._theme["text"], font=self._cell_font)
        self._entry_win = self._canvas_body.create_window(
            0, 0, anchor="nw", window=self._entry, width=10, height=10,
            state="hidden")
        self._entry.bind("<Return>",     self._entry_return)
        self._entry.bind("<KP_Enter>",   self._entry_return)
        self._entry.bind("<Tab>",        self._entry_tab)
        self._entry.bind("<Shift-Tab>",  self._entry_shift_tab)
        self._entry.bind("<Escape>",     lambda e: (self._cancel_edit(), "break"))
        self._entry.bind("<FocusOut>",   lambda e: self._commit_edit())

        cv = self._canvas_body
        cv.bind("<Button-1>",         self._on_click)
        cv.bind("<B1-Motion>",        self._on_drag)
        cv.bind("<Double-Button-1>",  self._on_double_click)
        cv.bind("<MouseWheel>",       self._on_wheel)
        cv.bind("<Configure>",        self._on_body_configure)
        cv.bind("<KeyPress>",         self._on_key)
        cv.bind("<Up>",               lambda e: self._move_active(-1, 0, self._shift(e)))
        cv.bind("<Down>",             lambda e: self._move_active(1, 0, self._shift(e)))
        cv.bind("<Left>",             lambda e: self._move_active(0, -1, self._shift(e)))
        cv.bind("<Right>",            lambda e: self._move_active(0, 1, self._shift(e)))
        cv.bind("<Tab>",              self._on_tab)
        cv.bind("<Shift-Tab>",        self._on_shift_tab)
        cv.bind("<Return>",           self._on_return_key)
        cv.bind("<F2>",               self._on_f2)
        cv.bind("<Delete>",           self._on_delete_key)
        cv.bind("<BackSpace>",        self._on_delete_key)
        cv.bind("<Control-c>",        self._on_copy)
        cv.bind("<Control-C>",        self._on_copy)
        cv.bind("<Control-v>",        self._on_paste)
        cv.bind("<Control-V>",        self._on_paste)
        cv.bind("<Button-1>",         lambda e: cv.focus_set(), add=True)

    @staticmethod
    def _shift(event):
        return bool(event.state & 0x0001)

    # ── layout de columnas (con estiramiento de la última si sobra ancho) ──
    def _recompute_layout(self, avail_width=None):
        widths = list(self._base_widths)
        total = sum(widths)
        if avail_width and avail_width > total and self._stretch_idx is not None:
            widths[self._stretch_idx] += (avail_width - total)
        self._col_widths = widths
        lefts, acc = [], 0
        for w in widths:
            lefts.append(acc)
            acc += w
        self._col_lefts   = lefts
        self._total_width = acc

    def _update_scrollregions(self):
        h = max(1, len(self._rows) * self.row_height)
        self._canvas_body.configure(scrollregion=(0, 0, self._total_width, h))
        self._canvas_header.configure(scrollregion=(0, 0, self._total_width, self.header_height))

    def _on_body_configure(self, event):
        self._recompute_layout(avail_width=event.width)
        self._update_scrollregions()
        self._draw_header()
        self._redraw()

    # ── scroll ───────────────────────────────────────────────────────────
    def _on_hbar(self, *args):
        self._canvas_body.xview(*args)

    def _on_body_yscroll(self, lo, hi):
        self._vbar.set(lo, hi)
        self._redraw()

    def _on_body_xscroll(self, lo, hi):
        self._hbar.set(lo, hi)
        self._canvas_header.xview_moveto(lo)
        self._redraw()

    def _on_wheel(self, event):
        self._canvas_body.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def focus_grid(self):
        self._canvas_body.focus_set()

    # ── datos públicos ───────────────────────────────────────────────────
    def set_rows(self, rows):
        """Reemplaza todos los datos (ej. al refrescar desde BD)."""
        self._commit_edit()
        self._rows = list(rows)
        n_rows = len(self._rows)
        first_editable = next(
            (i for i, c in enumerate(self._columns) if c.get("editable", True)), 0)
        self._anchor = self._active = (0, first_editable) if n_rows else (0, 0)
        self._update_scrollregions()
        self._canvas_body.yview_moveto(0)
        self._redraw()

    def append_row(self, row):
        """Agrega una fila nueva (ej. tras crearla en BD) y salta a editarla."""
        self._rows.append(row)
        self._update_scrollregions()
        r = len(self._rows) - 1
        first_editable = next(
            (i for i, c in enumerate(self._columns) if c.get("editable", True)), 0)
        self._anchor = self._active = (r, first_editable)
        self._ensure_visible(r, first_editable)
        self._redraw()
        self.focus_grid()

    def get_rows(self):
        return list(self._rows)

    # ── helpers de estado ────────────────────────────────────────────────
    def _toast(self, msg, level="ok"):
        if self._status_cb:
            try:
                self._status_cb(msg, level)
            except Exception:
                pass

    def _visible_row_range(self):
        if not self._rows:
            return 0, -1
        total_h = max(1, len(self._rows) * self.row_height)
        top_frac, bot_frac = self._canvas_body.yview()
        first = max(0, int((top_frac * total_h) // self.row_height) - 2)
        last  = min(len(self._rows) - 1, int((bot_frac * total_h) // self.row_height) + 2)
        return first, last

    def _sel_bounds(self):
        (ar, ac), (br, bc) = self._anchor, self._active
        return min(ar, br), max(ar, br), min(ac, bc), max(ac, bc)

    def _in_selection(self, r, c):
        r0, r1, c0, c1 = self._sel_bounds()
        return r0 <= r <= r1 and c0 <= c <= c1

    def _iter_selected_cells(self):
        r0, r1, c0, c1 = self._sel_bounds()
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                yield r, c

    def _xy_to_cell(self, event):
        cx = self._canvas_body.canvasx(event.x)
        cy = self._canvas_body.canvasy(event.y)
        r = int(cy // self.row_height)
        if r < 0 or r >= len(self._rows):
            return None
        for c, left in enumerate(self._col_lefts):
            if left <= cx < left + self._col_widths[c]:
                return r, c
        return None

    def _ensure_visible(self, r, c):
        total_h = max(1, len(self._rows) * self.row_height)
        y0, y1 = r * self.row_height, (r + 1) * self.row_height
        top_f, bot_f = self._canvas_body.yview()
        top_px, bot_px = top_f * total_h, bot_f * total_h
        if y0 < top_px:
            self._canvas_body.yview_moveto(y0 / total_h)
        elif y1 > bot_px:
            self._canvas_body.yview_moveto(max(0, y1 - (bot_px - top_px)) / total_h)

        x0, x1 = self._col_lefts[c], self._col_lefts[c] + self._col_widths[c]
        left_f, right_f = self._canvas_body.xview()
        left_px, right_px = left_f * self._total_width, right_f * self._total_width
        if x0 < left_px:
            self._canvas_body.xview_moveto(x0 / self._total_width)
        elif x1 > right_px:
            self._canvas_body.xview_moveto(max(0, x1 - (right_px - left_px)) / self._total_width)

    # ── dibujo ───────────────────────────────────────────────────────────
    def _draw_header(self):
        cv = self._canvas_header
        cv.delete("all")
        t = self._theme
        for c, col in enumerate(self._columns):
            x, w = self._col_lefts[c], self._col_widths[c]
            cv.create_rectangle(x, 0, x + w, self.header_height,
                                 fill=t["header_bg"], outline=t["border"])
            centrado = col.get("align") == "center"
            cv.create_text(x + w / 2 if centrado else x + 8, self.header_height / 2,
                            text=col["header"], anchor="center" if centrado else "w",
                            fill=t["header_fg"], font=self._header_font)

    def _truncate(self, text, max_px):
        if max_px <= 0:
            return ""
        if self._cell_font.measure(text) <= max_px:
            return text
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self._cell_font.measure(text[:mid] + "…") <= max_px:
                lo = mid
            else:
                hi = mid - 1
        return (text[:lo] + "…") if lo > 0 else "…"

    def _draw_checkbox(self, cv, x, y, w, h, checked):
        t = self._theme
        size = 16
        cx, cy = x + w / 2, y + h / 2
        x0, y0, x1, y1 = cx - size / 2, cy - size / 2, cx + size / 2, cy + size / 2
        cv.create_rectangle(x0, y0, x1, y1, outline=t["border"], width=1.5,
                             fill=t["accent"] if checked else t["input_bg"], tags="cell")
        if checked:
            cv.create_line(x0 + 3, cy, cx - 1, y1 - 3, fill="#ffffff", width=2, tags="cell")
            cv.create_line(cx - 1, y1 - 3, x1 - 3, y0 + 3, fill="#ffffff", width=2, tags="cell")

    def _redraw(self):
        cv, t = self._canvas_body, self._theme
        cv.delete("cell")

        if not self._rows:
            cv.create_text(20, 20, anchor="nw", fill=t["text_dim"],
                            font=self._cell_font, tags="cell",
                            text="Sin registros todavía — usa “➕ Nueva fila” para crear el primero.")
            return

        first, last = self._visible_row_range()
        for r in range(first, last + 1):
            row = self._rows[r]
            y0, y1 = r * self.row_height, (r + 1) * self.row_height
            zebra = t["row_bg_alt"] if r % 2 else t["row_bg"]
            for c, col in enumerate(self._columns):
                x, w = self._col_lefts[c], self._col_widths[c]
                bg = t["select_bg"] if self._in_selection(r, c) else zebra
                cv.create_rectangle(x, y0, x + w, y1, fill=bg, outline=t["border"],
                                     width=1, tags="cell")
                if col["type"] == "bool":
                    self._draw_checkbox(cv, x, y0, w, self.row_height, bool(row.get(col["key"])))
                else:
                    val = row.get(col["key"])
                    txt = self._truncate("" if val is None else str(val), w - 16)
                    color = col.get("color") or t["text"]
                    centrado = col.get("align") == "center"
                    cv.create_text(x + w / 2 if centrado else x + 8, (y0 + y1) / 2,
                                    text=txt, anchor="center" if centrado else "w",
                                    fill=color, font=self._cell_font, tags="cell")
        self._draw_active_border()

    def _draw_active_border(self):
        if not self._rows:
            return
        r, c = self._active
        if r >= len(self._rows):
            return
        x, w = self._col_lefts[c], self._col_widths[c]
        y0 = r * self.row_height
        self._canvas_body.create_rectangle(
            x + 1, y0 + 1, x + w - 1, y0 + self.row_height - 1,
            outline=self._theme["accent"], width=2, tags="cell")

    # ── selección / navegación con mouse ─────────────────────────────────
    def _on_click(self, event):
        cell = self._xy_to_cell(event)
        if cell is None:
            return
        r, c = cell
        if self._editing and (r, c) != self._editing:
            self._commit_edit()
        col = self._columns[c]
        if col["type"] == "bool" and col.get("editable", True):
            self._toggle_bool(r, c)
        self._anchor = self._active = (r, c)
        self._redraw()

    def _on_drag(self, event):
        cell = self._xy_to_cell(event)
        if cell is None:
            return
        self._active = cell
        self._redraw()

    def _on_double_click(self, event):
        cell = self._xy_to_cell(event)
        if cell is None:
            return
        r, c = cell
        col = self._columns[c]
        if col["type"] == "text" and col.get("editable", True):
            self._start_edit(r, c, select_all=True)

    # ── navegación con teclado ────────────────────────────────────────────
    _IGNORAR = {"Tab", "Return", "KP_Enter", "Up", "Down", "Left", "Right",
                "Escape", "Delete", "BackSpace", "F2",
                "Control_L", "Control_R", "Shift_L", "Shift_R", "Alt_L", "Alt_R"}

    def _on_key(self, event):
        if event.keysym in self._IGNORAR:
            return
        if not event.char or not event.char.isprintable():
            return
        r, c = self._active
        if r >= len(self._rows):
            return
        col = self._columns[c]
        if col["type"] == "text" and col.get("editable", True):
            self._start_edit(r, c, initial_text=event.char)

    def _move_active(self, dr, dc, extend=False):
        if not self._rows:
            return "break"
        self._commit_edit()
        r, c = self._active
        nr = min(max(0, r + dr), len(self._rows) - 1)
        nc = min(max(0, c + dc), len(self._columns) - 1)
        self._active = (nr, nc)
        if not extend:
            self._anchor = (nr, nc)
        self._ensure_visible(nr, nc)
        self._redraw()
        self._canvas_body.focus_set()
        return "break"

    def _on_tab(self, event):
        self._advance(1)
        return "break"

    def _on_shift_tab(self, event):
        self._advance(-1)
        return "break"

    def _advance(self, step):
        self._commit_edit()
        if not self._rows:
            return
        r, c = self._active
        c += step
        if c >= len(self._columns):
            c = 0
            r = min(r + 1, len(self._rows) - 1)
        elif c < 0:
            c = len(self._columns) - 1
            r = max(r - 1, 0)
        self._anchor = self._active = (r, c)
        self._ensure_visible(r, c)
        self._redraw()
        self._canvas_body.focus_set()

    def _on_return_key(self, event):
        self._move_active(1, 0)
        return "break"

    def _on_f2(self, event):
        r, c = self._active
        col = self._columns[c]
        if col["type"] == "text" and col.get("editable", True):
            self._start_edit(r, c, cursor_at_end=True)
        return "break"

    def _on_delete_key(self, event):
        if self._editing:
            return  # el Entry maneja su propio borrado
        for r, c in self._iter_selected_cells():
            col = self._columns[c]
            if not col.get("editable", True):
                continue
            nuevo = False if col["type"] == "bool" else ""
            self._commit_cell(r, c, nuevo)
        self._redraw()
        return "break"

    # ── edición de celdas de texto ───────────────────────────────────────
    def _start_edit(self, r, c, initial_text=None, select_all=False, cursor_at_end=False):
        self._commit_edit()
        col = self._columns[c]
        val = initial_text if initial_text is not None else str(self._rows[r].get(col["key"]) or "")
        self._entry.delete(0, "end")
        self._entry.insert(0, val)
        if select_all:
            self._entry.select_range(0, "end")
        if cursor_at_end or initial_text is not None:
            self._entry.icursor("end")

        x, w = self._col_lefts[c], self._col_widths[c]
        y = r * self.row_height
        self._canvas_body.coords(self._entry_win, x + 1, y + 1)
        self._canvas_body.itemconfigure(
            self._entry_win, width=w - 2, height=self.row_height - 2, state="normal")
        self._entry.focus_set()
        self._editing = (r, c)

    def _commit_edit(self):
        if self._editing is None:
            return
        r, c = self._editing
        self._editing = None
        newval = self._entry.get()
        try:
            self._canvas_body.itemconfigure(self._entry_win, state="hidden")
        except Exception:
            pass
        self._commit_cell(r, c, newval)
        self._redraw()

    def _cancel_edit(self):
        if self._editing is None:
            return
        self._editing = None
        try:
            self._canvas_body.itemconfigure(self._entry_win, state="hidden")
        except Exception:
            pass
        self._redraw()
        self._canvas_body.focus_set()

    def _entry_return(self, event):
        self._commit_edit()
        self._move_active(1, 0)
        return "break"

    def _entry_tab(self, event):
        self._commit_edit()
        self._advance(1)
        return "break"

    def _entry_shift_tab(self, event):
        self._commit_edit()
        self._advance(-1)
        return "break"

    def _toggle_bool(self, r, c):
        col = self._columns[c]
        actual = bool(self._rows[r].get(col["key"]))
        self._commit_cell(r, c, not actual)

    # ── aplicar cambios (delega en el callback del que use el widget) ────
    def _commit_cell(self, r, c, newval):
        col = self._columns[c]
        key = col["key"]
        old = self._rows[r].get(key)
        if old == newval or (old in (None, "") and newval in (None, "")):
            return True
        if self._on_edit:
            try:
                ok = bool(self._on_edit(self._rows[r], key, newval))
            except Exception as e:
                self._toast(f"✘ Error guardando: {e}", "error")
                return False
            if not ok:
                self._toast("✘ No se pudo guardar el cambio", "error")
                return False
        self._rows[r][key] = newval
        return True

    # ── copiar / pegar estilo Excel ──────────────────────────────────────
    def _on_copy(self, event):
        r0, r1, c0, c1 = self._sel_bounds()
        lineas = []
        for r in range(r0, r1 + 1):
            vals = []
            for c in range(c0, c1 + 1):
                col = self._columns[c]
                v = self._rows[r].get(col["key"])
                if col["type"] == "bool":
                    vals.append("1" if v else "0")
                else:
                    vals.append("" if v is None else str(v))
            lineas.append("\t".join(vals))
        texto = "\n".join(lineas)
        self.clipboard_clear()
        self.clipboard_append(texto)
        n = (r1 - r0 + 1) * (c1 - c0 + 1)
        self._toast(f"✔ Copiado ({n} celda{'s' if n != 1 else ''})")
        return "break"

    def _on_paste(self, event):
        try:
            clip = self.clipboard_get()
        except Exception:
            return "break"
        filas = clip.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if filas and filas[-1] == "":
            filas.pop()
        bloque = [f.split("\t") for f in filas]

        r0, c0 = self._active
        cambios = 0
        for i, vals in enumerate(bloque):
            r = r0 + i
            if r >= len(self._rows):
                break
            for j, val in enumerate(vals):
                c = c0 + j
                if c >= len(self._columns):
                    break
                col = self._columns[c]
                if not col.get("editable", True):
                    continue
                val = val.strip()
                nuevo = (val.lower() in _VERDADEROS) if col["type"] == "bool" else val
                if self._commit_cell(r, c, nuevo):
                    cambios += 1
        self._redraw()
        self._toast(f"✔ Pegado ({cambios} celdas actualizadas)")
        return "break"

