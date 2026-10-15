# -*- coding: utf-8 -*-
"""
Pestaña "Logos Laser" — grilla estilo Excel para llevar el registro de logos
láser: ID | Cliente | Producto | DXF | SJF | Validado | Cargado en Genesis |
Azure | App | Observaciones.

Archivo completamente aparte de agp_app.py a propósito — toda la lógica
visual y de negocio de esta pestaña vive acá. agp_app.py solo la importa y
la registra en PAGES / _PERMISOS / _PAGE_SUBTITLES (buscar "Logos Laser" en
ese archivo). Así cualquier cambio futuro a esta pestaña no puede romper
Crear Arte / Consultar BD / Gestión BD / Scanner, que ya están funcionando.

La grilla en sí (selección arrastrando, copiar/pegar, Tab, checkboxes) la
pone excel_grid.ExcelGrid — este archivo solo la conecta con mallas.LOGOS_LASER
a través de db_app/logos_laser.py.
"""

import threading
import customtkinter as ctk

from excel_grid import ExcelGrid
from db_app.logos_laser import listar, crear_fila, actualizar_campo

try:
    # Rastreo de actividad — mismo módulo que usa agp_app.py. Si por algo
    # no carga, esta pestaña sigue funcionando igual (solo no queda anotado).
    from db_app.tracking import registrar_evento as _registrar_evento
except Exception:
    def _registrar_evento(*_a, **_k): pass

# Paleta propia — a propósito NO se importa de agp_app.py (así esta pestaña
# queda 100% independiente, sin riesgo de import circular). Son los mismos
# valores que PAL en agp_app.py para que se vea igual de una vez la abras.
PAL_LL = {
    "bg":      "#0F1117",
    "card":    "#1C2333",
    "card2":   "#222A3E",
    "border":  "#2E3A55",
    "accent":  "#3B82F6",
    "accent2": "#1D4ED8",
    "green":   "#10B981",
    "red":     "#EF4444",
    "txt":     "#F1F5F9",
    "txt_mid": "#94A3B8",
    "txt_dim": "#475569",
}

FONT_LL = lambda s, w="normal": ctk.CTkFont(family="Segoe UI", size=s, weight=w)

# Columnas de la grilla, en el mismo orden del Excel que compartió Andrés.
# "id" no es editable (lo pone la BD solo), las de check son tipo "bool".
_COLUMNS = [
    {"key": "id",             "header": "ID",                "width": 90,
     "type": "text", "editable": False, "color": "#6ee7b7"},
    {"key": "cliente",        "header": "Cliente",           "width": 170, "type": "text"},
    {"key": "producto",       "header": "Producto",          "width": 170, "type": "text"},
    {"key": "dxf",             "header": "DXF",                "width": 70,  "type": "bool", "align": "center"},
    {"key": "sjf",             "header": "SJF",                "width": 70,  "type": "bool", "align": "center"},
    {"key": "validado",        "header": "Validado",           "width": 90,  "type": "bool", "align": "center"},
    {"key": "cargado_genesis", "header": "Cargado en Genesis", "width": 150, "type": "bool", "align": "center"},
    {"key": "azure",           "header": "Azure",              "width": 70,  "type": "bool", "align": "center"},
    {"key": "app",             "header": "App",                "width": 70,  "type": "bool", "align": "center"},
    {"key": "observaciones",   "header": "Observaciones",      "width": 260, "type": "text", "stretch": True},
]


class TabLogosLaser(ctk.CTkFrame):
    def __init__(self, parent, usuario_info=None, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._usuario = usuario_info or {}
        self._todas = []   # cache completo desde BD — el buscador filtra acá sin volver a consultar
        self._build()
        self.refresh()

    # ── UI ───────────────────────────────────────────────────────────────
    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Barra superior: título, buscador, nueva fila, refrescar
        top = ctk.CTkFrame(self, fg_color=PAL_LL["card"], corner_radius=10)
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))

        ctk.CTkLabel(top, text="🔶  Logos Laser", font=FONT_LL(15, "bold"),
                     text_color=PAL_LL["txt"]).pack(side="left", padx=16, pady=10)

        self._buscador = ctk.CTkEntry(
            top, placeholder_text="Buscar cliente / producto / ID...",
            width=240, height=32, fg_color=PAL_LL["card2"],
            border_width=0, font=FONT_LL(11))
        self._buscador.pack(side="left", padx=(4, 8), pady=10)
        self._buscador.bind("<KeyRelease>", lambda e: self._aplicar_filtro())

        ctk.CTkButton(top, text="↺  Actualizar", width=110, height=32,
                      fg_color=PAL_LL["card2"], hover_color=PAL_LL["border"],
                      font=FONT_LL(11), command=self.refresh
                      ).pack(side="right", padx=(4, 12), pady=10)
        ctk.CTkButton(top, text="➕  Nueva fila", width=130, height=32,
                      fg_color=PAL_LL["accent"], hover_color=PAL_LL["accent2"],
                      font=FONT_LL(11, "bold"), command=self._nueva_fila
                      ).pack(side="right", padx=4, pady=10)

        # Grilla — tarjeta que la envuelve
        card = ctk.CTkFrame(self, fg_color=PAL_LL["card"], corner_radius=10)
        card.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 6))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        tema_grid = {
            "bg": PAL_LL["card"], "header_bg": PAL_LL["card2"], "header_fg": PAL_LL["txt_mid"],
            "row_bg": PAL_LL["card"], "row_bg_alt": "#1A2136", "border": PAL_LL["border"],
            "text": PAL_LL["txt"], "text_dim": PAL_LL["txt_dim"], "accent": PAL_LL["accent"],
            "select_bg": "#22406E", "input_bg": PAL_LL["card2"],
        }
        self._grid = ExcelGrid(
            card, _COLUMNS, on_edit=self._on_edit, status_cb=self._status, theme=tema_grid)
        self._grid.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # Pie: recordatorio de atajos + estado de la última acción
        pie = ctk.CTkFrame(self, fg_color="transparent")
        pie.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 14))
        ctk.CTkLabel(
            pie, text="✏ Doble clic o F2 edita · Tab / flechas para moverte · "
                       "arrastra para seleccionar varias celdas · Ctrl+C / Ctrl+V "
                       "copia y pega como en Excel · un clic marca los check",
            font=FONT_LL(9), text_color=PAL_LL["txt_dim"]).pack(side="left")
        self._status_lbl = ctk.CTkLabel(pie, text="", font=FONT_LL(10, "bold"))
        self._status_lbl.pack(side="right")

    # ── feedback visual (toast simple en el pie) ──────────────────────────
    def _status(self, msg, level="ok"):
        color = {"ok": PAL_LL["green"], "error": PAL_LL["red"]}.get(level, PAL_LL["txt_mid"])
        self._status_lbl.configure(text=msg, text_color=color)
        self._status_lbl.after(2500, lambda: self._status_lbl.configure(text=""))

    # ── carga desde BD ────────────────────────────────────────────────────
    def refresh(self):
        """Recarga todo desde BD. Va en hilo aparte para no congelar la UI."""
        self._status("Cargando...", "info")
        threading.Thread(target=self._cargar_bg, daemon=True).start()

    def _cargar_bg(self):
        try:
            filas = listar()
            self.after(0, lambda: self._on_cargado(filas))
        except Exception as e:
            self.after(0, lambda: self._status(f"✘ Error cargando: {e}", "error"))

    def _on_cargado(self, filas):
        self._todas = filas
        self._aplicar_filtro()
        self._status(f"✔ {len(filas)} registro{'s' if len(filas) != 1 else ''}")

    def _aplicar_filtro(self):
        texto = (self._buscador.get() or "").strip().lower()
        if not texto:
            vistas = self._todas
        else:
            vistas = [
                f for f in self._todas
                if texto in (f.get("cliente") or "").lower()
                or texto in (f.get("producto") or "").lower()
                or texto in (f.get("id") or "").lower()
            ]
        self._grid.set_rows(vistas)

    # ── acciones del usuario ──────────────────────────────────────────────
    def _nueva_fila(self):
        responsable = self._usuario.get("nombre") or self._usuario.get("usuario")
        try:
            fila = crear_fila(responsable=responsable)
        except Exception as e:
            self._status(f"✘ No se pudo crear la fila: {e}", "error")
            return
        # La grilla está ordenada más reciente primero (igual que el resto
        # de mallas.*), así que la fila nueva va al principio y queda visible
        # de una — se limpia el buscador para asegurarse de que no quede
        # tapada por un filtro que no le calce.
        self._todas.insert(0, fila)
        self._buscador.delete(0, "end")
        self._grid.set_rows(self._todas)
        self._grid.focus_grid()
        self._status(f"✔ Fila {fila['id']} creada")
        _registrar_evento("nueva_fila_logos_laser", pestana="Logos Laser",
                           detalle=fila["id"], exito=True)

    def _on_edit(self, row, key, valor):
        """Callback de ExcelGrid: guarda un campo editado. True/False = éxito."""
        try:
            ok = actualizar_campo(row["id"], key, valor)
        except Exception as e:
            self._status(f"✘ Error guardando: {e}", "error")
            _registrar_evento("editar_logos_laser", pestana="Logos Laser",
                               detalle=f"{row['id']}.{key}: {e}"[:200], exito=False)
            return False
        _registrar_evento("editar_logos_laser", pestana="Logos Laser",
                           detalle=f"{row['id']}.{key}", exito=ok)
        return ok
