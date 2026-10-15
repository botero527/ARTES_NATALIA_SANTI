#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Ventana de login AGP Glass
"""
import threading
import customtkinter as ctk
from customtkinter import CTkFont

# Paleta igual que la app principal
PAL = {
    "bg":       "#0F1117",
    "card":     "#1C2333",
    "card2":    "#242C3D",
    "accent":   "#3B82F6",
    "accent_h": "#2563EB",
    "txt":      "#F1F5F9",
    "txt_mid":  "#94A3B8",
    "border":   "#2D3A4F",
    "err":      "#EF4444",
    "ok":       "#22C55E",
}


class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.usuario_info = None   # se rellena al hacer login exitoso

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("AGP Glass — Iniciar sesión")
        self.geometry("440x520")
        self.resizable(False, False)
        self.configure(fg_color=PAL["bg"])

        # Centrar en pantalla
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 440) // 2
        y = (self.winfo_screenheight() - 520) // 2
        self.geometry(f"440x520+{x}+{y}")

        self._build()
        self._inicializar_bd()

    # ──────────────────────────────────────────────
    def _build(self):
        # Card central
        card = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=20,
                            border_width=1, border_color=PAL["border"])
        card.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.88, relheight=0.90)

        # Logo / título
        ctk.CTkLabel(card, text="AGP", font=CTkFont(size=48, weight="bold"),
                     text_color=PAL["accent"]).pack(pady=(36, 0))
        ctk.CTkLabel(card, text="Glass Engineering",
                     font=CTkFont(size=14), text_color=PAL["txt_mid"]).pack(pady=(0, 6))
        ctk.CTkLabel(card, text="Iniciar sesión",
                     font=CTkFont(size=18, weight="bold"),
                     text_color=PAL["txt"]).pack(pady=(10, 28))

        # Campo usuario
        ctk.CTkLabel(card, text="Usuario", font=CTkFont(size=12),
                     text_color=PAL["txt_mid"], anchor="w").pack(padx=36, fill="x")
        self._ent_user = ctk.CTkEntry(
            card, height=44, corner_radius=10,
            fg_color=PAL["card2"], border_color=PAL["border"],
            text_color=PAL["txt"], font=CTkFont(size=13),
            placeholder_text="correo@agpglass.com",
        )
        self._ent_user.pack(padx=36, fill="x", pady=(4, 14))

        # Campo contraseña
        ctk.CTkLabel(card, text="Contraseña", font=CTkFont(size=12),
                     text_color=PAL["txt_mid"], anchor="w").pack(padx=36, fill="x")
        self._ent_pass = ctk.CTkEntry(
            card, height=44, corner_radius=10,
            fg_color=PAL["card2"], border_color=PAL["border"],
            text_color=PAL["txt"], font=CTkFont(size=13),
            show="●", placeholder_text="Contraseña",
        )
        self._ent_pass.pack(padx=36, fill="x", pady=(4, 6))

        # Label de estado / error
        self._lbl_estado = ctk.CTkLabel(
            card, text="", font=CTkFont(size=12),
            text_color=PAL["err"], height=20,
        )
        self._lbl_estado.pack(pady=(4, 12))

        # Botón login
        self._btn_login = ctk.CTkButton(
            card, text="Entrar", height=48, corner_radius=12,
            fg_color=PAL["accent"], hover_color=PAL["accent_h"],
            font=CTkFont(size=14, weight="bold"),
            command=self._on_login,
        )
        self._btn_login.pack(padx=36, fill="x")

        # Footer
        ctk.CTkLabel(card, text="AGP Group · Ingeniería Colombia",
                     font=CTkFont(size=10), text_color=PAL["txt_mid"]).pack(pady=(20, 0))

        # Enter en cualquier campo → login
        self._ent_user.bind("<Return>", lambda e: self._ent_pass.focus())
        self._ent_pass.bind("<Return>", lambda e: self._on_login())
        self._ent_user.focus()

    # ──────────────────────────────────────────────
    def _inicializar_bd(self):
        """Crea la tabla si no existe (en background para no bloquear la UI)."""
        def _init():
            try:
                from db_app.auth import crear_tabla
                crear_tabla()
            except Exception as e:
                print(f"[login] init BD error: {e}")
        threading.Thread(target=_init, daemon=True).start()

    # ──────────────────────────────────────────────
    def _on_login(self):
        usuario    = self._ent_user.get().strip()
        contrasenia = self._ent_pass.get().strip()

        if not usuario or not contrasenia:
            self._set_estado("Completa todos los campos", error=True)
            return

        self._btn_login.configure(state="disabled", text="Verificando...")
        self._set_estado("")

        def _worker():
            try:
                from db_app.auth import validar_login, sincronizar_background
                info = validar_login(usuario, contrasenia)
            except Exception as e:
                info = None
                self.after(0, lambda: self._set_estado(f"Error de conexión: {e}", error=True))
                self.after(0, lambda: self._btn_login.configure(state="normal", text="Entrar"))
                return

            if info:
                # Login OK → sync en background
                sincronizar_background()
                self.after(0, lambda: self._login_exitoso(info))
            else:
                self.after(0, lambda: self._set_estado(
                    "Usuario o contraseña incorrectos", error=True))
                self.after(0, lambda: self._btn_login.configure(state="normal", text="Entrar"))

        threading.Thread(target=_worker, daemon=True).start()

    def _login_exitoso(self, info):
        self._set_estado(f"Bienvenido, {info['nombre'].split()[0]}", error=False)
        self.usuario_info = info
        self.after(400, self.destroy)

    def _set_estado(self, msg, error=True):
        color = PAL["err"] if error else PAL["ok"]
        self._lbl_estado.configure(text=msg, text_color=color)
