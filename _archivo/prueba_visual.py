import tkinter as tk
from tkinter import messagebox

def mostrar_mensaje():
    texto = entrada.get().strip()
    if texto:
        messagebox.showinfo("Mensaje", f"Hola, {texto} aaa")
    else:
        messagebox.showwarning("Advertencia", "Por favor, ingresa tu nombre.")

