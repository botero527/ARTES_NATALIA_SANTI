import tkinter as tk
def saludar():
    print("golaaaa")

ventana = tk.Tk()
ventana.geometry("300x200")

boton = tk.Button(
    ventana,
    text="saludar",
    command =saludar
)
boton.pack(pady=20)

ventana.mainloop()