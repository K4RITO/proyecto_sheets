import tkinter as tk

ventana = tk.Tk()
ventana.title("Aplicacion de busqueda")
ventana.geometry("500x500")

#
buscar = tk.Button(text="Buscar", background="lightgreen", cursor="hand2")
salir = tk.Button(text="Salir", background="red", cursor="hand2")
actualizar = tk.Button(text="Actualizar", background="lightblue", cursor="hand2", activebackground="#1d6bd1")
etiqueta = tk.Label(text="Ingrese el DNI: ", cursor="xterm")
barra_busqueda = tk.Entry()
texto = tk.Text(width=25, height=10)
#
buscar.pack()
salir.pack()
actualizar.pack()
etiqueta.pack()
barra_busqueda.pack()
texto.pack()

#
respuesta = "PENE\n"
texto.insert(tk.END, respuesta)
texto.config(state="disabled")
texto.config(state="normal")
texto.insert(tk.END, respuesta)
texto.config(state="disabled")



ventana.mainloop()