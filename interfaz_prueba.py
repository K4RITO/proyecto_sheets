"""
Buscador de beneficiarios IMDEL - Interfaz Tkinter
----------------------------------------------------
Interfaz grafica simple para buscar un DNI en las distintas hojas de
la planilla "Tablas IMDEL Prototipo" y mostrar en que programas figura
como beneficiario. Tambien permite actualizar las bases de datos
(volver a traer los datos desde Google Sheets) con un boton.

Reemplazá CREDENTIALS_FILE por el nombre real de tu archivo de
credenciales de la cuenta de servicio si cambia.
"""

import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import gspread

# ------------------------------------------------------------------
# Configuracion de conexion
# ------------------------------------------------------------------
CREDENTIALS_FILE = "proyecto-sheets-505320-bb8ae069f096.json"
SHEET_NAME = "Tablas IMDEL Prototipo"


# ------------------------------------------------------------------
# Funciones de limpieza de DNI (replican la logica del script original
# para las hojas donde el DNI viene mezclado con otro texto)
# ------------------------------------------------------------------
def _limpiar_recorte(valor, min_len):
    """Extrae los digitos de un valor y descarta los primeros 2 y el
    ultimo digito encontrado. Usado en formulario_inscripcion y
    nominalizacion."""
    if len(valor) >= min_len:
        digitos = [c for c in valor if c.isdigit()]
        if len(digitos) >= 10:
            digitos = digitos[2:-1]
        return "".join(digitos)
    return valor


def _limpiar_solo_digitos(valor, min_len):
    """Extrae solo los digitos de un valor. Usado en usuarios."""
    if len(valor) >= min_len:
        return "".join(c for c in valor if c.isdigit())
    return valor

# Configuracion de cada hoja
# Cada entrada define:
#   - nombre_interno: clave para guardar los datos cargados
#   - worksheet_index: indice de la hoja dentro del Google Sheet (0-based)
#   - col_dni: indice de la columna donde esta el DNI (0-based)
#   - limpiar: funcion opcional (valor, min_len) -> valor limpio,
#              aplicada antes de comparar (y se cachea en la fila,
#              igual que en el script original)
#   - min_len: longitud minima para disparar la limpieza (si aplica)
#   - formatear: funcion que recibe la fila (dato) y devuelve el texto
#                a mostrar cuando hay coincidencia
# ------------------------------------------------------------------
SHEETS_CONFIG = [
    {
        "nombre_interno": "oficina_empleo",
        "worksheet_index": 0,
        "col_dni": 2,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[1]}\nPrograma: Oficina empleo"
        ),
    },
    {
        "nombre_interno": "desarrollo_agrario",
        "worksheet_index": 1,
        "col_dni": 2,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[1]} {dato[0]}\n"
            f"Coordinacion: Desarrollo agrario\n"
            f"Programa: Huertas familiares"
        ),
    },
    {
        "nombre_interno": "registro_recepcion",
        "worksheet_index": 2,
        "col_dni": 6,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[3]} {dato[2]}\n"
            f"Coordinacion: Capacitacion laboral y empleo\n"
            f"Programa: Fortalecimiento de Trayectorias Laborales"
        ),
    },
    {
        "nombre_interno": "inscripcion_cuatrimestral",
        "worksheet_index": 3,
        "col_dni": 5,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[3]} {dato[2]}\n"
            f"Coordinacion: Capacitacion laboral y empleo\n"
            f"Programa: Capacitacion laboral"
        ),
    },
    {
        "nombre_interno": "formulario_inscripcion",
        "worksheet_index": 4,
        "col_dni": 39,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[37]} {dato[38]}\n"
            f"Coordinacion: Capacitacion laboral y empleo\n"
            f"Programa: Insercion laboral"
        ),
    },
    {
        "nombre_interno": "nominalizacion",
        "worksheet_index": 5,
        "col_dni": 9,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[7]} {dato[8]}\n"
            f"Coordinacion: Capacitacion laboral y empleo\n"
            f"Programa: Plan fines"
        ),
    },
    {
        "nombre_interno": "usuarios",
        "worksheet_index": 6,
        "col_dni": 5,
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[2]} {dato[1]}\n"
            f"Coordinacion: Economia popular\n"
            f"Programa: Registro de Trabajadores de la Economia Popular"
        ),
    },
]


class BuscadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de beneficiarios - IMDEL")
        self.root.geometry("640x560")
        self.root.resizable(True, True)

        # Estado interno
        self.gc = None
        self.proyecto_sheets = None
        self.worksheets = {}   # nombre_interno -> objeto worksheet
        self.registros = {}    # nombre_interno -> lista de filas (get_all_values)
        self.conectado = False

        self._armar_interfaz()

        # Conectar en segundo plano al iniciar para no congelar la UI
        self._conectar_en_hilo()

    # ------------------------------------------------------------------
    # Construccion de la interfaz
    # ------------------------------------------------------------------
    def _armar_interfaz(self):
        contenedor = ttk.Frame(self.root, padding=12)
        contenedor.pack(fill="both", expand=True)

        # --- Fila de estado de conexion ---
        self.estado_var = tk.StringVar(value="Conectando a Google Sheets...")
        estado_label = ttk.Label(contenedor, textvariable=self.estado_var, foreground="gray")
        estado_label.pack(anchor="w", pady=(0, 8))

        # --- Fila de busqueda ---
        fila_busqueda = ttk.Frame(contenedor)
        fila_busqueda.pack(fill="x", pady=(0, 8))

        ttk.Label(fila_busqueda, text="DNI:").pack(side="left")

        validar_dni = (self.root.register(self._validar_entrada_dni), "%P")
        self.dni_var = tk.StringVar()
        self.dni_entry = ttk.Entry(
            fila_busqueda, textvariable=self.dni_var, width=20,
            validate="key", validatecommand=validar_dni,
        )
        self.dni_entry.pack(side="left", padx=(6, 6))
        self.dni_entry.bind("<Return>", lambda e: self.buscar_dni())

        self.btn_buscar = ttk.Button(fila_busqueda, text="Buscar", command=self.buscar_dni)
        self.btn_buscar.pack(side="left", padx=(0, 6))

        self.btn_actualizar = ttk.Button(
            fila_busqueda, text="Actualizar bases de datos", command=self.actualizar_bases
        )
        self.btn_actualizar.pack(side="left")

        # --- Area de resultados ---
        ttk.Label(contenedor, text="Resultados:").pack(anchor="w")
        self.resultado_text = scrolledtext.ScrolledText(
            contenedor, wrap="word", height=24, state="disabled", font=("Consolas", 10)
        )
        self.resultado_text.pack(fill="both", expand=True, pady=(4, 0))

    def _validar_entrada_dni(self, valor_propuesto):
        # Permite vacio (para poder borrar) o solo digitos
        return valor_propuesto == "" or valor_propuesto.isdigit()

    # ------------------------------------------------------------------
    # Conexion y carga de datos
    # ------------------------------------------------------------------
    def _conectar_en_hilo(self):
        self._set_controles_habilitados(False)
        hilo = threading.Thread(target=self._conectar, daemon=True)
        hilo.start()

    def _conectar(self):
        try:
            self.gc = gspread.service_account(filename=CREDENTIALS_FILE)
            self.proyecto_sheets = self.gc.open(SHEET_NAME)

            for hoja in SHEETS_CONFIG:
                ws = self.proyecto_sheets.get_worksheet(hoja["worksheet_index"])
                self.worksheets[hoja["nombre_interno"]] = ws

            self._cargar_datos()
            self.conectado = True
            self.root.after(0, lambda: self.estado_var.set(
                "Conectado. Bases de datos cargadas."
            ))
        except Exception as e:
            self.root.after(0, lambda: self._mostrar_error_conexion(e))
        finally:
            self.root.after(0, lambda: self._set_controles_habilitados(True))

    def _cargar_datos(self):
        for nombre_interno, ws in self.worksheets.items():
            self.registros[nombre_interno] = ws.get_all_values()

    def _mostrar_error_conexion(self, error):
        self.estado_var.set("Error al conectar. Ver detalle.")
        messagebox.showerror(
            "Error de conexion",
            f"No se pudo conectar a Google Sheets:\n{error}",
        )

    def _set_controles_habilitados(self, habilitados):
        estado = "normal" if habilitados else "disabled"
        self.btn_buscar.config(state=estado)
        self.btn_actualizar.config(state=estado)
        self.dni_entry.config(state=estado)

    # ------------------------------------------------------------------
    # Boton: actualizar bases
    # ------------------------------------------------------------------
    def actualizar_bases(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return
        self.estado_var.set("Actualizando bases de datos...")
        self._set_controles_habilitados(False)
        hilo = threading.Thread(target=self._actualizar_bases_hilo, daemon=True)
        hilo.start()

    def _actualizar_bases_hilo(self):
        try:
            self._cargar_datos()
            self.root.after(0, lambda: self.estado_var.set("Bases de datos actualizadas."))
        except Exception as e:
            self.root.after(0, lambda: self._mostrar_error_conexion(e))
        finally:
            self.root.after(0, lambda: self._set_controles_habilitados(True))

    # ------------------------------------------------------------------
    # Boton: buscar DNI
    # ------------------------------------------------------------------
    def buscar_dni(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return

        dni_buscar = self.dni_var.get().strip()
        if not dni_buscar or not dni_buscar.isdigit():
            messagebox.showwarning("DNI invalido", "Ingrese solo numeros (ej: 12345678).")
            return

        self._escribir_resultado("", limpiar=True)
        encontrado = False

        for hoja in SHEETS_CONFIG:
            col_dni = hoja["col_dni"]
            nombre_interno = hoja["nombre_interno"]
            limpiar_fn = hoja.get("limpiar")
            min_len = hoja.get("min_len")

            datos_hoja = self.registros.get(nombre_interno, [])
            contador = 0
            primer_mensaje = None

            for dato in datos_hoja:
                if len(dato) <= col_dni:
                    continue

                valor = dato[col_dni]
                if limpiar_fn is not None:
                    valor = limpiar_fn(valor, min_len)
                    dato[col_dni] = valor  # cachear, igual que en el script original

                if valor == dni_buscar:
                    encontrado = True
                    contador += 1
                    if contador == 1:
                        try:
                            primer_mensaje = hoja["formatear"](dato)
                        except IndexError:
                            primer_mensaje = (
                                f"(Fila encontrada en {nombre_interno} pero con "
                                f"columnas insuficientes para mostrar el detalle)"
                            )

            if contador > 0:
                self._escribir_resultado(primer_mensaje)
                if contador > 1:
                    self._escribir_resultado(f"Se encontro al beneficiario {contador} veces")
                self._escribir_resultado("-" * 40)

        if not encontrado:
            self._escribir_resultado(
                f"El DNI ingresado {dni_buscar} no se encontro en las bases de datos."
            )

    def _escribir_resultado(self, texto, limpiar=False):
        self.resultado_text.config(state="normal")
        if limpiar:
            self.resultado_text.delete("1.0", "end")
        else:
            self.resultado_text.insert("end", texto + "\n")
        self.resultado_text.config(state="disabled")
        self.resultado_text.see("end")


def main():
    root = tk.Tk()
    BuscadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()