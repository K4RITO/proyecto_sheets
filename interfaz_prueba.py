"""
Buscador de beneficiarios IMDEL - Interfaz Tkinter
----------------------------------------------------
Interfaz grafica para buscar un beneficiario de los distintos programas del IMDEL 
en las distintas hojas de la planilla "Tablas IMDEL Prototipo 2" 
mostrar en que programas figura como beneficiario. 
"""

import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import re
import csv
import os
from datetime import datetime

import gspread

# Configuracion de conexion
CREDENTIALS_FILE = "proyecto-sheets-505320-bb8ae069f096.json"
SHEET_NAME = "Tablas IMDEL Prototipo 2"

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
#   - domicilio_index: posicion en la tabla de los campos domicilio y altura
#   - limpiar: funcion opcional (valor, min_len) -> valor limpio,
#              aplicada antes de comparar (y se cachea en la fila,
#              igual que en el script original)
#   - min_len: longitud minima para disparar la limpieza (si aplica)
#   - formatear: funcion que recibe la fila (dato) y devuelve el texto
#                a mostrar cuando hay coincidencia
# ------------------------------------------------------------------
SHEETS_CONFIG = [
    {
        "nombre_interno": "INSCRIPCIÓN CUATRIMESTRAL CAPACITACION EN OFICIOS",
        "worksheet_index": 0,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "Coordinacion": "Capacitación laboral y empleo",
        "Programa": "Capacitación laboral",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[2]} {dato[1]}\n"
            f"Coordinacion: Capacitación laboral y empleo\n"
            f"Programa: Capacitación laboral\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    },
    {
        "nombre_interno": "OFICINA EMPLEO",
        "worksheet_index": 1,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,            
        "Coordinacion":"Capacitación laboral y empleo",
        "Programa":"Inserción laboral",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[1]} {dato[2]}\n"
            f"Coordinacion: Capacitación laboral y empleo\n"
            "Programa: Inserción laboral\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    },
    {
        "nombre_interno": "NOMINALIZACIÓN",
        "worksheet_index": 2,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "Coordinacion": "Capacitacón laboral y empleo",
        "Programa": "Plan FINES",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[2]} {dato[1]}\n"
            f"Coordinacion: Capacitacón laboral y empleo\n"
            f"Programa: Plan FINES\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    },
    {
        "nombre_interno": "DESARROLLO AGRARIO",
        "worksheet_index": 3,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "Coordinacion": "Desarrollo agrario",
        "Programa": "Huertas familiares - Entregas KIT de semillas",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[2]} {dato[1]}\n"
            f"Coordinacion: Desarrollo agrario\n"
            f"Programa: Huertas familiares - Entregas KIT de semillas\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    },
    {
        "nombre_interno": "AVICULTURA EN COMUNIDAD",
        "worksheet_index": 4,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "Coordinacion": "Desarrollo agrario",
        "Programa": "Avicultura en comunidad",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[1]} {dato[2]}\n"
            f"Coordinacion: Desarrollo agrario\n"
            f"Programa: Avicultura en comunidad\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    },    
    {
        "nombre_interno": "PRODUCTORES",
        "worksheet_index": 5,
        "col_dni": 0,
        "domicilio_index": (3, 4),
        "limpiar": _limpiar_recorte,
        "min_len": 9,
        "Coordinacion": "Desarrollo agrario",
        "Programa": "RETEP - Registro de trabajadores de la economía popular",
        "formatear": lambda dato: (
            f"Apellido y nombre: {dato[1]} {dato[2]}\n"
            f"Coordinacion: Desarrollo agrario\n"
            f"Programa: RETEP - Registro de trabajadores de la economía popular\n"
            "Datos de contacto:\n"
            f"Telefono: {dato[5]}\n"
            f"Email: {dato[6]}\n"
            f"Domicilio: {dato[3]} {dato[4]}"
        ),
    }, 
]


class BuscadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de beneficiarios - IMDEL")
        self.root.geometry("800x600")
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

        # --- Fila de busqueda por DNI---
        fila_busqueda_dni = ttk.Frame(contenedor)
        fila_busqueda_dni.pack(fill="x", pady=(0, 8))

        ttk.Label(fila_busqueda_dni, text="DNI:").pack(side="left")

        validar_dni = (self.root.register(self._validar_entrada_dni), "%P")
        self.dni_var = tk.StringVar()
        self.dni_entry = ttk.Entry(
            fila_busqueda_dni, textvariable=self.dni_var, width=20,
            validate="key", validatecommand=validar_dni,
        )
        self.dni_entry.pack(side="left", padx=(6, 6))
        self.dni_entry.bind("<Return>", lambda e: self.buscar_dni())

        self.btn_buscar = ttk.Button(fila_busqueda_dni, text="Buscar", command=self.buscar_dni)
        self.btn_buscar.pack(side="left", padx=(0, 6))

        # --- Fila de busqueda por Nombre y apellido---
        fila_busqueda_nombre = ttk.Frame(contenedor)
        fila_busqueda_nombre.pack(fill="x", pady=(0, 8))

        ttk.Label(fila_busqueda_nombre, text="Apellido y Nombre:").pack(side="left")

        validar_nombre = (self.root.register(self._validar_entrada_nombre), "%P")
        self.nombre_var = tk.StringVar()
        self.nombre_entry = ttk.Entry(
            fila_busqueda_nombre, textvariable=self.nombre_var, width=20,
            validate="key", validatecommand=validar_nombre,
        )
        self.nombre_entry.pack(side="left", padx=(6, 6))
        self.nombre_entry.bind("<Return>", lambda e: self.buscar_nombre())

        self.btn_buscar_nombre = ttk.Button(fila_busqueda_nombre, text="Buscar", command=self.buscar_nombre)
        self.btn_buscar_nombre.pack(side="left", padx=(0, 6))

        # --- Fila de busqueda por Calle y Altura---
        fila_busqueda_domicilio = ttk.Frame(contenedor)
        fila_busqueda_domicilio.pack(fill="x", pady=(0, 8))

        ttk.Label(fila_busqueda_domicilio, text="Calle y Altura:").pack(side="left")

        validar_domicilio = (self.root.register(self._validar_entrada_domicilio), "%P")
        self.domicilio_var = tk.StringVar()
        self.domicilio_entry = ttk.Entry(
            fila_busqueda_domicilio, textvariable=self.domicilio_var, width=20,
            validate="key", validatecommand=validar_domicilio,
        )
        self.domicilio_entry.pack(side="left", padx=(6, 6))
        self.domicilio_entry.bind("<Return>", lambda e: self.buscar_domicilio())

        self.btn_buscar_domicilio = ttk.Button(fila_busqueda_domicilio, text="Buscar", command=self.buscar_domicilio)
        self.btn_buscar_domicilio.pack(side="left", padx=(0, 6))

        # --- Boton actualizar bases de datos ---
        self.btn_actualizar = ttk.Button(
            fila_busqueda_domicilio, text="Actualizar bases de datos", command=self.actualizar_bases
        )
        self.btn_actualizar.pack(side="left")

        # --- Area de resultados ---
        ttk.Label(contenedor, text="Resultados:").pack(anchor="w")
        self.resultado_text = scrolledtext.ScrolledText(
            contenedor, wrap="word", height=24, state="disabled", font=("Consolas", 10)
        )
        self.resultado_text.pack(fill="both", expand=True, pady=(4, 0))

        # --- Boton para imprimir resultados ---
        self.resultados_guardados = ""
        fila_busqueda_descargar_resultados = ttk.Frame(contenedor)
        fila_busqueda_descargar_resultados.pack(fill="x", pady=(0, 8))

        self.btn_descargar_resultados = ttk.Button(
            fila_busqueda_descargar_resultados, text="Descargar resultados", command=self.descargar_resultados
        )
        self.btn_descargar_resultados.pack(side="left")

    # --- Validaciones ---

    def _validar_entrada_dni(self, valor_propuesto):
        # Permite vacio (para poder borrar) o solo digitos
        return valor_propuesto == "" or valor_propuesto.isdigit()

    def _validar_entrada_nombre(self, valor_propuesto):
        return valor_propuesto == "" or valor_propuesto.replace(" ", "").isalpha()
    
    def _validar_entrada_domicilio(self, valor_propuesto):
        return valor_propuesto == "" or valor_propuesto.replace(" ", "").isalnum()

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
    # Botones de busqueda: 
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
                    dato[col_dni] = valor

                if valor == dni_buscar:
                    encontrado = True
                    campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                    self.guardar_resultados(campos_guardar)
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
        
    def buscar_nombre(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return

        nombre_buscar = self.nombre_var.get().strip().lower()
        if not nombre_buscar:
            messagebox.showwarning("Nombre invalido", "Debe ingresar texto (ej: Benitez Nicolas).")
            return

        if len(nombre_buscar.split()) < 2:
            messagebox.showwarning("Entrada invalida", "Debe ingresar 2 palabras (ej: Benitez Nicolas).")
            return

        self._escribir_resultado("", limpiar=True)
        encontrado = False
        coincidencias = []

        for hoja in SHEETS_CONFIG:
            col_nombre = 1
            col_apellido = 2
            nombre_interno = hoja["nombre_interno"]
            # limpiar_fn = hoja.get("limpiar")
            # min_len = hoja.get("min_len")
            datos_hoja = self.registros.get(nombre_interno, [])
            contador = 0
            primer_mensaje = None

            for dato in datos_hoja:
                if len(dato) <= col_nombre:
                    continue

                if (len(dato[col_apellido]) < 1 or len(dato[col_nombre]) < 1):
                    continue
                # entra aca cuando hay un espacio
                if (not dato[col_apellido].split()[0].isalpha() or not dato[col_nombre].split()[0].isalpha()):
                    continue

                # print(dato[col_apellido])
                # print(re.match(r'\w+', dato[col_apellido]))
                primer_apellido = re.match(r'\w+', dato[col_apellido].strip()).group().lower()
                # print(primer_apellido)
                primer_nombre = re.match(r'\w+', dato[col_nombre].strip()).group().lower()

                valor = f"{primer_apellido} {primer_nombre}"

                # TODO: validar si es necesario limpiar el string
                # if limpiar_fn is not None:
                #     # valor = limpiar_fn(valor, min_len)
                #     dato[col_dni] = valor
                if valor == nombre_buscar:
                    encontrado = True
                    campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                    coincidencias.append(campos_guardar)
                    contador += 1
                    if contador == 1:
                        try:
                            dato[1] = dato[1].capitalize()
                            dato[2] = dato[2].capitalize()
                            primer_mensaje = f"DNI: {dato[0]} \n{hoja["formatear"](dato)} \n{"-" * 40}"
                        except IndexError:
                            primer_mensaje = (
                                f"(Fila encontrada en {nombre_interno} pero con "
                                f"columnas insuficientes para mostrar el detalle)"
                            )
                    else:
                        dato[1] = dato[1].capitalize()
                        dato[2] = dato[2].capitalize()
                        primer_mensaje = primer_mensaje + f"\nDNI: {dato[0]}\n{hoja["formatear"](dato)} \n{"-" * 40}"

            if contador > 0:
                self._escribir_resultado(primer_mensaje)
                # self._escribir_resultado("-" * 40)
        
        self.guardar_resultados(coincidencias)
        
        if not encontrado:
            self._escribir_resultado(
                f"El Nombre y apellido ingresado '{nombre_buscar.capitalize()}' no se encontro en las bases de datos."
            )     

    def buscar_domicilio(self):
            if not self.conectado:
                messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
                return
    
            domicilio_buscar = self.domicilio_var.get().strip().lower()
            if not domicilio_buscar:
                messagebox.showwarning("Domicilio invalido", "Debe ingresar calle y altura (ej: Cnel. Pedro Aquino 2356).")
                return
    
            if len(domicilio_buscar.split()) < 2:
                messagebox.showwarning("Entrada invalida", "Debe ingresar minimo 2 palabras (ej: Cnel. Pedro Aquino 2356).")
                return
    
            self._escribir_resultado("", limpiar=True)
            encontrado = False
            coincidencias = []
    
            for hoja in SHEETS_CONFIG:
                col_calle = 3
                col_altura = 4
                nombre_interno = hoja["nombre_interno"]
                # limpiar_fn = hoja.get("limpiar")
                # min_len = hoja.get("min_len")
                datos_hoja = self.registros.get(nombre_interno, [])
                primer_mensaje = None
                contador = 0

                for dato in datos_hoja:
                    if len(dato) <= col_calle:
                        continue

                    if (len(dato[col_calle]) < 1 or len(dato[col_altura]) < 1):
                        continue
                    # entra aca cuando hay un espacio
                    if (not dato[col_calle].replace(" ", "").isalnum() or not dato[col_altura].split()[0].isdigit()):
                        continue
    
                    valor = f"{dato[col_calle].lower()} {dato[col_altura]}"
                    # TODO: validar si es necesario limpiar el string
                    # if limpiar_fn is not None:
                    #     # valor = limpiar_fn(valor, min_len)
                    #     dato[col_dni] = valor
                    if valor == domicilio_buscar:
                        encontrado = True
                        campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                        coincidencias.append(campos_guardar)
                        contador += 1
                        if contador == 1:
                            try:
                                dato[1] = dato[1].capitalize()
                                dato[2] = dato[2].capitalize()
                                primer_mensaje = f"DNI: {dato[0]} \n{hoja["formatear"](dato)} \n{"-" * 40}"
                            except IndexError:
                                primer_mensaje = (
                                    f"(Fila encontrada en {nombre_interno} pero con "
                                    f"columnas insuficientes para mostrar el detalle)"
                                )
                        else:
                            dato[1] = dato[1].capitalize()
                            dato[2] = dato[2].capitalize()
                            primer_mensaje = primer_mensaje + f"\nDNI: {dato[0]}\n{hoja["formatear"](dato)} \n{"-" * 40}"
    
                if contador > 0:
                    self._escribir_resultado(primer_mensaje)
                    # self._escribir_resultado("-" * 40)
            
            self.guardar_resultados(coincidencias)

            if not encontrado:
                self._escribir_resultado(
                    f"El Domicilio ingresado '{domicilio_buscar.capitalize()}' no se encontro en las bases de datos."
                )    

    # ------------------------------------------------------------------
    # Funciones de los resultados: 
    # ------------------------------------------------------------------

    def guardar_resultados(self, datos):
        self.resultados_guardados = datos

    def _escribir_resultado(self, texto, limpiar=False):
        self.resultado_text.config(state="normal")

        if limpiar:
            self.resultado_text.delete("1.0", "end")
        else:
            self.resultado_text.insert("end", texto + "\n")
        self.resultado_text.config(state="disabled")
        self.resultado_text.see("end")

    def descargar_resultados(self):
        # formatear el texto : resultado_text
        # en resultado_text se guarda el resultado actual
        # Obtener carpeta Descargas del usuario
        carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")

        fecha = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        nombre_archivo = f"Registros de busqueda de beneficiarios - {fecha}hs.csv"

        # Ruta completa del archivo
        ruta_archivo = os.path.join(carpeta_descargas, nombre_archivo)

        # Encabezados del archivo CSV
        encabezados = ['DNI', 'Nombre', 'Apellido', 'Calle', 'Altura', 'Numero Telefono', 'Email', 'Coordinacion', 'Programa']
        
        with open(ruta_archivo, "w", newline="", encoding="utf-8") as archivo:
            escritor = csv.writer(archivo)

            escritor.writerow(encabezados)

            for registro in self.resultados_guardados:
                escritor.writerow(registro)

        messagebox.showinfo(
            "Descarga completada",
            f"El archivo se guardó correctamente en:\n{ruta_archivo}"
        )

def main():
    root = tk.Tk()
    BuscadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()