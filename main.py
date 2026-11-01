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
import tkinter.font as tkfont

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

def agrandar_fuentes(root, tamano=12):
    # Fuentes por defecto de Tk: todos los widgets las usan
    for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                "TkHeadingFont", "TkCaptionFont", "TkFixedFont"):
        tkfont.nametofont(nombre).configure(size=tamano)

    # Treeview: la altura de fila no crece sola con la fuente
    estilo = ttk.Style(root)
    estilo.configure("Treeview", rowheight=int(tamano * 2.2))
    estilo.configure("Treeview.Heading", font=("TkDefaultFont", tamano, "bold"))

class BuscadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de beneficiarios - IMDEL")
        self.root.geometry("800x650")
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

        # --- Grilla de busqueda: (label, entry, label2, entry2, button) ---
        grilla = ttk.Frame(contenedor)
        grilla.pack(fill="x", pady=(0, 8))

        # Las columnas de los entries se estiran; label y boton toman su ancho natural
        grilla.columnconfigure(1, weight=1)
        grilla.columnconfigure(3, weight=1)

        opciones_label = {"sticky": "w", "padx": (0, 6), "pady": 4}
        opciones_entry = {"sticky": "ew", "padx": (0, 12), "pady": 4}
        opciones_boton = {"sticky": "ew", "pady": 4}

        validar_dni = (self.root.register(self._validar_entrada_dni), "%P")
        validar_texto = (self.root.register(self._validar_entrada_nombre), "%P")
        validar_calle = (self.root.register(self._validar_entrada_calle), "%P")
        validar_altura = (self.root.register(self._validar_entrada_altura), "%P")

        # Fila 0: DNI
        ttk.Label(grilla, text="Numero de DNI:").grid(row=0, column=0, **opciones_label)
        self.dni_var = tk.StringVar()
        self.dni_entry = ttk.Entry(
            grilla, textvariable=self.dni_var,
            validate="key", validatecommand=validar_dni,
        )
        self.dni_entry.grid(row=0, column=1, **opciones_entry)
        self.dni_entry.bind("<Return>", lambda e: self.buscar_dni())

        self.btn_buscar = ttk.Button(grilla, text="Buscar por DNI", command=self.buscar_dni)
        self.btn_buscar.grid(row=0, column=4, **opciones_boton)

        # Fila 1: Apellido / Nombre
        ttk.Label(grilla, text="Apellido:").grid(row=1, column=0, **opciones_label)
        self.apellido_var = tk.StringVar()
        self.apellido_entry = ttk.Entry(
            grilla, textvariable=self.apellido_var,
            validate="key", validatecommand=validar_texto,
        )
        self.apellido_entry.grid(row=1, column=1, **opciones_entry)
        self.apellido_entry.bind("<Return>", lambda e: self.buscar_nombre())

        ttk.Label(grilla, text="Nombre:").grid(row=1, column=2, **opciones_label)
        self.nombre_var = tk.StringVar()
        self.nombre_entry = ttk.Entry(
            grilla, textvariable=self.nombre_var,
            validate="key", validatecommand=validar_texto,
        )
        self.nombre_entry.grid(row=1, column=3, **opciones_entry)
        self.nombre_entry.bind("<Return>", lambda e: self.buscar_nombre())

        self.btn_buscar_nombre = ttk.Button(grilla, text="Buscar por nombre", command=self.buscar_nombre)
        self.btn_buscar_nombre.grid(row=1, column=4, **opciones_boton)

        # Fila 2: Calle / Altura
        ttk.Label(grilla, text="Calle:").grid(row=2, column=0, **opciones_label)
        self.calle_var = tk.StringVar()
        self.calle_entry = ttk.Entry(
            grilla, textvariable=self.calle_var,
            validate="key", validatecommand=validar_calle,
        )
        self.calle_entry.grid(row=2, column=1, **opciones_entry)
        self.calle_entry.bind("<Return>", lambda e: self.buscar_domicilio())

        ttk.Label(grilla, text="Altura:").grid(row=2, column=2, **opciones_label)
        self.altura_var = tk.StringVar()
        self.altura_entry = ttk.Entry(
            grilla, textvariable=self.altura_var,
            validate="key", validatecommand=validar_altura,
        )
        self.altura_entry.grid(row=2, column=3, **opciones_entry)
        self.altura_entry.bind("<Return>", lambda e: self.buscar_domicilio())

        # Boton deshabilitado hasta que haya una calle
        self.btn_buscar_domicilio = ttk.Button(
            grilla, text="Buscar por calle y altura", command=self.buscar_domicilio,
        )
        self.btn_buscar_domicilio.grid(row=2, column=4, **opciones_boton)
        self.btn_buscar_domicilio.state(["disabled"])
        self.calle_var.trace_add("write", self._actualizar_estado_btn_domicilio)

        # --- Area de resultados ---
        ttk.Label(contenedor, text="Resultados:").pack(anchor="w")
        self.resultado_text = scrolledtext.ScrolledText(
            contenedor, wrap="word", height=20, state="disabled", font=("Consolas", 12) # cambiar aca para agrandar/achicar el tamaño de la letra en el textarea
        )
        self.resultado_text.pack(fill="both", expand=True, pady=(4, 0))

        # --- Botones inferiores ---
        self.resultados_guardados = ""
        fila_inferior = ttk.Frame(contenedor)
        fila_inferior.pack(fill="x", pady=(8, 0))

        self.btn_descargar_resultados = ttk.Button(
            fila_inferior, text="Descargar resultados", command=self.descargar_resultados
        )
        self.btn_descargar_resultados.pack(side="left")

        self.btn_actualizar = ttk.Button(
            fila_inferior, text="Actualizar bases de datos", command=self.actualizar_bases
        )
        self.btn_actualizar.pack(side="right")

    # --- Validaciones ---

    def _validar_entrada_dni(self, valor_propuesto):
        return valor_propuesto == "" or valor_propuesto.isdigit()

    def _validar_entrada_nombre(self, valor_propuesto):
        return valor_propuesto == "" or valor_propuesto.replace(" ", "").isalpha()
    
    # def _validar_entrada_domicilio(self, valor_propuesto):
    #     return valor_propuesto == "" or valor_propuesto.replace(" ", "").isalnum()

    def _validar_entrada_calle(self, texto):
        # Letras (con acentos y ñ), números, espacios y puntos. Máx. 40 caracteres.
        if texto == "":
            return True
        return (
            len(texto) <= 40
            and re.fullmatch(r"[A-Za-zÁÉÍÓÚáéíóúÑñÜü0-9. ]+", texto) is not None
        )

    def _validar_entrada_altura(self, texto):
        # Solo dígitos, entero positivo (sin 0 inicial), máx. 6 dígitos.
        if texto == "":
            return True
        return (
            texto.isascii()
            and texto.isdigit()
            and not texto.startswith("0")
            and len(texto) <= 6
        )

    # ---------- Estado del botón ----------
    def _actualizar_estado_btn_domicilio(self, *args):
        if self.calle_var.get().strip():
            self.btn_buscar_domicilio.state(["!disabled"])
        else:
            self.btn_buscar_domicilio.state(["disabled"])


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
        self.apellido_entry.config(state=estado)
        self.nombre_entry.config(state=estado)
        self.btn_buscar_nombre.config(state=estado)
        self.calle_entry.config(state=estado)
        self.altura_entry.config(state=estado)
        self.btn_buscar_domicilio.config(state=estado)
        self.btn_descargar_resultados.config(state=estado)

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
    # Botones de busqueda
    # ------------------------------------------------------------------
    def buscar_dni(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return

        self.limpiar_campos("dni")

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
                    self.guardar_resultados([campos_guardar])
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

    # TODO: Dividir en 2 funciones buscar nombre y buscar apellido.  
    def buscar_nombre(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return

        self.limpiar_campos("nombre")

        apellido_buscar = self.apellido_var.get().strip().lower()
        nombre_buscar = self.nombre_var.get().strip().lower()

        if not apellido_buscar and not nombre_buscar:
            messagebox.showwarning("Entrada invalida", "Debe ingresar al menos un apellido o un nombre.")
            return

        self._escribir_resultado("", limpiar=True)
        encontrado = False
        coincidencias = []

        for hoja in SHEETS_CONFIG:
            col_nombre = 1
            col_apellido = 2
            nombre_interno = hoja["nombre_interno"]
            datos_hoja = self.registros.get(nombre_interno, [])
            contador = 0
            primer_mensaje = ""

            for dato in datos_hoja:
                if len(dato) <= max(col_nombre, col_apellido):
                    continue

                if not dato[col_apellido].strip() or not dato[col_nombre].strip():
                    continue

                apellido_cell = dato[col_apellido].strip().lower()
                nombre_cell = dato[col_nombre].strip().lower()

                # Condicion de coincidencia segun los campos llenos
                coincide_apellido = apellido_buscar in apellido_cell if apellido_buscar else True
                coincide_nombre = nombre_buscar in nombre_cell if nombre_buscar else True

                if coincide_apellido and coincide_nombre:
                    encontrado = True
                    campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                    coincidencias.append(campos_guardar)
                    contador += 1

                    dato_formateado = list(dato)
                    dato_formateado[1] = dato_formateado[1].capitalize()
                    dato_formateado[2] = dato_formateado[2].capitalize()

                    try:
                        mensaje_item = f"DNI: {dato_formateado[0]}\n{hoja['formatear'](dato_formateado)}\n{'-'* 40}"
                    except IndexError:
                        mensaje_item = f"(Fila encontrada en {nombre_interno} pero con columnas insuficientes)"

                    if contador == 1:
                        primer_mensaje = mensaje_item
                    else:
                        primer_mensaje += f"\n{mensaje_item}"

            if contador > 0:
                self._escribir_resultado(primer_mensaje)

        self.guardar_resultados(coincidencias)
        
        if not encontrado:
            busqueda_str = f"{self.apellido_var.get().strip()} {self.nombre_var.get().strip()}".strip()
            self._escribir_resultado(
                f"El registro '{busqueda_str}' no se encontro en las bases de datos."
            )     
    
    def limpiar_campos(self, evitar):
        """
        Limpia todos los campos, evita el campo ingresado por parametro.
        Valores: "dni", "nombre", "domicilio"
        """
        if (not evitar == "dni"):
            self.dni_var.set("")
        if (not evitar == "nombre"):
            self.apellido_var.set("")
            self.nombre_var.set("")
        if (not evitar == "domicilio"):
            self.calle_var.set("")
            self.altura_var.set("")

    def buscar_domicilio(self):
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return

        self.limpiar_campos("domicilio")
        
        calle_buscar = self.calle_var.get().strip().lower()
        altura_buscar = self.altura_var.get().strip().lower()

        if not calle_buscar:
            messagebox.showwarning("Domicilio invalido", "Debe ingresar calle siempre (ej: Cnel. Pedro Aquino).")
            return

        self._escribir_resultado("", limpiar=True)
        encontrado = False
        coincidencias = []

        for hoja in SHEETS_CONFIG:
            col_calle = 3
            col_altura = 4
            nombre_interno = hoja["nombre_interno"]
            datos_hoja = self.registros.get(nombre_interno, [])
            primer_mensaje = ""
            contador = 0

            for dato in datos_hoja:
                if len(dato) <= col_calle:
                    continue

                if not dato[col_calle] or not dato[col_altura]:
                    continue

                if not dato[col_calle].replace(" ", "").isalnum() or not dato[col_altura].split()[0].isdigit():
                    continue

                calle = dato[col_calle].lower()
                altura = dato[col_altura]
                def hay_coincidencia(entrada, texto):
                    return re.search(re.escape(entrada.strip()), texto, re.IGNORECASE) is not None

                if hay_coincidencia(calle_buscar, calle):
                    if (altura_buscar and altura == altura_buscar):
                        encontrado = True
                        campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                        coincidencias.append(campos_guardar)
                        contador += 1
    
                        dato_formateado = list(dato)
                        dato_formateado[1] = dato_formateado[1].capitalize()
                        dato_formateado[2] = dato_formateado[2].capitalize()
    
                        try:
                            mensaje_item = f"DNI: {dato_formateado[0]}\n{hoja['formatear'](dato_formateado)}\n{'-' * 40}"
                        except IndexError:
                            mensaje_item = f"(Fila encontrada en {nombre_interno} pero con columnas insuficientes)"
    
                        if contador == 1:
                            primer_mensaje = mensaje_item
                        else:
                            primer_mensaje += f"\n{mensaje_item}"
                    elif (not altura_buscar):
                        encontrado = True
                        campos_guardar = [dato[0], dato[1], dato[2], dato[3], dato[4], dato[5], dato[6], hoja["Coordinacion"] , hoja["Programa"]]
                        coincidencias.append(campos_guardar)
                        contador += 1

                        dato_formateado = list(dato)
                        dato_formateado[1] = dato_formateado[1].capitalize()
                        dato_formateado[2] = dato_formateado[2].capitalize()

                        try:
                            mensaje_item = f"DNI: {dato_formateado[0]}\n{hoja['formatear'](dato_formateado)}\n{'-' * 40}"
                        except IndexError:
                            mensaje_item = f"(Fila encontrada en {nombre_interno} pero con columnas insuficientes)"

                        if contador == 1:
                            primer_mensaje = mensaje_item
                        else:
                            primer_mensaje += f"\n{mensaje_item}"

            if contador > 0:
                self._escribir_resultado(primer_mensaje)
        
        self.guardar_resultados(coincidencias)

        if not encontrado:
            self._escribir_resultado(
                f"El Domicilio ingresado '{calle_buscar.capitalize()}{" " + altura_buscar.capitalize() if altura_buscar else ""}' no se encontro en las bases de datos."
            )    

    # ------------------------------------------------------------------
    # Funciones de los resultados
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
        carpeta_descargas = r"C:\Reportes IMDEL"
        if not os.path.exists(carpeta_descargas):
            os.makedirs(carpeta_descargas)
        fecha = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        nombre_archivo = f"Registros de busqueda de beneficiarios - {fecha}hs.csv"

        ruta_archivo = os.path.join(carpeta_descargas, nombre_archivo)
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
    agrandar_fuentes(root, 12)
    BuscadorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()