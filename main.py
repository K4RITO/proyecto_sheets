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
import ast
import glob
import shutil
from pathlib import Path

from PIL import Image, ImageTk
import gspread

# Configuracion de conexion
CREDENTIALS_FILE = "Credenciales.json"
SHEET_NAME = "Tablas IMDEL"

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

def vaciar_carpeta(ruta):
    carpeta = Path(ruta)
    for item in carpeta.iterdir():
        if item.is_dir() and not item.is_symlink():
            shutil.rmtree(item)   # borra subcarpetas con todo su contenido
        else:
            item.unlink()  

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
COLUMNAS = ("DNI", "Nombre", "Apellido", "Calle", "Altura",
            "Teléfono", "Email", "Coordinación", "Programa") # Columnas de la tabla de resultados
BACKUP_DIR = r"C:\Reportes IMDEL\backup"
PASO_SCROLL_H = 40  # pixeles por "click" de la rueda; subilo o bajalo a gusto

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
        self.gc = None # Conexion con Google Cloud a traves de GSpread
        self.proyecto_sheets = None 
        self.worksheets = {}   # nombre_interno -> objeto worksheet
        self.registros = {}    # nombre_interno -> lista de filas (get_all_values)
        self.conectado = False
        self.modo_offline = False

        # Variables del menu
        self.tamano_var = tk.IntVar(value=12) # Tamaño de la fuente, por defecto 12pt
        self.tema_oscuro = tk.BooleanVar(value=False) # Variable que determina si tema es oscuro
        self.tema_original = ttk.Style(root).theme_use() # Variable que guarda el tema de la aplicacion por defecto.
        self.vista_var = tk.StringVar(value="texto")  # Determina si los resultados aparecen como "Texto" o "Tabla"

        # Arma la interfaz general
        self._armar_interfaz()

        # Construye el menu
        self._armar_menu()

        # Conecta en otro hilo al iniciar para no congelar la UI
        self._conectar_en_hilo()

    # ------------------------------------------------------------------
    # Construccion de la interfaz
    # ------------------------------------------------------------------

    def _armar_interfaz(self):
        """Construye todos los widgets de la ventana principal.

        Secciones, de arriba hacia abajo:
        - Encabezado: icono de la ventana (logo.ico), logo PNG (logo.png.png)
          y etiqueta de estado de conexión (self.estado_var).
        - Grilla de búsqueda: campos y botones para buscar por DNI, por
          apellido/nombre y por calle/altura. Cada campo tiene validación en
          tiempo real y dispara su búsqueda con Enter. El botón de domicilio
          permanece deshabilitado hasta que se ingresa una calle.
        - Área de resultados: dos vistas intercambiables, texto
          (self.resultado_text) y tabla (self.tabla), con sus scrollbars.
          La tabla permite scroll horizontal con Shift + rueda.
        - Etiqueta de mensajes (self.msg_var) debajo de los resultados.

        Al final llama a cambiar_vista() para mostrar la vista seleccionada
        en self.vista_var. Requiere que self.tamano_var y self.vista_var ya
        existan, por lo que debe llamarse después de crearlas en __init__.
        """

        # 1. Establecer el icono de la ventana (.ico)
        try:
            self.root.iconbitmap("logo.ico")
        except Exception as e:
            print(f"No se pudo cargar logo.ico: {e}")

        contenedor = ttk.Frame(self.root, padding=12)
        contenedor.pack(fill="both", expand=True)

        # --- Encabezado: Logo PNG + Estado de conexión ---
        header_frame = ttk.Frame(contenedor)
        header_frame.pack(fill="x", pady=(0, 8))

        # 2. Cargar e insertar el logo PNG en el diseño
        try:
            img_logo = Image.open("logo.png.png")
            # Ajusta (ancho, alto) según lo que necesites para tu interfaz
            img_logo = img_logo.resize((80, 80), Image.Resampling.LANCZOS)
            self.logo_tk = ImageTk.PhotoImage(img_logo)

            lbl_logo = ttk.Label(header_frame, image=self.logo_tk)
            lbl_logo.pack(side="left", padx=(0, 10))
        except Exception as e:
            print(f"No se pudo cargar el logo.png.png: {e}")
        # Para el icono de la ventana:
        self.root.iconbitmap("logo.ico")

        # Para el logo de la interfaz:
        img_logo = Image.open("logo.png.png")
        # --- Estado de conexión ---
        self.estado_var = tk.StringVar(value="Conectando a Google Sheets...")
        estado_label = ttk.Label(header_frame, textvariable=self.estado_var, foreground="gray")
        estado_label.pack(side="left", anchor="w")

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
        self.marco_resultados = ttk.Frame(contenedor)
        self.marco_resultados.pack(fill="both", expand=True, pady=(4, 0))

        # Vista texto
        self.marco_texto = ttk.Frame(self.marco_resultados)
        self.resultado_text = scrolledtext.ScrolledText(
            self.marco_texto, wrap="word", height=20, state="disabled",
            font=("Consolas", self.tamano_var.get()),
        )
        self.resultado_text.pack(fill="both", expand=True)

        # Vista tabla
        self.marco_tabla = ttk.Frame(self.marco_resultados)
        self.tabla = ttk.Treeview(self.marco_tabla, columns=COLUMNAS, show="headings")
        for col, ancho in zip(COLUMNAS, (90, 100, 100, 140, 60, 110, 160, 170, 230)):
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=ancho, minwidth=50, stretch=False)
        sy = ttk.Scrollbar(self.marco_tabla, orient="vertical", command=self.tabla.yview)
        sx = ttk.Scrollbar(self.marco_tabla, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        self.marco_tabla.rowconfigure(0, weight=1)
        self.marco_tabla.columnconfigure(0, weight=1)

        self.tabla.bind("<Shift-MouseWheel>", self._scroll_horizontal)

        self.msg_var = tk.StringVar()
        ttk.Label(contenedor, textvariable=self.msg_var).pack(anchor="w", pady=(4, 0))

        self.cambiar_vista()
        # # --- Botones inferiores ---
        # self.resultados_guardados = ""
        # fila_inferior = ttk.Frame(contenedor)
        # fila_inferior.pack(fill="x", pady=(8, 0))

        # self.btn_descargar_resultados = ttk.Button(
        #     fila_inferior, text="Descargar resultados", command=self.descargar_resultados
        # )
        # self.btn_descargar_resultados.pack(side="left")

        # self.btn_actualizar = ttk.Button(
        #     fila_inferior, text="Actualizar bases de datos", command=self.actualizar_bases
        # )
        # self.btn_actualizar.pack(side="right")

    # ------------------------------------------------------------------
    # Menu
    # ------------------------------------------------------------------

    def _armar_menu(self):
        """Crea la barra de menú superior de la ventana.

        Menús:
        - Archivo: descargar resultados, actualizar bases de datos y salir.
        - Ver: tamaño de letra (10 a 18 pt), modo oscuro y vista de
          resultados (Texto o Tabla).
        - Ayuda: ventana "Acerca de".

        Las opciones están enlazadas a self.tamano_var, self.tema_oscuro y
        self.vista_var. Para agregar funciones nuevas, basta sumar entradas
        al menú correspondiente.
        """

        barra = tk.Menu(self.root)
        self.root.config(menu=barra)

        # Sub menu de la seccion "Archivo"
        m_archivo = tk.Menu(barra, tearoff=0)
        m_archivo.add_command(label="Descargar resultados", command=self.descargar_resultados)
        m_archivo.add_command(label="Actualizar bases de datos", command=self.actualizar_bases)
        m_archivo.add_separator()
        m_archivo.add_command(label="Salir", command=self.root.destroy)
        barra.add_cascade(label="Archivo", menu=m_archivo)

        # Sub Menu de la seccion "Ver"
        m_ver = tk.Menu(barra, tearoff=0)
        m_fuente = tk.Menu(m_ver, tearoff=0)
        for t in (10, 12, 14, 16, 18):
            m_fuente.add_radiobutton(
                label=f"{t} pt", variable=self.tamano_var, value=t,
                command=lambda: self.cambiar_fuente(self.tamano_var.get()),
            )
        m_ver.add_cascade(label="Tamaño de letra", menu=m_fuente)
        m_ver.add_checkbutton(label="Modo oscuro", variable=self.tema_oscuro,
                              command=self.aplicar_tema)
        
        # Para configurar los resultados como Texto o Tabla
        m_vista = tk.Menu(m_ver, tearoff=0)
        m_vista.add_radiobutton(label="Texto", variable=self.vista_var,
                                value="texto", command=self.cambiar_vista)
        m_vista.add_radiobutton(label="Tabla", variable=self.vista_var,
                                value="tabla", command=self.cambiar_vista)
        m_ver.add_cascade(label="Vista de resultados", menu=m_vista)
        barra.add_cascade(label="Ver", menu=m_ver)

        # Ayuda (para futuras opciones)
        m_ayuda = tk.Menu(barra, tearoff=0)
        m_ayuda.add_command(
            label="Acerca de",
            command=lambda: messagebox.showinfo("Acerca de", "Buscador de beneficiarios - IMDEL"),
        )
        barra.add_cascade(label="Ayuda", menu=m_ayuda)

    def cambiar_vista(self):
        """Cambia el modo de vista de resultados, alterna entre vista como lista y como Tabla"""
        self.marco_texto.pack_forget()
        self.marco_tabla.pack_forget()
        if self.vista_var.get() == "tabla":
            self.marco_tabla.pack(fill="both", expand=True)
        else:
            self.marco_texto.pack(fill="both", expand=True)

    # Utilidades

    def _escribir_resultado(self, texto, limpiar=False):
        """Guarda los resultados para que puedan aparecer tanto en el Textarea como en la tabla"""
        if limpiar:
            self._limpiar_resultados()
        else:
            self.msg_var.set(texto)

    def _scroll_horizontal(self, event):
        """Configura la velocidad del Scroll Horizontal"""
        pasos = -event.delta / 120 * PASO_SCROLL_H
        if pasos:
            self.tabla.xview_scroll(int(pasos) or (1 if pasos > 0 else -1), "units")
        return "break"

    def _buscar(self, coincide, no_encontrado):
        """Realiza una busqueda entre las bases cargadas en memoria y guarda el resultado de esta\n
        Primero Limpia el campo de resultados y luego realiza la busqueda para guardar el resultado y mostrarlo en pantalla
        """
        self._limpiar_resultados()
        
        filas = []
        self.resultado_text.config(state="normal")

        for hoja in SHEETS_CONFIG:
            for dato in self.registros.get(hoja["nombre_interno"], []):
                if len(dato) < 7 or not coincide(hoja, dato):
                    continue
                fila = [*dato[:7], hoja["Coordinacion"], hoja["Programa"]]
                filas.append(fila)
                self.tabla.insert("", "end", values=fila)

                d = list(dato)
                d[1], d[2] = d[1].capitalize(), d[2].capitalize()
                self.resultado_text.insert(
                    "end", f"DNI: {d[0]}\n{hoja['formatear'](d)}\n{'-' * 40}\n")
        self.resultado_text.config(state="disabled")
        self.guardar_resultados(filas)
        self._escribir_resultado(f"{len(filas)} resultado(s)." if filas else no_encontrado)
    
    def _limpiar_resultados(self):
        """Limpia el campo de resultados"""
        self.tabla.delete(*self.tabla.get_children())
        self.resultado_text.config(state="normal")
        self.resultado_text.delete("1.0", "end")
        self.resultado_text.config(state="disabled")
        self.msg_var.set("")
        self.resultados_guardados = []

    def cambiar_fuente(self, tamano):
        """Cambia el tamaño de la fuente de toda la aplicacion"""
        agrandar_fuentes(self.root, tamano)
        self.resultado_text.config(font=("Consolas", tamano))

    def aplicar_tema(self):
        """Alterna los temas de la aplicacion, actualmente existen los temas 'claro' y 'oscuro' """
        estilo = ttk.Style(self.root)

        if self.tema_oscuro.get():
            bg, fg, campo = "#2b2b2b", "#e6e6e6", "#3c3f41"
            estilo.theme_use("clam")
            estilo.configure(".", background=bg, foreground=fg)
            estilo.configure("TEntry", fieldbackground=campo, foreground=fg)
            estilo.configure("Treeview", background=campo, fieldbackground=campo, foreground=fg)
            estilo.configure("Treeview.Heading", background="#4a4d50", foreground=fg)
            estilo.map("Treeview", background=[("selected", "#2f65ca")])
            estilo.map(".", foreground=[("disabled", "#888888")])
            estilo.map("TButton", background=[("active", "#4a4d50")])
        else:
            bg, fg, campo = "#f0f0f0", "black", "white"
            estilo.theme_use(self.tema_original)

        self.root.configure(bg=bg)
        self.resultado_text.config(bg=campo, fg=fg, insertbackground=fg)

        agrandar_fuentes(self.root, self.tamano_var.get())
    # --- Validaciones ---

    def _validar_entrada_dni(self, valor_propuesto):
        """Valida el valor ingresado en el campo de entrada de DNI"""
        return valor_propuesto == "" or valor_propuesto.isdigit()

    def _validar_entrada_nombre(self, valor_propuesto):
        """Valida el valor ingresado en el campo de entrada de Nombre y de Apellido"""
        return valor_propuesto == "" or valor_propuesto.replace(" ", "").isalpha()
    
    def _validar_entrada_calle(self, texto):
        """Valida el valor ingresado en el campo de entrada de Calle"""
        # Letras (con acentos y ñ), números, espacios y puntos. Máx. 40 caracteres.
        if texto == "":
            return True
        return (
            len(texto) <= 40
            and re.fullmatch(r"[A-Za-zÁÉÍÓÚáéíóúÑñÜü0-9. ]+", texto) is not None
        )

    def _validar_entrada_altura(self, texto):
        """Valida el valor ingresado en el campo de entrada de Altura"""
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
        """Habilita el campo 'Altura' solo si hay  """
        if self.calle_var.get().strip():
            self.btn_buscar_domicilio.state(["!disabled"])
        else:
            self.btn_buscar_domicilio.state(["disabled"])


    # ------------------------------------------------------------------
    # Conexion y carga de datos
    # ------------------------------------------------------------------
    def _conectar_en_hilo(self):
        """Inicia la conexión a Google Sheets en un hilo secundario.

        Deshabilita los controles de la interfaz mientras este conectando
        para que la ventana no se congele ni el usuario busque sin datos.
        """
        self._set_controles_habilitados(False)
        hilo = threading.Thread(target=self._conectar, daemon=True)
        hilo.start()

    def _conectar(self):
        """Conecta con Google Sheets, carga las hojas y guarda un backup local.

        Flujo:
        - Con conexión: autentica con CREDENTIALS_FILE, abre la planilla
          SHEET_NAME, obtiene las hojas definidas en SHEETS_CONFIG, carga sus
          datos en self.registros y genera un backup en BACKUP_DIR.
        - Sin conexión (cualquier error): intenta cargar el backup más
          reciente y activa el modo offline. Si no hay backup, muestra el
          error de conexión.

        Siempre rehabilita los controles al terminar. Se ejecuta en un hilo
        secundario, por lo que las actualizaciones de la interfaz se
        delegan al hilo principal.
        """
        try:
            self.gc = gspread.service_account(filename=CREDENTIALS_FILE)
            self.proyecto_sheets = self.gc.open(SHEET_NAME)
            for hoja in SHEETS_CONFIG:
                ws = self.proyecto_sheets.get_worksheet(hoja["worksheet_index"])
                self.worksheets[hoja["nombre_interno"]] = ws

            self._cargar_datos()
            # raise Exception("Prueba") 
            self.guardar_backup(self.registros)
            self.conectado = True
            self.modo_offline = False
            self.root.after(0, lambda: self.estado_var.set("Conectado. Bases de datos cargadas."))
        except Exception as e:
            try:
                archivo = self._cargar_backup()
            except Exception:
                archivo = None
            if archivo:
                self.conectado = True
                self.modo_offline = True
                self.root.after(0, lambda: self.estado_var.set(
                    f"SIN CONEXIÓN. Usando backup: {archivo}"))
            else:
                self.root.after(0, lambda err=e: self._mostrar_error_conexion(err))
        finally:
            self.root.after(0, lambda: self._set_controles_habilitados(True))

    def _cargar_datos(self):
        """Carga los datos de cada hoja del Sheet en los self.registros como un diccionario
        
        Las claves de self.registros es el nombre interno de cada hoja y el valor de este es todo su contenido.
        """
        for nombre_interno, ws in self.worksheets.items():
            self.registros[nombre_interno] = ws.get_all_values()

    def _mostrar_error_conexion(self, error):
        """Muestra un error al fallar la conexion con Google Cloud"""
        self.estado_var.set("Error al conectar. Ver detalle.")
        messagebox.showerror(
            "Error de conexion",
            f"No se pudo conectar a Google Sheets:\n{error}",
        )

    def _set_controles_habilitados(self, habilitados):
        """Deshabilita todos los controles de la aplicacion para evitar que el usuario utilice los campos y botones cuando estan estableciendo la conexion con Google Cloud"""
        estado = "normal" if habilitados else "disabled"
        self.btn_buscar.config(state=estado)
        self.dni_entry.config(state=estado)
        self.apellido_entry.config(state=estado)
        self.nombre_entry.config(state=estado)
        self.btn_buscar_nombre.config(state=estado)
        self.calle_entry.config(state=estado)
        self.altura_entry.config(state=estado)
        self.btn_buscar_domicilio.config(state=estado)

    # GUARDAR BACKUPS

    def guardar_backup(self, datos):
        """Guarda un backup local de todas las hojas en un único archivo CSV.

        Crea BACKUP_DIR si no existe; si existe, lo vacía antes de guardar,
        de modo que siempre quede un solo backup (el más reciente). El
        archivo se nombra "backup - dd-mm-YYYY_HH-MM-SShs.csv".

        Formato: cada fila del CSV corresponde a una hoja (en el orden de
        self.worksheets) y cada celda contiene el str() de una fila de esa
        hoja, que luego _cargar_backup reconstruye con ast.literal_eval.

        Args:
            datos (dict): Diccionario {nombre_interno: lista de filas},
                normalmente self.registros.

        Muestra un aviso con la ruta del archivo al finalizar.
        """
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
        else:
            vaciar_carpeta(BACKUP_DIR)

        fecha = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")       
        nombre_archivo = f"backup - {fecha}hs.csv"
        ruta_archivo = os.path.join(BACKUP_DIR, nombre_archivo)

        nombres_internos = []
        for nombre_interno, _ in self.worksheets.items():
            nombres_internos.append(nombre_interno)

        with open(ruta_archivo, "w", newline="", encoding="utf-8") as archivo:
            escritor = csv.writer(archivo)

            for i in range(len(nombres_internos)):
                escritor.writerow(datos[nombres_internos[i]])

        messagebox.showinfo(
            "Descarga completada",
            f"El archivo se guardó correctamente en:\n{ruta_archivo}"
        )
    
    def _cargar_backup(self):
        """Carga el backup más reciente de BACKUP_DIR en self.registros.

        Busca los archivos "backup - *.csv", toma el más
        reciente y reconstruye los registros de cada hoja según el
        orden de SHEETS_CONFIG, convirtiendo cada celda de una fila
        otra vez en lista.

        Returns:
            str | None: Nombre del archivo cargado, o None si no hay ningún backup disponible.
        """
        archivos = glob.glob(os.path.join(BACKUP_DIR, "backup - *.csv"))

        if not archivos:
            return None
        
        ultimo = max(archivos, key=os.path.getmtime)

        with open(ultimo, newline="", encoding="utf-8") as f:
            filas = list(csv.reader(f))
            
        # Cada fila del CSV = una hoja; cada celda = str() de una fila de la hoja
        for hoja, fila in zip(SHEETS_CONFIG, filas):
            self.registros[hoja["nombre_interno"]] = [ast.literal_eval(c) for c in fila]

        return os.path.basename(ultimo)
    # ------------------------------------------------------------------
    # Boton: actualizar bases
    # ------------------------------------------------------------------
    def actualizar_bases(self):
        """Recarga los datos de las hojas desde Google Sheets.

        - En modo offline: reintenta la conexión completa.
        - Si la conexión inicial aún no terminó: avisa al usuario y no hace nada.
        - En condiciones normales: deshabilita los controles y lanza la
          recarga en un hilo secundario para no congelar la interfaz.
        """
        if self.modo_offline:
            self.estado_var.set("Reintentando conexión...")
            self._conectar_en_hilo()
            return
        
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return
        
        self.estado_var.set("Actualizando bases de datos...")
        self._set_controles_habilitados(False)
        hilo = threading.Thread(target=self._actualizar_bases_hilo, daemon=True)
        hilo.start()

    def _actualizar_bases_hilo(self):
        """Vuelve a cargar los registros de todas las hojas (hilo secundario).

        Si la carga es exitosa, informa en la barra de estado; si falla,
        muestra el error de conexión. En ambos casos rehabilita los
        controles al terminar. Las actualizaciones de la interfaz se
        delegan al hilo principal.
        """
        try:
            self._cargar_datos()
            self.guardar_backup(self.registros)
            self.root.after(0, lambda: self.estado_var.set("Bases de datos actualizadas."))
        except Exception as e:
            self._cargar_backup()
            self.root.after(0, lambda err=e: self._mostrar_error_conexion(err))
        finally:
            self.root.after(0, lambda: self._set_controles_habilitados(True))

    # ------------------------------------------------------------------
    # Botones de busqueda
    # ------------------------------------------------------------------

    def buscar_dni(self):
        """Realiza la busqueda exacta por DNI en las bases cargadas en memoria
        
        Returns:
            str: Registro del beneficiario de cada hoja, la cantidad de apariciones por hoja o un mensaje indicando que no se lo encontro.
        """
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return
        self.limpiar_campos("dni")
        dni = self.dni_var.get().strip()
        if not dni or not dni.isdigit():
            messagebox.showwarning("DNI invalido", "Ingrese solo numeros (ej: 12345678).")
            return

        def coincide(hoja, dato):
            valor = hoja["limpiar"](dato[0], hoja["min_len"])
            dato[0] = valor
            return valor == dni

        self._buscar(coincide, f"El DNI ingresado {dni} no se encontro en las bases de datos.")

    def buscar_nombre(self):
        """Realiza una busqueda blanda por un texto ingresado en las bases cargadas en memoria
        
        Returns:
            str: Registros de los beneficiarios que contengan el patron dentro del nombre o apellido (depende del campo utilizado) de cada hoja o un mensaje indicando que no se lo encontro.
        """
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return
        self.limpiar_campos("nombre")
        apellido = self.apellido_var.get().strip().lower()
        nombre = self.nombre_var.get().strip().lower()
        if not apellido and not nombre:
            messagebox.showwarning("Entrada invalida", "Debe ingresar al menos un apellido o un nombre.")
            return

        def coincide(hoja, dato):
            n, a = dato[1].strip().lower(), dato[2].strip().lower()
            return bool(n and a) and nombre in n and apellido in a

        busqueda = f"{apellido} {nombre}".strip()
        self._buscar(coincide, f"El registro '{busqueda}' no se encontro en las bases de datos.")

    def buscar_domicilio(self):
        """Realiza una busqueda blanda por un texto ingresado en las bases cargadas en memoria
        
        Returns:
            str: Registros de los beneficiarios que contengan el patron dentro del nombre o apellido (depende del campo utilizado) de cada hoja o un mensaje indicando que no se lo encontro.
        """
        if not self.conectado:
            messagebox.showwarning("Sin conexion", "Todavia no se completo la conexion inicial.")
            return
        self.limpiar_campos("domicilio")
        calle_buscar = self.calle_var.get().strip().lower()
        altura_buscar = self.altura_var.get().strip()
        if not calle_buscar:
            messagebox.showwarning("Domicilio invalido", "Debe ingresar calle siempre (ej: Cnel. Pedro Aquino).")
            return

        def coincide(hoja, dato):
            calle, altura = dato[3], dato[4]
            if not calle or not altura.split():
                return False
            if not calle.replace(" ", "").isalnum() or not altura.split()[0].isdigit():
                return False
            if calle_buscar not in calle.lower():
                return False
            return not altura_buscar or altura == altura_buscar

        domicilio = f"{calle_buscar.capitalize()} {altura_buscar}".strip()
        self._buscar(coincide, f"El Domicilio '{domicilio}' no se encontro en las bases de datos.")    
    
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

    # ------------------------------------------------------------------
    # Funciones de los resultados
    # ------------------------------------------------------------------

    def guardar_resultados(self, datos):
        """Guarda los resultados pasados por parametros"""
        self.resultados_guardados = datos

    def descargar_resultados(self):
        """Exporta los resultados de la última búsqueda a un archivo CSV.

        Crea la carpeta C:\\Reportes IMDEL si no existe y guarda allí el
        archivo "Registros de busqueda de beneficiarios - dd-mm-YYYY_HH-MM-SShs.csv".
        La primera fila contiene los encabezados (DNI, Nombre, Apellido, Calle,
        Altura, Numero Telefono, Email, Coordinacion, Programa) y las
        siguientes, los registros guardados en self.resultados_guardados.

        Al finalizar muestra un aviso con la ruta del archivo generado.
        """

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