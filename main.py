import gspread

# Conectarse a Google Sheets
gc = gspread.service_account(filename="proyecto-sheets-505320-bb8ae069f096.json")

# Abrir el archivo de Google Sheets
oficina_empleo = gc.open("Tablas IMDEL Prototipo").get_worksheet(0)
contador = 0
# Obtener todos los datos
datos = oficina_empleo.get_all_values()

"""Validar entrada de datos dni. 
    Agregar un bucle para ejecucion continua
    separar la carga de datos con el funcionamiento del backend(que solo cargue los datos la primera vez)
"""
dni_buscar = input("Ingrese DNI de la persona que quiere buscar: ")

for dato in datos:
    if dato[2] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[1]} \nPrograma: Oficina empleo")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")
 

desarrollo_agrario = gc.open("Tablas IMDEL Prototipo").get_worksheet(1)
contador = 0
datos = desarrollo_agrario.get_all_values()

for dato in datos:
    if dato[2] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[1]} {dato[0]} \nCoordinacion: Desarrollo agrario \nPrograma: Huertas familiares")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")

registro_recepcion = gc.open("Tablas IMDEL Prototipo").get_worksheet(2)
contador = 0
datos = registro_recepcion.get_all_values()

for dato in datos:
    if dato[6] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[3]} {dato[2]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Fortalecimiento de Trayectorias Laborales")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")

inscripcion_cuatrimestral = gc.open("Tablas IMDEL Prototipo").get_worksheet(3)
contador = 0
datos = inscripcion_cuatrimestral.get_all_values()

for dato in datos:
    if dato[5] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[3]} {dato[2]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Capacitacion laboral")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")

formulario_inscripcion = gc.open("Tablas IMDEL Prototipo").get_worksheet(4)
contador = 0
datos = formulario_inscripcion.get_all_values()

for dato in datos:
    if len(dato[39]) >= 10:
        dni = []
        for caracter in dato[39]:
            if caracter.isdigit() == True:
                dni.append(caracter)
        dni.pop(0)
        dni.pop(0)
        dni.pop(-1)
        result_dni = "".join(dni)
        dato[39] = result_dni

    if dato[39] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[37]} {dato[38]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Insercion laboral")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")

nominalizacion = gc.open("Tablas IMDEL Prototipo").get_worksheet(5)
contador = 0
datos = nominalizacion.get_all_values()

for dato in datos:
    if len(dato[9]) >= 10:
        dni = []
        for caracter in dato[9]:
            if caracter.isdigit() == True:
                dni.append(caracter)
        dni.pop(0)
        dni.pop(0)
        dni.pop(-1)
        result_dni = "".join(dni)
        dato[9] = result_dni

    if dato[9] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[7]} {dato[8]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Plan fines")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")

usuarios = gc.open("Tablas IMDEL Prototipo").get_worksheet(6)
contador = 0
datos = usuarios.get_all_values()

for dato in datos:
    if len(dato[5]) >= 9:
        dni = []
        for caracter in dato[5]:
            if caracter.isdigit() == True:
                dni.append(caracter)
            result_dni = "".join(dni)
            dato[5] = result_dni
            
    if dato[5] == dni_buscar:
        contador += 1
        if contador == 1:
            print(f"Apellido y nombre: {dato[2]} {dato[1]} \nCoordinacion: Economia popular \nPrograma: Registro de Trabajadores de la Economía Popular")
if contador > 1:
    print(f"Se encontro al beneficiario {contador} veces")