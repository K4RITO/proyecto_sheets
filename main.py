import gspread

dni_buscar = input("Ingrese DNI de la persona que quiere buscar: ")

while not dni_buscar.isdigit():
    print("DNI invalido. Ingrese solo numeros (12345678).")
    dni_buscar = input("Ingrese DNI de la persona que quiere buscar: ")

encontrado = False

# Conectarse a Google Sheets
print("Conectando a Google Sheets...")
gc = gspread.service_account(filename="proyecto-sheets-505320-bb8ae069f096.json")

proyecto_sheets = gc.open("Tablas IMDEL Prototipo")

print("Cargando hojas...")
# Abrir todas las tablas en simultaneo
oficina_empleo = proyecto_sheets.get_worksheet(0)
desarrollo_agrario = proyecto_sheets.get_worksheet(1)
registro_recepcion = proyecto_sheets.get_worksheet(2)
inscripcion_cuatrimestral = proyecto_sheets.get_worksheet(3)
formulario_inscripcion = proyecto_sheets.get_worksheet(4)
nominalizacion = proyecto_sheets.get_worksheet(5)
usuarios = proyecto_sheets.get_worksheet(6)

print("Cargando datos de las hojas...")
contador = 0
# Obtener todos los datos
registros_oficina_empleo = oficina_empleo.get_all_values()
registros_desarrollo_agrario = desarrollo_agrario.get_all_values()
registros_registro_recepcion = registro_recepcion.get_all_values()
registros_inscripcion_cuatrimestral = inscripcion_cuatrimestral.get_all_values()
registros_formulario_inscripcion = formulario_inscripcion.get_all_values()
registros_nominalizacion = nominalizacion.get_all_values()
registros_usuarios = usuarios.get_all_values()

while True:
    print("TABLA OFICINA EMPLEO")

    for dato in registros_oficina_empleo:
        if dato[2] == dni_buscar:
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[1]} \nPrograma: Oficina empleo")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")
    

    print("TABLA DESARROLLO AGRARIO")
    contador = 0

    for dato in registros_desarrollo_agrario:
        if dato[2] == dni_buscar:
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[1]} {dato[0]} \nCoordinacion: Desarrollo agrario \nPrograma: Huertas familiares")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    print("TABLA RECEPCION")
    contador = 0

    for dato in registros_registro_recepcion:
        if dato[6] == dni_buscar:
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[3]} {dato[2]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Fortalecimiento de Trayectorias Laborales")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    print("TABLA INSCRIPCION CUATRIMESTRAL")
    contador = 0

    for dato in registros_inscripcion_cuatrimestral:
        if dato[5] == dni_buscar:
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[3]} {dato[2]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Capacitacion laboral")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    print("TABLA FORMULARIO INSCRIPCION")
    contador = 0

    for dato in registros_formulario_inscripcion:
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
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[37]} {dato[38]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Insercion laboral")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    print("TABLA NOMINALIZACION")
    contador = 0

    for dato in registros_nominalizacion:
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
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[7]} {dato[8]} \nCoordinacion: Capacitacion laboral y empleo \nPrograma: Plan fines")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    print("TABLA USUARIOS")
    contador = 0

    for dato in registros_usuarios:
        if len(dato[5]) >= 9:
            dni = []
            for caracter in dato[5]:
                if caracter.isdigit() == True:
                    dni.append(caracter)
                result_dni = "".join(dni)
                dato[5] = result_dni
                
        if dato[5] == dni_buscar:
            encontrado = True
            contador += 1
            if contador == 1:
                print(f"Apellido y nombre: {dato[2]} {dato[1]} \nCoordinacion: Economia popular \nPrograma: Registro de Trabajadores de la Economía Popular")
    if contador > 1:
        print(f"Se encontro al beneficiario {contador} veces")

    if (not encontrado): print(f"El DNI ingresado {dni_buscar} no se encontro en las bases de datos.")

    dni_buscar = input("Ingrese DNI de la persona que quiere buscar: ")

    while not dni_buscar.isdigit():
        print("DNI invalido. Ingrese solo numeros (12345678).")
        dni_buscar = input("Ingrese DNI de la persona que quiere buscar: ")

    encontrado = False

    if (dni_buscar == "0"): break