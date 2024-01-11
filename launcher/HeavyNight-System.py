import os
import re
import requests
import sys
import subprocess
import glob
import shutil
import webbrowser
import tkinter as tk
import winsound
from tkinter import ttk


ini_file_path = "data/HeavyNight.ini"
local_file_path = "data/categorias.txt"
launchcfg_path = "data/launchcfg"

####################### FUNCIONES NECESARIAS #############################
def popup_message(title, message, icon_path, button_color):
    # Crear una nueva ventana
    popup = tk.Tk()
    popup.withdraw()  # Ocultar temporalmente la ventana para evitar que se muestre en una posición no deseada
    popup.title(title)
    popup.iconbitmap('data/icono.ico')  # Asegúrate de que la ruta al icono sea correcta

    # Establecer la geometría de la ventana y evitar que se redimensione
    #popup.geometry('400x200')  # Ancho x Alto
    popup.resizable(0, 0)

    # Estilo personalizado para el botón
    style = ttk.Style()
    style.configure('TButton', font=('Calibri', 10), background=button_color)

    # Contenido de la ventana
    content_frame = tk.Frame(popup, bg='white', padx=10, pady=10)
    content_frame.pack(expand=True, fill='both')

    # Ícono en el lado izquierdo
    icon = tk.PhotoImage(file=icon_path)  # Utiliza el argumento icon_path para la imagen
    icon_label = tk.Label(content_frame, image=icon, bg='white')
    icon_label.image = icon  # Referencia para evitar que la imagen sea recolectada por el recolector de basura
    icon_label.pack(side='left', padx=(10, 20))

    # Mensaje de texto
    message_label = tk.Label(content_frame, text=message, bg='white', fg='black', justify='left', anchor='w')
    message_label.pack(side='left', expand=True)

    # Botón "Aceptar" para cerrar la ventana emergente
    button_close = ttk.Button(popup, text="Aceptar", style='TButton', command=popup.destroy)
    button_close.pack(side='bottom', pady=10)

    # Centrar la ventana en la pantalla
    popup.update_idletasks()  # Actualiza las tareas pendientes de Tkinter para obtener las dimensiones correctas
    width = popup.winfo_reqwidth() + 20  # Agrega 20 píxeles de ancho adicional
    height = popup.winfo_reqheight() + 20  # Agrega 20 píxeles de alto adicional
    x = (popup.winfo_screenwidth() // 2) - (width // 2)
    y = (popup.winfo_screenheight() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")

    # Reproducir sonido de alerta
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

    # Mostrar la ventana después de configurar todo
    popup.deiconify()

    popup.mainloop()

def popup_question(title, message, icon_path, button_color):
    # Crear una nueva ventana
    popup = tk.Tk()
    popup.withdraw()  # Ocultar temporalmente la ventana para evitar que se muestre en una posición no deseada
    popup.title(title)
    popup.iconbitmap('data/icono.ico')  # Asegúrate de que la ruta al icono sea correcta

    # Establecer la geometría de la ventana y evitar que se redimensione
    #popup.geometry('400x100')  # Ancho x Alto
    popup.resizable(0, 0)

    # Estilo personalizado para el botón
    style = ttk.Style()
    style.configure('TButton', font=('Calibri', 10), background=button_color)

    # Variable para almacenar la respuesta del usuario
    user_response = tk.BooleanVar(value=False)

    # Funciones para manejar la respuesta del usuario
    def yes_action():
        user_response.set(True)
        popup.destroy()

    def no_action():
        user_response.set(False)
        popup.destroy()

    # Contenido de la ventana
    content_frame = tk.Frame(popup, bg='white', padx=10, pady=10)
    content_frame.pack(expand=True, fill='both')

    # Ícono en el lado izquierdo
    icon = tk.PhotoImage(file=icon_path)  # Utiliza el argumento icon_path para la imagen
    icon_label = tk.Label(content_frame, image=icon, bg='white')
    icon_label.image = icon  # Referencia para evitar que la imagen sea recolectada por el recolector de basura
    icon_label.pack(side='left', padx=(10, 20))

    # Mensaje de texto
    message_label = tk.Label(content_frame, text=message, bg='white', fg='black', justify='left', anchor='nw', wraplength=300)
    message_label.pack(side='left', expand=True, fill='both')

    # Botones "Sí" y "No" para la respuesta del usuario
    button_yes = ttk.Button(popup, text="Sí", style='TButton', command=yes_action)
    button_no = ttk.Button(popup, text="No", style='TButton', command=no_action)
    button_yes.pack(side='left', padx=10, pady=10)
    button_no.pack(side='right', padx=10, pady=10)

    # Centrar la ventana en la pantalla
    popup.update_idletasks()  # Actualiza las tareas pendientes de Tkinter para obtener las dimensiones correctas
    width = popup.winfo_reqwidth() + 20  # Agrega 20 píxeles de ancho adicional
    height = popup.winfo_reqheight() + 20  # Agrega 20 píxeles de alto adicional
    x = (popup.winfo_screenwidth() // 2) - (width // 2)
    y = (popup.winfo_screenheight() // 2) - (height // 2)
    popup.geometry(f"{width}x{height}+{x}+{y}")

    # Reproducir sonido de alerta
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

    # Mostrar la ventana después de configurar todo
    popup.deiconify()

    # Ejecutar el bucle principal de Tkinter y esperar a que la ventana se cierre
    popup.mainloop()

    # Después de cerrar la ventana, devolver la respuesta del usuario
    return user_response.get()

def ejecutar_seccion(seccion):
    if seccion in SECCIONES_CON_LOCK:
        if not check_lock_file(seccion):
            create_lock_file(seccion)
            try:
                ejecutar_seccion_con_manejo_de_errores(seccion)
            finally:
                delete_lock_file(seccion)
        else:
            print(f"La sección {seccion} ya está en ejecución.")
    else:
        # Si la sección no está en la lista, simplemente la ejecuta
        ejecutar_seccion_con_manejo_de_errores(seccion)

SECCIONES_CON_LOCK = ["desinstala", "mantenimiento", "iniciarc1", "parcharc1", "desinstalarc1", "iniciarc2", "parcharc2", "desinstalarc2", "iniciarc3", "parcharc3", "desinstalarc3", "iniciarc4", "parcharc4", "desinstalarc4"]  # Agrega todas las secciones que quieras aquí

def check_lock_file(seccion):
    lock_file = os.path.join(get_data_folder(), f"Lock_{seccion}.lock")
    return os.path.exists(lock_file)

def get_data_folder():
    script_folder = get_script_folder()
    return os.path.abspath(os.path.join(script_folder, "../data"))

def get_script_folder():
    script_path = sys.argv[0]
    return os.path.dirname(script_path)

def create_lock_file(seccion):
    data_folder = get_data_folder()
    lock_file = os.path.join(data_folder, f"Lock_{seccion}.lock")
    
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
    
    with open(lock_file, 'w') as file:
        file.write('')  # Crea el archivo vacío

def delete_lock_file(seccion):
    lock_file = os.path.join(get_data_folder(), f"Lock_{seccion}.lock")
    
    if os.path.exists(lock_file):
        os.remove(lock_file)

def perform_replacements_in_memory(file_path, replacements):
    with open(file_path, 'r', encoding='utf-8') as file:
        file_content = file.read()

    for search_str, replace_str in replacements:
        file_content = file_content.replace(search_str, replace_str)
    
    return file_content

def edit_launch_cfg_multiple_replacements(file_path, replacements):
    new_content = perform_replacements_in_memory(file_path, replacements)
    
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(new_content)

def extract_text_between_brackets(string):
    match = re.search(r'\[(.*?)\]', string)
    return match.group(1) if match else ""

def read_ini_into_dict(ini_file_path):
    with open(ini_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    data_dict = {}
    for line in lines:
        if "=" in line:
            key, value = line.split("=", 1)
            data_dict[key.strip()] = value.strip()
    return data_dict

def read_ini(ini_file_path):
    ini_data = {}
    current_section = None
    with open(ini_file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if line.startswith('[') and line.endswith(']'):
                current_section = line[1:-1]
                ini_data[current_section] = {}
            elif '=' in line and current_section:
                key, value = line.split('=', 1)
                ini_data[current_section][key] = value
    return ini_data

def edit_launchcfg_based_on_ini(ini_data, launchcfg_file_path):
    with open(launchcfg_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # Preparar las líneas modificadas
    modified_lines = []

    for line in lines:
        line_stripped = line.strip()
        
        # Verificar si la línea es un comentario o una clave
        if line_stripped.startswith("//[") or line_stripped.startswith("["):
            # Extraer la clave sin corchetes
            ini_key = line_stripped.replace("//", "").strip("[]")

            # Comparar con el estado en el archivo INI
            if ini_key in ini_data['Botones']:
                if ini_data['Botones'][ini_key] == "off":
                    modified_lines.append(f"//[{ini_key}]\n")  # Comentar la línea
                else:
                    modified_lines.append(f"[{ini_key}]\n")  # Descomentar la línea
            else:
                modified_lines.append(line)  # Mantener la línea como está
        else:
            modified_lines.append(line)  # Mantener la línea como está

    # Escribir las líneas modificadas de nuevo al archivo
    with open(launchcfg_file_path, 'w', encoding='utf-8') as file:
        file.writelines(modified_lines)

def edit_file_based_on_dict(launchcfg_path, data_dict):
    with open(launchcfg_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    new_lines = []
    for line in lines:
        line = line.strip()
        key = extract_text_between_brackets(line)
        if key:
            if key in data_dict:
                if data_dict[key] == "on":
                    new_lines.append(f"[{key}]\n")
                else:
                    new_lines.append(f"//[{key}]\n")
            else:
                new_lines.append(line + "\n")
        else:
            new_lines.append(line + "\n")
    
    with open(launchcfg_path, 'w', encoding='utf-8') as file:
        file.writelines(new_lines)

def write_ini(section, data):
    # Leer el archivo .ini existente
    content = {}
    if os.path.exists(ini_file_path):
        with open(ini_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            current_section = None
            for line in lines:
                line = line.strip()
                if line.startswith('[') and line.endswith(']'):
                    current_section = line[1:-1]
                    content[current_section] = {}
                elif '=' in line and current_section:
                    key, value = line.split('=', 1)
                    content[current_section][key] = value

    # Fusionar/actualizar el contenido existente con los nuevos datos
    if section not in content:
        content[section] = {}
    content[section].update(data)
    
    # Escribir de nuevo al archivo .ini
    with open(ini_file_path, 'w', encoding='utf-8') as file:
        for sec, values in content.items():
            file.write(f"[{sec}]\n")
            for key, value in values.items():
                file.write(f"{key}={value}\n")

def get_user_agent_headers():
    return {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

def download_file(url, destination):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Referer": "https://www.heavynight.com/",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    }
    
    response = requests.get(url, headers=headers, stream=True)
    response.raise_for_status()  # Lanza una excepción si la solicitud no fue exitosa

    with open(destination, 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):  # Guardar el archivo en "chunks" para manejar archivos grandes
            file.write(chunk)

def mover_y_actualizar_archivos(categoriavieja, nuevacategoria):
    str_folder = f"launcher/{categoriavieja}"
    str_dest_folder = "launcher/zboveda"
    str_new_folder_name = nuevacategoria

    if not os.path.exists(str_dest_folder):
        os.makedirs(str_dest_folder)

    if os.path.exists(str_folder):
        source_folder_name = os.path.basename(str_folder)
        dest_folder = os.path.join(str_dest_folder, source_folder_name)

        # Aquí creamos la carpeta de destino si no existe
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)

        # Listamos las subcarpetas que vamos a mover
        arr_folders = ["config", "mods", "saves", "scripts"]

        for subfolder in arr_folders:
            source_path = os.path.join(str_folder, subfolder)
            dest_path = os.path.join(dest_folder, subfolder)

            if os.path.exists(source_path):
                shutil.move(source_path, dest_path)

        # Movemos el archivo version.txt
        source_version_file = os.path.join(str_folder, "version.txt")
        dest_version_file = os.path.join(dest_folder, "version.txt")

        if os.path.exists(source_version_file):
            shutil.move(source_version_file, dest_version_file)

        # Renombramos la carpeta
        os.rename(str_folder, os.path.join(os.path.dirname(str_folder), str_new_folder_name))

def ejecutar_seccion_con_manejo_de_errores(seccion):
    try:
        if seccion == "instalar":
            seccion_a1()
        elif seccion == "desinstala":
            seccion_a2()
        elif seccion == "errorversion":
            seccion_a3()
        elif seccion == "mantenimiento":
            seccion_a4()
        elif seccion == "update":
            seccion_a5()
        elif seccion == "instalarc1":
            seccion_b1()
        elif seccion == "desinstalarc1":
            seccion_b2()
        elif seccion == "iniciarc1":
            seccion_b3()
        elif seccion == "parcharc1":
            seccion_b4()
        elif seccion == "nuevainstanciac1":
            seccion_b5()
        elif seccion == "carpetac1":
            seccion_b6()
        elif seccion == "infowebc1":
            seccion_b7()
        elif seccion == "tiendawebc1":
            seccion_b8()
        elif seccion == "ipc1":
            seccion_b9()
        elif seccion == "instalarc2":
            seccion_c1()
        elif seccion == "desinstalarc2":
            seccion_c2()
        elif seccion == "iniciarc2":
            seccion_c3()
        elif seccion == "parcharc2":
            seccion_c4()
        elif seccion == "nuevainstanciac2":
            seccion_c5()
        elif seccion == "carpetac2":
            seccion_c6()
        elif seccion == "infowebc2":
            seccion_c7()
        elif seccion == "tiendawebc2":
            seccion_c8()
        elif seccion == "ipc2":
            seccion_c9()
        elif seccion == "instalarc3":
            seccion_d1()
        elif seccion == "desinstalarc3":
            seccion_d2()
        elif seccion == "iniciarc3":
            seccion_d3()
        elif seccion == "parcharc3":
            seccion_d4()
        elif seccion == "nuevainstanciac3":
            seccion_d5()
        elif seccion == "carpetac3":
            seccion_d6()
        elif seccion == "infowebc3":
            seccion_d7()
        elif seccion == "tiendawebc3":
            seccion_d8()
        elif seccion == "ipc3":
            seccion_d9()
        elif seccion == "instalarc4":
            seccion_e1()
        elif seccion == "desinstalarc4":
            seccion_e2()
        elif seccion == "iniciarc4":
            seccion_e3()
        elif seccion == "parcharc4":
            seccion_e4()
        elif seccion == "nuevainstanciac4":
            seccion_e5()
        elif seccion == "carpetac4":
            seccion_e6()
        elif seccion == "infowebc4":
            seccion_e7()
        elif seccion == "tiendawebc4":
            seccion_e8()
        elif seccion == "ipc4":
            seccion_e9()
        else:
            print("Valor no válido para hn. Use una 'case'.")
    except Exception as e:
        print(f"Error al ejecutar la sección: {e}")

####################### FUNCIONES PARA TODAS LAS CATEGORIAS #############################
def instalar_categoria(categoria, replacements, botones):
    headers = get_user_agent_headers()
    response = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Category-Name.php", headers=headers)
    
    if response.status_code == 200:
        response_lines = response.text.split("<br>")
        carpeta = response_lines[0].strip()
        
        if os.path.exists(f"launcher/{carpeta}/assets"):
            # Borrar archivos
            files_to_delete = ["data/instancia.zip", "data/mods.zip"]
            for file in files_to_delete:
                try:
                    os.remove(file)
                except FileNotFoundError:
                    pass

            # Ediciones en el archivo launchcfg
            edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)

            # Actualizar valores en el archivo .ini
            write_ini("Botones", botones)
            
            popup_message(
            title="HeavyNight",
            message=f"!La instalación de {carpeta} fue exitosa! Abriendo launcher...",
            icon_path='data/info.png',
            button_color='blue'
            )
            subprocess.Popen(["HeavyNight.exe"])
        else:
            respuesta = popup_question(
            title=f"Instalación - {carpeta}",
            message=f"Algo salió mal porque no se reconoció la carpeta {carpeta}. ¿Quieres reportarlo con nuestro soporte?",
            icon_path='data/error.png',
            button_color='red'
            )
        
            # Actuar en función de la respuesta del usuario
            if respuesta:
                webbrowser.open("http://heavynight.com/")
            subprocess.Popen(["HeavyNight.exe"])
    else:
        popup_message(
            title="HeavyNight - Error!",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',
            button_color='red'
            )

def eliminar_categoria(categoria, replacements, botones):
    headers = get_user_agent_headers()
    response = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Category-Name.php", headers=headers)
    
    if response.status_code == 200:
        response_lines = response.text.split("<br>")
        carpeta = response_lines[0].strip()  # Obtiene la primera línea
        
        if os.path.exists(f"launcher/{carpeta}/assets"):
            
            # Usar popup_question para la confirmación
            result = popup_question(
                title="HeavyNight - Desinstalador",
                message="Esta acción eliminará por completo la instancia y no habrá vuelta atrás. Tardará unos segundos y cuando haya terminado se abrirá el launcher nuevamente. ¿Estás seguro?",
                icon_path='launcher/alerta.png',
                button_color='yellow'
            )
            
            if result:
                subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
                
                try:
                    shutil.rmtree(f"launcher/{carpeta}")
                    shutil.rmtree("launcher/forge")
                except Exception as e:
                    # Usar popup_message para mostrar errores
                    popup_message(
                        title="Error",
                        message=f"No se pudo eliminar la(s) carpeta(s): {e}",
                        icon_path='data/error.png',
                        button_color='red'
                    )
                    return  # No continuar si falla la eliminación
                
                # Ediciones en el archivo launchcfg y actualizaciones
                edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)
                write_ini("Botones", botones)

                # Usar popup_message para mostrar información
                popup_message(
                    title="HeavyNight - " + carpeta,
                    message="Se eliminaron los archivos con éxito. Abriendo launcher...",
                    icon_path='data/info.png',
                    button_color='blue'
                )
                subprocess.Popen(["HeavyNight.exe"])
            else:
                # Usar popup_message para cancelación
                popup_message(
                    title="HeavyNight",
                    message="Operación cancelada.",
                    icon_path='data/info.png',
                    button_color='blue'
                )
        else:
            # Usar popup_message para error de carpeta no encontrada
            popup_message(
                title="HeavyNight - Error",
                message=f"No se encontró la carpeta {carpeta}.",
                icon_path='data/error.png',
                button_color='red'
            )
            subprocess.Popen(["HeavyNight.exe"])
    else:
        # Usar popup_message para error de conexión
        popup_message(
            title="HeavyNight - Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',
            button_color='red'
        )

def iniciar_categoria(categoria):
    headers = get_user_agent_headers()
    response = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Category-Name.php", headers=headers)

    if response.status_code == 200:
        response_lines = response.text.split("<br>")
        carpeta, cversion, _, ip, forge, cjava = response_lines[:6]
        
        if os.path.exists(f"launcher/{carpeta}"):
            if os.path.exists("C:/Program Files/java/jdk-17.0.6/bin/javaw.exe"):
                dest_path = f"launcher/{carpeta}/version.txt"
                if not os.path.exists(dest_path):
                    with open(dest_path, "w") as file:
                        file.write("1.0.0")

                with open(dest_path, "r") as file:
                    version = file.readline().strip()

                remote_version = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Modpack/{carpeta}/version.txt", headers=headers).text.strip()
                
                if version == remote_version:
                    subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
                    download_file("https://www.heavynight.com/launcherV5/launcher_configs.js", "launcher/resources/app/launcher_config.js")
                    
                    with open("launcher/resources/app/launcher_config.js", "r") as file:
                        config_content = file.read()
                    
                    config_content = config_content.replace("{category-ip}", ip)
                    config_content = config_content.replace("{category-java}", cjava)
                    config_content = config_content.replace("{category-name}", carpeta)
                    config_content = config_content.replace("{category-version}", cversion)
                    config_content = config_content.replace("{category-forge}", forge)
                    
                    with open("launcher/resources/app/launcher_config.js", "w") as file:
                        file.write(config_content)

                    original_cwd = os.getcwd()  # Guarda el directorio de trabajo actual
                    os.chdir(os.path.join(original_cwd, 'launcher'))
                    subprocess.Popen(["login.exe"])
                    os.chdir(original_cwd)  # Restablece el directorio de trabajo original

                else:
                    # Usar popup_question para confirmación de actualización
                    result = popup_question(
                        title=f"HeavyNight - {carpeta}",
                        message="¡Hay una actualización pendiente! ¿Quieres actualizarla?",
                        icon_path='data/alerta.png',
                        button_color='yellow'
                    )
                    if result:
                        actualizar_y_parchear(carpeta)
            else:
                # Usar popup_message para error de Java
                popup_message(
                    title="HeavyNight - Error de inicio",
                    message=f"{carpeta} necesita Java 17 y parece que algo ha fallado en la integración de java. Por favor, contacta con nuestro soporte o vuelve a reinstalar el launcher.",
                    icon_path='data/error.png',
                    button_color='red'
                )
                if popup_question(
                    title="HeavyNight - Soporte",
                    message="¿Quieres contactar con nuestro soporte?",
                    icon_path='data/info.png',
                    button_color='blue'
                ):
                    webbrowser.open("http://heavynight.com/")
        else:
            # Usar popup_message para información de descarga pendiente
            popup_message(
                title="HeavyNight!",
                message=f"Aún no tienes descargado {carpeta}.",
                icon_path='data/alerta.png',
                button_color='yellow'
            )
    else:
        # Usar popup_message para error de conexión
        popup_message(
            title="Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',
            button_color='red'
        )

def actualizar_y_parchear(carpeta):
    if os.path.exists("C:/Program Files/java/Jre_8/bin/javaw.exe"):
        subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
        
        # Actualizar archivo config_sync.json
        download_file("https://www.heavynight.com/launcherV5/config_sync.json", "launcher/config_sync.json")
        with open("launcher/config_sync.json", "r") as file:
            config_content = file.read()
        config_content = config_content.replace("{category-name}", carpeta)
        with open("launcher/config_sync.json", "w") as file:
            file.write(config_content)

        subprocess.run(["launcher/server_sync.exe", "c1serversync"])
        subprocess.Popen(["HeavyNight.exe"])

        # Usar popup_message para mostrar información
        popup_message(
            title="HeavyNight - Parches",
            message="El parche ha terminado.",
            icon_path='data/info.png',
            button_color='blue'
        )
    else:
        # Usar popup_message para error de Java
        popup_message(
            title="HeavyNight - Error",
            message="Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.",
            icon_path='data/error.png',
            button_color='red'
        )
        if popup_question(
            title="HeavyNight - Soporte",
            message="¿Quieres contactar con nuestro soporte?",
            icon_path='data/info.png',
            button_color='blue'
        ):
            webbrowser.open("http://heavynight.com/")

def parchear_categoria(categoria):
    headers = get_user_agent_headers()
    response = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Category-Name.php", headers=headers)

    if response.status_code == 200:
        response_lines = response.text.split("<br>")
        carpeta = response_lines[0].strip()
        
        dest_path = f"launcher/{carpeta}/version.txt"
        
        if not os.path.exists(dest_path):
            with open(dest_path, "w") as file:
                file.write("1.0.0")

        with open(dest_path, "r") as file:
            version = file.readline().strip()

        remote_version = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Modpack/{carpeta}/version.txt", headers=headers).text.strip()
        
        if version == remote_version:
            result = popup_question(
                title=f"HeavyNight - {carpeta}",
                message="¡Ya tienes la última actualización! ¿Quieres actualizarla igualmente?",
                icon_path='launcher/actualizacion.png',
                button_color='yellow'
            )
            if result:
                actualizar_y_parchear(carpeta)
        else:
            actualizar_y_parchear(carpeta)
    else:
        # Usar popup_message para mostrar error
        popup_message(
            title="HeavyNight - Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',
            button_color='red'
        )

def actualizar_categoria(categoria, categoria_ini_key, image_suffix, replacements, botones, ini_file_path):
    # Leer la categoría vieja del archivo
    ini_data = read_ini(ini_file_path)
    categoriavieja = ini_data.get("Categorias", {}).get(categoria_ini_key, "")
    
    headers = get_user_agent_headers()
    response = requests.get(f"https://www.heavynightlauncher.com/Launcher-Categorias/{categoria}/Category-Name.php", headers=headers)
    
    if response.status_code == 200:
        nuevacategoria = response.text.split('<br>')[0].strip()

        print (f"{categoriavieja} y {nuevacategoria}")
        
        # Comprobar si las categorías son diferentes
        if categoriavieja.lower() != nuevacategoria.lower():
            carpeta_vieja_path = os.path.join("launcher", categoriavieja)

            result= popup_question(
                    title="HeavyNight - Categorias",
                    message=f"Hemos marcado la categoria {categoriavieja} como 'CERRADA' ya que hay una nueva disponible actualmente llamada {nuevacategoria}.\n¿Quieres actualizar a la nueva categoria?",
                    icon_path='data/alerta.png',
                    button_color='yellow'
                )
            if result:
               if os.path.exists(carpeta_vieja_path):
                   result = popup_question(
                       title="HeavyNight - Categorias",
                       message=f"Hemos identificado que existe la carpeta {categoriavieja}.\n¿Quieres realizarle una copia de seguridad de tus archivos y progresos a tu vobeda?",
                       icon_path='data/alerta.png',
                       button_color='yellow'
                   )
                   if result:
                       subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
                       mover_y_actualizar_archivos(categoriavieja, nuevacategoria)
   
                       # Actualización de imágenes y configuraciones
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/logo.png", f"data/logo-C{image_suffix}.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/titulo.png", f"data/titulo-C{image_suffix}.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono1.png", f"data/b_icono-c{image_suffix}_a.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono2.png", f"data/b_icono-c{image_suffix}_b.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono3.png", f"data/b_icono-c{image_suffix}_c.png")
   
                       # Actualizando el archivo .ini con la nueva categoría
                       ini_data["Categorias"][categoria_ini_key] = nuevacategoria
                       write_ini("Categorias", ini_data["Categorias"])
                       
                       # Ediciones en el archivo launchcfg
                       edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)
   
                       # Actualizar valores en el archivo .ini
                       write_ini("Botones", botones)    
   
                       subprocess.Popen(["HeavyNight.exe"])
                       popup_message(
                           title="HeavyNight",
                           message=f"La categoria {categoriavieja} ha cambiado y se ha creado una copia en su bóveda para presentarte a nuestra nueva categoría {nuevacategoria}.",
                           icon_path='data/info.png',
                           button_color='blue'
                       )
                   else:
                       for subFolder in ["config", "mods", "saves", "scripts"]:
                           folderDel = os.path.join(carpeta_vieja_path, subFolder)
                           if os.path.exists(folderDel):
                               shutil.rmtree(folderDel)
   
                       # Renombrar carpeta vieja
                       os.rename(carpeta_vieja_path, os.path.join("launcher", nuevacategoria))
   
                       # Actualización de imágenes y configuraciones
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/logo.png", f"data/logo-C{image_suffix}.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/titulo.png", f"data/titulo-C{image_suffix}.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono1.png", f"data/b_icono-c{image_suffix}_a.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono2.png", f"data/b_icono-c{image_suffix}_b.png")
                       download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono3.png", f"data/b_icono-c{image_suffix}_c.png")
       
                       # Actualizando el archivo .ini con la nueva categoría
                       ini_data["Categorias"][categoria_ini_key] = nuevacategoria
                       write_ini("Categorias", ini_data["Categorias"])
       
                       # Ediciones en el archivo launchcfg
                       edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)
       
                       # Actualizar valores en el archivo .ini
                       write_ini("Botones", botones)
       
                       popup_message(
                           title="HeavyNight",
                           message=f"La categoria {categoriavieja} ha cambiado y ahora te presentamos {nuevacategoria}.",
                           icon_path='data/info.png',
                           button_color='blue'
                       )
               else:
                   subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
                   # Actualización de imágenes y configuraciones
                   download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/logo.png", f"data/logo-C{image_suffix}.png")
                   download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/titulo.png", f"data/titulo-C{image_suffix}.png")
                   download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono1.png", f"data/b_icono-c{image_suffix}_a.png")
                   download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono2.png", f"data/b_icono-c{image_suffix}_b.png")
                   download_file(f"https://heavynightlauncher.com/Launcher-Categorias/{categoria}/imagenes/icono3.png", f"data/b_icono-c{image_suffix}_c.png")
    
                   # Actualizando el archivo .ini con la nueva categoría
                   ini_data["Categorias"][categoria_ini_key] = nuevacategoria
                   write_ini("Categorias", ini_data["Categorias"])

                   subprocess.Popen(["HeavyNight.exe"])
                   popup_message(
                       title="HeavyNight",
                       message=f"La categoria {categoriavieja} ha cambiado y ahora te presentamos {nuevacategoria}.",
                       icon_path='data/info.png',
                       button_color='blue'
                   )
            else:
                pass
        else:
            pass

def abrir_carpeta_categoria(categoria):
    base_url = "https://www.heavynightlauncher.com/Launcher-Categorias"
    headers = get_user_agent_headers()
    full_url = f"{base_url}/{categoria}/Category-Name.php"
    
    response = requests.get(full_url, headers=headers)
    
    if response.status_code == 200:
        response_lines = response.text.split("<br>") 
        carpeta = response_lines[0].strip()
        
        path_to_folder = os.path.join("launcher", carpeta)
        if os.path.exists(path_to_folder):
            subprocess.Popen(f'explorer "{path_to_folder}"', shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            popup_message(
                title="HeavyNight - Error",
                message=f"Parece que aún no tienes instalada la categoría o no existe la carpeta {carpeta}.",
                icon_path='data/error.png',  # Cambia la ruta a la nueva imagen de ícono aquí
                button_color='red'  # Cambia el color de fondo del botón aquí
            )
    else:
        popup_message(
            title="HeavyNight - Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',  # Cambia la ruta a la nueva imagen de ícono aquí
            button_color='red'  # Cambia el color de fondo del botón aquí
        )

def obtener_y_abrir_url(categoria, url_parte):
    base_url = "https://www.heavynightlauncher.com/Launcher-Categorias"
    headers = get_user_agent_headers()
    full_url = f"{base_url}/{categoria}/Category-Name.php"
    
    response = requests.get(full_url, headers=headers)
    
    if response.status_code == 200:
        response_lines = response.text.split("<br>") 
        paginaweb = response_lines[0].strip()

        link = f"https://www.heavynight.com/{url_parte}/{paginaweb}"
        webbrowser.open(link)
    else:
        popup_message(
            title="HeavyNight - Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',  # Cambia la ruta a la nueva imagen de ícono aquí
            button_color='red'  # Cambia el color de fondo del botón aquí
        )

def obtener_y_copiar_ip(categoria):
    base_url = "https://www.heavynightlauncher.com/Launcher-Categorias"
    headers = get_user_agent_headers()
    full_url = f"{base_url}/{categoria}/Category-Name.php"
    
    response = requests.get(full_url, headers=headers)
    
    if response.status_code == 200:
        response_lines = response.text.split("<br>") 
        ipserver = response_lines[3].strip()  
        
        # Copiar IP al portapapeles usando tkinter
        root = tk.Tk()
        root.withdraw()  # esconde la ventana principal
        root.clipboard_clear()
        root.clipboard_append(ipserver)
        root.update()  # ahora está almacenado en el portapapeles y puedes pegar el contenido en otra aplicación o ventana
        root.destroy()

        # Mostrar mensaje de confirmación
        popup_message(
            title="HeavyNight!",
            message="La IP ha sido copiada a tu portapapeles.",
            icon_path='data/info.png',  # Cambia la ruta a la nueva imagen de ícono aquí
            button_color='blue'  # Cambia el color de fondo del botón aquí
        )
    else:
        popup_message(
            title="HeavyNight - Error",
            message=f"No se pudo obtener el valor de la URL. Código de estado: {response.status_code}",
            icon_path='data/error.png',  # Cambia la ruta a la nueva imagen de ícono aquí
            button_color='red'  # Cambia el color de fondo del botón aquí
        )


####################### FUNCIONES DEL LAUNCHER #############################
# Instala el launcher
def seccion_a1():
    # Eliminar archivos si existen
    for filename in ["data/categorias.zip", "data/java.zip"]:
        try:
            os.remove(filename)
        except FileNotFoundError:
            pass

    # Editar el archivo "launchcfg"
    replacements = [
        ("[p_0_img_button_8]", "//[p_0_img_button_8]"),
        ("//[p_0_img_button_9]", "[p_0_img_button_9]"),
        ("//[p_0_img_button_6]", "[p_0_img_button_6]"),
        ("//[p_0_img_button_7]", "[p_0_img_button_7]"),
        ("//[p_0_img_button_11]", "[p_0_img_button_11]"),
        ("//[p_0_img_button_12]", "[p_0_img_button_12]")
    ]
    edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)

    # Registros para archivo ini
    botones = {
        "p_0_img_button_8": "off",
        "p_0_img_button_9": "on",
        "p_0_img_button_6": "on",
        "p_0_img_button_7": "on",
        "p_0_img_button_11": "on",
        "p_0_img_button_12": "on"
    }
    write_ini("Botones", botones)

    # Mostrar mensaje de éxito
    popup_message(
        title="HeavyNight!",
        message="¡La instalación fue exitosa! Iniciando launcher...",
        icon_path='data/info.png',  # Cambia la ruta a la nueva imagen de ícono aquí
        button_color='blue'  # Cambia el color de fondo del botón aquí
    )
    
    # Iniciar el programa HeavyNight.exe
    subprocess.Popen(["HeavyNight.exe"])

# Desinstala el launcher
def seccion_a2():
    if os.path.exists("launcher/resources"):
        result = popup_question(
            title="HeavyNight - Desinstalador",
            message="Esta acción eliminará por completo las categorías y no habrá vuelta atrás. Tardará unos segundos y cuando haya terminado se abrirá el launcher nuevamente. ¿Estás seguro?",
            icon_path='data/alerta.png',
            button_color='yellow'
        )

        if result:
            subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)

            # Eliminar carpetas una por una
            folders_to_delete = [
                r"C:\Program Files\Java\jdk-17.0.6",
                r"C:\Program Files\Java\Jre_8",
                "launcher/locales",
                "launcher/resources",
                "launcher/forge"
            ]
            
            for folder in folders_to_delete:
                if os.path.exists(folder):
                    # Quita el atributo de solo lectura de todos los archivos y subcarpetas recursivamente
                    subprocess.call(f'attrib -R {folder}\* /S /D', shell=True)
                    # Elimina la carpeta y todo su contenido
                    subprocess.call(f'rmdir /s /q "{folder}"', shell=True)

            # Eliminar tipos de archivos específicos uno por uno
            file_types_to_delete = [".exe", ".dat", ".json", ".pak", ".bin", ".dll"]
            for file_type in file_types_to_delete:
                files = glob.glob(f"launcher/*{file_type}")
                for file in files:
                    if os.path.basename(file).lower() != "heavynight-system.exe":
                        # Quita el atributo de solo lectura del archivo
                        subprocess.call(f'attrib -R "{file}"', shell=True)
                        # Elimina el archivo
                        subprocess.call(f'del /f /q "{file}"', shell=True)

            replacements = [
                ("//[p_0_img_button_8]", "[p_0_img_button_8]"),
                ("[p_0_img_button_9]", "//[p_0_img_button_9]"),
                ("[p_0_img_button_6]", "//[p_0_img_button_6]"),
                ("[p_0_img_button_7]", "//[p_0_img_button_7]"),
                ("[p_0_img_button_11]", "//[p_0_img_button_11]"),
                ("[p_0_img_button_12]", "//[p_0_img_button_12]")
            ]
            edit_launch_cfg_multiple_replacements(launchcfg_path, replacements)

            botones = {
                "p_0_img_button_8": "on",
                "p_0_img_button_9": "off",
                "p_0_img_button_6": "off",
                "p_0_img_button_7": "off",
                "p_0_img_button_11": "off",
                "p_0_img_button_12": "off"
            }
            write_ini("Botones", botones)

            # Mensaje de éxito
            popup_message(
                title="HeavyNight!",
                message="Se eliminaron los archivos con éxito. Abriendo launcher...",
                icon_path='data/info.png',
                button_color='blue'
            )
            subprocess.Popen(["HeavyNight.exe"])
        else:
            # Mensaje de operación cancelada
            popup_message(
                title="HeavyNight!",
                message="Operación cancelada.",
                icon_path='data/info.png',
                button_color='blue'
            )
    else:
        # Mensaje de error si el directorio 'resources' no existe
        popup_message(
            title="HeavyNight!",
            message="El directorio 'resources' no existe.",
            icon_path='data/error.png',
            button_color='red'
        )

# Cuando no se encuentra el archivo version.txt
def seccion_a3():
    popup_message(
    title="Aviso de HeavyNight",
    message="!No se pudo obtener el archivo version.txt y es posible que no recibas actualizaciones futuras.",
    icon_path='data/alerta.png',  # Cambia la ruta a la nueva imagen de ícono aquí
    button_color='orange'  # Cambia el color de fondo del botón aquí
    )

# Cuando el launcher está en mantenimiento
def seccion_a4():
    headers = get_user_agent_headers()
    # Hacer la solicitud con los encabezados definidos
    response = requests.get("https://www.heavynight.com/launcherV5/Mantenimiento.php", headers=headers)
    
    maintenance = response.text.strip().lower()
    
    # Imprimir la respuesta para depurar
    print(f"Respuesta de mantenimiento: {maintenance}")

    if maintenance == "true":
        # Cerrar el programa HeavyNight.exe si está abierto
        subprocess.run("taskkill /f /im HeavyNight.exe", shell=True, check=True)
        
        # Mostrar el mensaje en una ventana emergente
        respuesta = popup_question(
        title="HeavyNight - Mantenimiento!",
        message="El launcher está en mantenimiento.\n\n¿Te muestro el canal de discord para que te enteres cuando el launcher deja de estar en mantenimiento?",
        icon_path='data/alerta.png',
        button_color='yellow'
        )
        
        # Actuar en función de la respuesta del usuario
        if respuesta:
            webbrowser.open("https://discord.com/channels/860007074695610398/1015623724600414218")

# Para un ubdate
def seccion_a5():
    # Abrir la URL en el navegador predeterminado
    webbrowser.open("https://www.heavynight.com/changelog/categories/9")

    # Eliminar la carpeta "HeavyNight.exe.WebView2" si existe
    folder_to_delete = "HeavyNight.exe.WebView2"
    if os.path.exists(folder_to_delete):
        shutil.rmtree(folder_to_delete)

    # Leer el archivo INI y almacenar los valores en un diccionario
    ini_data = read_ini(ini_file_path)

    # Editar el archivo launchcfg basado en los datos del archivo INI
    edit_launchcfg_based_on_ini(ini_data, launchcfg_path)

    # Intentar ejecutar el programa HeavyNight.exe
    try:
        subprocess.Popen(["HeavyNight.exe"])
    except Exception as e:
        # Mostrar mensaje de error
        popup_message(
            title="Error al ejecutar HeavyNight.exe",
            message=str(e),
            icon_path='data/error.png',
            button_color='red'
        )

####################### FUNCIONES PARA CATEGORIA 1 #############################

# Intala la categoria
def seccion_b1():
    replacements_categoria_1 = [
        ("[p_1_img_button_5]", "//[p_1_img_button_5]"), #boton de descargar
        ("//[p_1_img_button_4]", "[p_1_img_button_4]"), #Boton de jugar
        ("//[p_1_img_button_11]", "[p_1_img_button_11]"), #Boton de desinstalar
        ("//[p_1_img_button_12]", "[p_1_img_button_12]") #Boton de parches
    ]
    botones_categoria_1 = {
        "p_1_img_button_5": "off",
        "p_1_img_button_4": "on",
        "p_1_img_button_11": "on",
        "p_1_img_button_12": "on"
    }
    instalar_categoria("Categoria1", replacements_categoria_1, botones_categoria_1)

# Desinstala la categoria
def seccion_b2():
    replacements_categoria_1 = [
        ("//[p_1_img_button_5]", "[p_1_img_button_5]"), #Boton de descargar
        ("[p_1_img_button_4]", "//[p_1_img_button_4]"), #Boton de jugar
        ("[p_1_img_button_11]", "//[p_1_img_button_11]"), #Boton de desinstalar
        ("[p_1_img_button_12]", "//[p_1_img_button_12]") #Boton de parches
    ]
    botones_categoria_1 = {
        "p_1_img_button_5": "on",
        "p_1_img_button_4": "off",
        "p_1_img_button_11": "off",
        "p_1_img_button_12": "off"
    }
    eliminar_categoria("Categoria1", replacements_categoria_1, botones_categoria_1)

# Inicia la categoria
def seccion_b3():
    iniciar_categoria("Categoria1")

# Parcha el modpack de la categoria
def seccion_b4():
    parchear_categoria("Categoria1")

# Update de la instancia completa
def seccion_b5():
    # Llamada a la función con los argumentos específicos para la categoría 1
    replacements_categoria_1 = [
        ("[p_1_img_button_12]", "//[p_1_img_button_12]"), #Boton de parches
        ("[p_1_img_button_11]", "//[p_1_img_button_11]"), #Boton de desinstalar
        ("[p_1_img_button_4]", "//[p_1_img_button_4]"), #Boton de jugar
        ("[p_1_img_button_5]", "//[p_1_img_button_5]"), #Boton de descargar
        ("//[p_1_img_button_13]", "[p_1_img_button_13]") #Boton de intalar
    ]
    botones_categoria_1 = {
        "p_1_img_button_12": "off",
        "p_1_img_button_11": "off",
        "p_1_img_button_4": "off",
        "p_1_img_button_5": "off",
        "p_1_img_button_13": "on"
    }
    actualizar_categoria("Categoria1", "C1", "1", replacements_categoria_1, botones_categoria_1, ini_file_path)

# Abre la carpeta de la categoria
def seccion_b6():
    abrir_carpeta_categoria("Categoria1")

# Abre la infoweb de la categoria
def seccion_b7():
    obtener_y_abrir_url("Categoria1", "news")

# Abre la tiendaweb de la categoria
def seccion_b8():
    obtener_y_abrir_url("Categoria1", "shop/categories")

# Copia la ip de la categoria
def seccion_b9():
    obtener_y_copiar_ip("Categoria1")

####################### FUNCIONES PARA CATEGORIA 2 #############################

# Intala la categoria
def seccion_c1():
    replacements_categoria_2 = [
        ("[p_2_img_button_5]", "//[p_2_img_button_5]"), #boton de descargar
        ("//[p_2_img_button_4]", "[p_2_img_button_4]"), #Boton de jugar
        ("//[p_2_img_button_10]", "[p_2_img_button_10]"), #Boton de desinstalar
        ("//[p_2_img_button_1]", "[p_2_img_button_1]") #Boton de parches
    ]
    botones_categoria_2 = {
        "p_2_img_button_5": "off",
        "p_2_img_button_4": "on",
        "p_2_img_button_10": "on",
        "p_2_img_button_1": "on"
    }
    instalar_categoria("Categoria2", replacements_categoria_2, botones_categoria_2)

# Desinstala la categoria
def seccion_c2():
    replacements_categoria_2 = [
        ("//[p_2_img_button_5]", "[p_2_img_button_5]"), #Boton de descargar
        ("[p_2_img_button_4]", "//[p_2_img_button_4]"), #Boton de jugar
        ("[p_2_img_button_10]", "//[p_2_img_button_10]"), #Boton de desinstalar
        ("[p_2_img_button_1]", "//[p_2_img_button_1]") #Boton de parches
    ]
    botones_categoria_2 = {
        "p_2_img_button_5": "on",
        "p_2_img_button_4": "off",
        "p_2_img_button_10": "off",
        "p_2_img_button_1": "off"
    }
    eliminar_categoria("Categoria2", replacements_categoria_2, botones_categoria_2)

# Inicia la categoria
def seccion_c3():
    iniciar_categoria("Categoria2")

# Parcha el modpack de la categoria
def seccion_c4():
    parchear_categoria("Categoria2")

# Update de la instancia completa
def seccion_c5():
    replacements_categoria_2 = [
        ("[p_2_img_button_1]", "//[p_2_img_button_1]"), #Boton de parches
        ("[p_2_img_button_10]", "//[p_2_img_button_10]"), #Boton de desinstalar
        ("[p_2_img_button_4]", "//[p_2_img_button_4]"), #Boton de jugar
        ("[p_2_img_button_5]", "//[p_2_img_button_5]"), #Boton de descargar
        ("//[p_2_img_button_11]", "[p_2_img_button_11]") #Boton de intalar
    ]
    botones_categoria_2 = {
        "p_2_img_button_1": "off",
        "p_2_img_button_10": "off",
        "p_2_img_button_4": "off",
        "p_2_img_button_5": "off",
        "p_2_img_button_11": "on"
    }
    actualizar_categoria("Categoria2", "C2", "2", replacements_categoria_2, botones_categoria_2, ini_file_path)

# Abre la carpeta de la categoria
def seccion_c6():
    abrir_carpeta_categoria("Categoria2")

# Abre la infoweb de la categoria
def seccion_c7():
    obtener_y_abrir_url("Categoria2", "news")

# Abre la tiendaweb de la categoria
def seccion_c8():
    obtener_y_abrir_url("Categoria2", "shop/categories")

# Copia la ip de la categoria
def seccion_c9():
    obtener_y_copiar_ip("Categoria2")

####################### FUNCIONES PARA CATEGORIA 3 #############################

# Intala la categoria
def seccion_d1():
    replacements_categoria_3 = [
        ("[p_3_img_button_7]", "//[p_3_img_button_7]"), #boton de descargar
        ("//[p_3_img_button_8]", "[p_3_img_button_8]"), #Boton de jugar
        ("//[p_3_img_button_10]", "[p_3_img_button_10]"), #Boton de desinstalar
        ("//[p_3_img_button_4]", "[p_3_img_button_4]") #Boton de parches
    ]
    botones_categoria_3 = {
        "p_3_img_button_7": "off",
        "p_3_img_button_8": "on",
        "p_3_img_button_10": "on",
        "p_3_img_button_4": "on"
    }
    instalar_categoria("Categoria3", replacements_categoria_3, botones_categoria_3)

# Desinstala la categoria
def seccion_d2():
    replacements_categoria_3 = [
        ("//[p_3_img_button_7]", "[p_3_img_button_7]"), #Boton de descargar
        ("[p_3_img_button_8]", "//[p_3_img_button_8]"), #Boton de jugar
        ("[p_3_img_button_10]", "//[p_3_img_button_10]"), #Boton de desinstalar
        ("[p_3_img_button_4]", "//[p_3_img_button_4]") #Boton de parches
    ]
    botones_categoria_3 = {
        "p_3_img_button_7": "on",
        "p_3_img_button_8": "off",
        "p_3_img_button_10": "off",
        "p_3_img_button_4": "off"
    }
    eliminar_categoria("Categoria3", replacements_categoria_3, botones_categoria_3)

# Inicia la categoria
def seccion_d3():
    iniciar_categoria("Categoria3")

# Parcha el modpack de la categoria
def seccion_d4():
    parchear_categoria("Categoria3")

# Update de la instancia completa
def seccion_d5():
    # Llamada a la función con los argumentos específicos para la categoría 1
    replacements_categoria_3 = [
        ("[p_3_img_button_4]", "//[p_3_img_button_4]"), #Boton de parches
        ("[p_3_img_button_10]", "//[p_3_img_button_10]"), #Boton de desinstalar
        ("[p_3_img_button_8]", "//[p_3_img_button_8]"), #Boton de jugar
        ("[p_3_img_button_7]", "//[p_3_img_button_7]"), #Boton de descargar
        ("//[p_3_img_button_9]", "[p_3_img_button_9]") #Boton de intalar
    ]
    botones_categoria_3 = {
        "p_3_img_button_4": "off",
        "p_3_img_button_10": "off",
        "p_3_img_button_8": "off",
        "p_3_img_button_7": "off",
        "p_3_img_button_9": "on"
    }
    actualizar_categoria("Categoria3", "C3", "3", replacements_categoria_3, botones_categoria_3, ini_file_path)

# Abre la carpeta de la categoria
def seccion_d6():
    abrir_carpeta_categoria("Categoria3")

# Abre la infoweb de la categoria
def seccion_d7():
    obtener_y_abrir_url("Categoria3", "news")

# Abre la tiendaweb de la categoria
def seccion_d8():
    obtener_y_abrir_url("Categoria3", "shop/categories")

# Copia la ip de la categoria
def seccion_d9():
    obtener_y_copiar_ip("Categoria3")

####################### FUNCIONES PARA CATEGORIA 4 #############################

# Intala la categoria
def seccion_e1():
    replacements_categoria_4 = [
        ("[p_6_img_button_3]", "//[p_6_img_button_3]"), #boton de descargar
        ("//[p_6_img_button_11]", "[p_6_img_button_11]"), #Boton de jugar
        ("//[p_6_img_button_6]", "[p_6_img_button_6]"), #Boton de desinstalar
        ("//[p_6_img_button_10]", "[p_6_img_button_10]") #Boton de parches
    ]
    botones_categoria_4 = {
        "p_6_img_button_3": "off",
        "p_6_img_button_11": "on",
        "p_6_img_button_6": "on",
        "p_6_img_button_10": "on"
    }
    instalar_categoria("Categoria4", replacements_categoria_4, botones_categoria_4)

# Desinstala la categoria
def seccion_e2():
    replacements_categoria_4 = [
        ("//[p_6_img_button_3]", "[p_6_img_button_3]"), #Boton de descargar
        ("[p_6_img_button_11]", "//[p_6_img_button_11]"), #Boton de jugar
        ("[p_6_img_button_6]", "//[p_6_img_button_6]"), #Boton de desinstalar
        ("[p_6_img_button_10]", "//[p_6_img_button_10]") #Boton de parches
    ]
    botones_categoria_4 = {
        "p_6_img_button_3": "on",
        "p_6_img_button_11": "off",
        "p_6_img_button_6": "off",
        "p_6_img_button_10": "off"
    }
    eliminar_categoria("Categoria4", replacements_categoria_4, botones_categoria_4)

# Inicia la categoria
def seccion_e3():
    iniciar_categoria("Categoria4")

# Parcha el modpack de la categoria
def seccion_e4():
    parchear_categoria("Categoria4")

# Update de la instancia completa
def seccion_e5():
    replacements_categoria_4 = [
        ("[p_6_img_button_10]", "//[p_6_img_button_10]"), #Boton de parches
        ("[p_6_img_button_6]", "//[p_6_img_button_6]"), #Boton de desinstalar
        ("[p_6_img_button_11]", "//[p_6_img_button_11]"), #Boton de jugar
        ("[p_6_img_button_3]", "//[p_6_img_button_3]"), #Boton de descargar
        ("//[p_6_img_button_12]", "[p_6_img_button_12]") #Boton de intalar
    ]
    botones_categoria_4 = {
        "p_6_img_button_10": "off",
        "p_6_img_button_6": "off",
        "p_6_img_button_11": "off",
        "p_6_img_button_3": "off",
        "p_6_img_button_12": "on"
    }
    actualizar_categoria("Categoria4", "C4", "4", replacements_categoria_4, botones_categoria_4, ini_file_path)

# Abre la carpeta de la categoria
def seccion_e6():
    abrir_carpeta_categoria("Categoria4")

# Abre la infoweb de la categoria
def seccion_e7():
    obtener_y_abrir_url("Categoria4", "news")

# Abre la tiendaweb de la categoria
def seccion_e8():
    obtener_y_abrir_url("Categoria4", "shop/categories")

# Copia la ip de la categoria
def seccion_e9():
    obtener_y_copiar_ip("Categoria4")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Error: No se ha detectado un argumento de inicio valido.")
        sys.exit(1)

    hn = sys.argv[1]
    ejecutar_seccion(hn)