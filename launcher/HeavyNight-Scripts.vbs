' Habilita la comprobación explícita de variables no declaradas
Option Explicit
Set objFSO = CreateObject("Scripting.FileSystemObject")

' \\\\\\\\\\\ SECCIÓN DE TODAS LAS VARIABLES \\\\\\\\\\\
 Dim hn, objXmlHttp, objADOStream, objFSO, objFolder, strLocalFolderPath, strUrl, paginaweb, scriptPath, configFile, configContent
 Dim strLocalFolderName, strRemoteFolderName, objShell, result, oShell, strFolder, response, scriptFolder, ip, forge, cversion
 Dim strDestFolder, strNewFolderName, sourceFolderName, destFolder, WshShell, link, request, lockFile, cjava, utf16Stream
 Dim Return, FolderDel, rename_file, obj, texto, MyBox, fso, carpeta, respuesta, file, maintenance, dataFolder
 Dim categoriavieja, nuevacategoria, arrFolders, subFolder, destPath, fileContent, i, line, winHttpReq, carpetaViejaPath
 Dim fs, currentFolder, versionFolderPath, versionPath, versionFile, version, url, objFile, urlRemota
 Dim xmlhttp, remoteVersion, responseLines, lineNumber, rutamods, resultado, currentDir, UrlList, ipserver
 Dim strLocalFilePath, colProcesses, objProcess, lines, shell, returnCode, objWMIService, strFolderPath
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' Obtiene el argumento "hn" proporcionado en la línea de comandos
hn = WScript.Arguments.Named("hn")

' Verifica si se proporcionó el argumento "hn"
If hn = "" Then
    WScript.Echo "Error: No se ha detectado un argumento de inicio valido."
    WScript.Quit
End If

' \\\\\\\\\\\ SECCIÓN DE TODAS LAS FUNCIONES REQUERIDAS \\\\\\\\\\\

 ' Verifica si la sección está en ejecución antes de ejecutarla
 Function EjecutarSeccion(Seccion)
     ' Verificar si el archivo de bloqueo existe antes de ejecutar la sección
     If Not CheckLockFile(Seccion) Then
         ' Crear el archivo de bloqueo antes de ejecutar la sección
         CreateLockFile(Seccion)
 
         ' Intentar ejecutar la sección con manejo de errores
         EjecutarSeccionConManejoDeErrores(Seccion)
 
         ' Eliminar el archivo de bloqueo si no ocurrió ningún error
         If Err.Number = 0 Then
             DeleteLockFile(Seccion)
         End If
     End If
 End Function

 ' Función para obtener la carpeta
 Function GetScriptFolder()
     Set fso = CreateObject("Scripting.FileSystemObject")
     scriptPath = WScript.Arguments(0)
     GetScriptFolder = fso.GetParentFolderName(scriptPath)
 End Function
 
 ' Función para obtener la carpeta Data
 Function GetDataFolder()
     Set fso = CreateObject("Scripting.FileSystemObject")
     scriptFolder = GetScriptFolder()
     GetDataFolder = fso.BuildPath(scriptFolder, "Data")
 End Function

 ' Función para verificar el archivo de bloqueo por ejecuccion de seccion
 Function CheckLockFile(Seccion)
     Set fso = CreateObject("Scripting.FileSystemObject")
     lockFile = fso.BuildPath(GetDataFolder(), "Lock_" & Seccion & ".lock")
     
     If fso.FileExists(lockFile) Then
         CheckLockFile = True ' El archivo de bloqueo existe, lo cual indica que la sección está en ejecución
     Else
         CheckLockFile = False ' El archivo de bloqueo no existe, lo cual indica que la sección no está en ejecución
     End If
 End Function

 ' Función para crear el archivo de bloqueo por ejecuccion de seccion
 Sub CreateLockFile(Seccion)
     Set fso = CreateObject("Scripting.FileSystemObject")
     dataFolder = GetDataFolder()
     lockFile = fso.BuildPath(dataFolder, "Lock_" & Seccion & ".lock")
     
     ' Crea la carpeta "Data" en el directorio del llamador si no existe
     If Not fso.FolderExists(dataFolder) Then
         fso.CreateFolder(dataFolder)
     End If
     
     ' Crea el archivo de bloqueo
     fso.CreateTextFile(lockFile).Close
 End Sub

 ' Función para eliminar el archivo de bloqueo por ejecuccion de seccion
 Sub DeleteLockFile(Seccion)
     Set fso = CreateObject("Scripting.FileSystemObject")
     lockFile = fso.BuildPath(GetDataFolder(), "Lock_" & Seccion & ".lock")
     
     ' Elimina el archivo de bloqueo si existe
     If fso.FileExists(lockFile) Then
         fso.DeleteFile(lockFile)
     End If
 End Sub
 
 ' Función para obtener el nombre de la carpeta remota desde el archivo PHP en línea
 Function GetRemoteFolderName(strUrl)
 
     Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
     objXmlHttp.Open "GET", strUrl, False
     objXmlHttp.Send
 
     If objXmlHttp.Status = 200 Then
         GetRemoteFolderName = Trim(objXmlHttp.responseText)
     Else
         GetRemoteFolderName = ""
     End If
 
     Set objXmlHttp = Nothing
 End Function
 
 ' Función para cerrar un programa dado su nombre de archivo
 Sub CloseProgram(programName)
 
     Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
     Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & programName & "'")
 
     For Each objProcess in colProcesses
         objProcess.Terminate()
     Next
 End Sub
 
 ' Definición de la subrutina "DownloadFile"
 Sub DownloadFile(url, destPath)
     Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
     Set objADOStream = CreateObject("ADODB.Stream")
 
     ' Descargar el archivo
     objXMLHTTP.Open "GET", url, False
     objXMLHTTP.Send
 
     ' Guardar el archivo descargado
     If objXMLHTTP.Status = 200 Then
         objADOStream.Type = 1
         objADOStream.Open
         objADOStream.Write objXMLHTTP.ResponseBody
         objADOStream.SaveToFile destPath, 2
         objADOStream.Close
     Else
         MsgBox "No se pudo descargar el archivo. Error: " & objXMLHTTP.Status & " " & objXMLHTTP.statusText, vbCritical, "Error de descarga"
     End If
 
     Set objXMLHTTP = Nothing
     Set objADOStream = Nothing
 End Sub

 ' Edicion de archivo
 Sub EditLaunchCfgFile(filePath, searchStr, replaceStr)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(filePath) Then
        ' Leer el contenido del archivo con codificación UTF-16 LE
        Set objFile = objFSO.OpenTextFile(filePath, 1, False, -1)  ' 1: Lectura, -1: UTF-16 LE
        fileContent = objFile.ReadAll()
        objFile.Close
        
        ' Realizar la sustitución de texto en el contenido UTF-16 LE
        fileContent = Replace(fileContent, searchStr, replaceStr)
        
        ' Crear un objeto ADODB.Stream para guardar el contenido modificado como UTF-16 LE
        Set utf16Stream = CreateObject("ADODB.Stream")
        utf16Stream.Open
        utf16Stream.Charset = "UTF-16"
        utf16Stream.WriteText fileContent
        utf16Stream.SaveToFile filePath, 2  ' 2: Escritura, sobreescribir el archivo
        utf16Stream.Close
    End If
    Set objFSO = Nothing
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ SECCIÓN DE FUNCIONES GENERAL DEL LAUNCHER \\\\\\\\\\\

 ' INSTALA EL LAUNCHER
 Sub SeccionA1()
     If WScript.Arguments.length = 0 Then
     Set objShell = CreateObject("Shell.Application")
     objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
 
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("data/categorias.zip")
     obj.DeleteFile("data/java.zip")

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_8]", "//[p_0_img_button_8]"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_9]", "[p_0_img_button_9]"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_6]", "[p_0_img_button_6]"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_7]", "[p_0_img_button_7]"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_11]", "[p_0_img_button_11]"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_12]", "[p_0_img_button_12]"
 
     texto = "!La instalacion fue exitosa!, Iniciando laucher..."
     MyBox = MsgBox(texto, 266304, "HeavyNight!")
 
     Set WshShell = CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     End If
 End Sub
 
 ' DESINSTALA EL LAUNCHER
 Sub SeccionA2()
     If WScript.Arguments.length = 0 Then
     Set objShell = CreateObject("Shell.Application")
     objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\resources") Then
     result = msgbox("Esta accion eliminara por completo las categorias y no habra vuelta atras. Tardara unos segundos y cuando haya terminado se abrira el launcher nuevamente." & vbCrLf & "" & vbCrLf & "¿Estas seguro?",4+48, "HeavyNiht - Desinstalador")
     If result=6 then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell")
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
     '
     ' Crea un objeto WshShell
     Set wshShell = CreateObject("WScript.Shell")
 
     ' Guarda el directorio actual
     currentDir = wshShell.CurrentDirectory
 
     ' Cambia al directorio C:\Program Files\Java
     wshShell.CurrentDirectory = "C:\Program Files\Java\"
 
     ' Borra las carpetas jdk1.8.0_281 y Jre_8
     Set fso = CreateObject("Scripting.FileSystemObject")
 
     ' Utiliza Try-Catch para evitar el error en caso de que los archivos no se encuentren
     On Error Resume Next
     fso.DeleteFolder "jdk-17.0.6", True ' True indica que se borren las subcarpetas y archivos
     fso.DeleteFolder "Jre_8", True ' True indica que se borren las subcarpetas y archivos
     If Err.Number <> 0 Then
     ' Si se genera un error, muestra un mensaje y continúa con el resto del código
     MsgBox "No se encontraron algunos archivos a eliminar", 0, "HeavyNight!"
     Err.Clear
     End If
     On Error Goto 0
 
     ' Regresa al directorio anterior
     wshShell.CurrentDirectory = currentDir
     '
     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "//[p_0_img_button_8]", "[p_0_img_button_8]"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_9]", "//[p_0_img_button_9]"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_6]", "//[p_0_img_button_6]"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_7]", "//[p_0_img_button_7]"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_11]", "//[p_0_img_button_11]"
     EditLaunchCfgFile "data\launchcfg", "[p_0_img_button_12]", "//[p_0_img_button_12]"
     '
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("launcher\*.exe")
     obj.DeleteFile("launcher\*.dat")
     obj.DeleteFile("launcher\*.json")
     obj.DeleteFile("launcher\*.pak")
     obj.DeleteFile("launcher\*.bin")
     obj.DeleteFile("launcher\*.dll")
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\locales"
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\resources"
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\forge"
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     texto = "Se eliminaron los archivos con exito!. Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     else
     
     end if
     else
     
     End If
     End If
 End Sub
 
 ' ERROR CUANDO NO ENCUENTRE EL ARCHIVO VERSION DEL LAUNCHER
 Sub SeccionA3()
     texto = "!No se pudo obtener el archivo version.txt y es posible que no recibas actualizaciones futuras."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
 End Sub
 
 ' LAUNCHER EN MANTENIMIENTO
 Sub SeccionA4()
     ' Obtener el contenido de la URL PHP
     Set request = CreateObject("MSXML2.XMLHTTP")
     request.Open "GET", "https://www.heavynight.com/launcherV5/Mantenimiento.php", False
     request.Send
     
     ' Obtener el valor de mantenimiento del archivo PHP
     maintenance = Trim(request.responseText)
     
     ' Verificar el valor de mantenimiento
     If LCase(maintenance) = "true" Then
         ' Cerrar el programa HeavyNight.exe si está abierto
         CloseProgram "HeavyNight.exe"
         
     ' Abrir la página web de más información
     respuesta = MsgBox("El launcher esta en mantenimiento." & vbNewLine & vbNewLine & "Te muestro el canal de discord para que te enteres cuando el launcher deja de estar en mantenimiento?", vbYesNo + vbQuestion, "HeavyNight - Mantenimiento!")
     
     ' Si se hace clic en Aceptar, abrir la página web de más información
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "https://discord.com/channels/860007074695610398/1015623724600414218"
     End If
     End If
 End Sub
 
 ' LAUNCHER UPDATE
 Sub SeccionA5()
     Set objShell = CreateObject("WScript.Shell")
     url = "https://www.heavynight.com/changelog/categories/10"
     objShell.Run url
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="HeavyNight.exe.WebView2"
     If fso.FolderExists(FolderDel) Then ' Verificar si la carpeta existe
         fso.DeleteFolder(FolderDel) ' Eliminar la carpeta si existe
     End If
     Set fso=nothing
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ SECCIÓN DE FUNCIONES DEL MODPACK A LA CATEGORIA 1 DEL LAUNCHER \\\\\\\\\\\

 ' INSTALA LA CATEGORIA 1
 Sub SeccionB1()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
     '
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("data/instancia.zip")
     obj.DeleteFile("data/mods.zip")

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_3]", "//[p_6_img_button_3]" '(b_descargar_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_1_img_button_4]", "[p_1_img_button_4]" '(b_juguar_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_1_img_button_11]", "[p_1_img_button_11]" '(b_delete_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_1_img_button_12]", "[p_1_img_button_12]" '(b_parches_a.png)
     '
     texto = "!La instalacion fue exitosa!, Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     Else
     '
     respuesta = MsgBox("Algo salio mal porque no se reconocio la carpeta " & carpeta & ". " & vbCrLf & "" & vbCrLf & "Quieres reportarlo con nuestro soporte?!", vbYesNo + vbQuestion, "Instalacion - " & carpeta & "!")
     If respuesta = vbYes Then
     CreateObject("WScript.Shell").Run "http://heavynight.com/"
     end If
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     End If
 End Sub
 
 ' DESINSTALA LA CATEGORIA 1
 Sub SeccionB2()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     If WScript.Arguments.length = 0 Then
         Set objShell = CreateObject("Shell.Application")
         objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
     
     result = msgbox("Esta accion eliminara por completo la instancia y no habra vuelta atras. Tardara unos segundos y cuando haya terminado se abrira el launcher nuevamente." & vbCrLf & "" & vbCrLf & "¿Estas seguro?",4+48, "HeavyNiht - Desinstalador")
     If result=6 then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
     
     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_3]", "[p_6_img_button_3]" '(b_descargar_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_4]", "//[p_1_img_button_4]" '(b_juguar_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_11]", "//[p_1_img_button_11]" '(b_delete_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_12]", "//[p_1_img_button_12]" '(b_parches_a.png)
     
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\" & carpeta & ""
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     texto = "Se eliminaron los archivos con exito!. Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight - " & carpeta & "!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     else
     
     end if
     else
     
     End If
     End If
 End Sub
 
 ' INICIA LA CATEGORIA 1
 Sub SeccionB3()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
             ip = responseLines(3)
             forge = responseLines(4)
             cjava = responseLines(5)
             cversion = responseLines(1)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la instalacion de la categoria.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FolderExists("launcher\" & carpeta & "") Then
     ' ////Comprovacion en la instalacion de java 17.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists("C:\Program Files\java\jdk-17.0.6\bin\javaw.exe") Then
     
     ' ////Comprovacion de la version de parches.////
       ' Nombre y ruta del archivo de destino
       destPath = "launcher\" & carpeta & "\version.txt"
       
       ' Contenido del archivo
       fileContent = "1.0.0"
       
       ' Crea un objeto FileSystemObject para comprobar si el archivo existe
       Set fs = CreateObject("Scripting.FileSystemObject")
       If Not fs.FileExists(destPath) Then
       
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
       
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
       
       End If
       
       ' Obtener la ruta actual del directorio donde se está ejecutando el script
       Set fso = CreateObject("Scripting.FileSystemObject")
       currentFolder = fso.GetAbsolutePathName(".")
       
       ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
       versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
       
       ' Especificar la ruta completa del archivo version.txt
       versionPath = versionFolderPath & "version.txt"
       
       ' Leer el contenido del archivo version.txt
       Set versionFile = fso.OpenTextFile(versionPath, 1)
       version = versionFile.ReadLine
       versionFile.Close
       
       ' Especificar la URL de la versión del archivo txt
       urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Modpack/" & carpeta & "/version.txt"
       
       ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
       Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
       winHttpReq.Open "GET", urlRemota, False
       winHttpReq.Send
       
       ' Obtener el contenido del archivo de versión desde la URL remota
       remoteVersion = winHttpReq.responseText
       
       ' Comparar la versión obtenida con la versión actual
       If version = remoteVersion Then
       
       ' ////Si la versión coincide, ejecuta la instancia de juego////
         
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
         '
         ' Llamar a la subrutina "DownloadFile"
         DownloadFile "https://www.heavynight.com/launcherV5/launcher_configs.js", "launcher\resources\app\launcher_config.js"

         ' Leer el contenido del archivo descargado
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 1)
         configContent = configFile.ReadAll
         configFile.Close
         
         ' Realizar las sustituciones en el contenido del archivo
         configContent = Replace(configContent, "{category-ip}", ip)
         configContent = Replace(configContent, "{category-name}", carpeta)
         configContent = Replace(configContent, "{category-version}", cversion)
         configContent = Replace(configContent, "{category-forge}", forge)
         configContent = Replace(configContent, "{category-java}", cjava)
         
         ' Guardar el contenido modificado de vuelta al archivo
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 2)
         configFile.Write configContent
         configFile.Close
         '
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c cd launcher & login.exe", 0, False
         
       ' ////Si la versión NO coincide, mostrar una alerta para actualizar el parche////
         Else
 
         'Set objShell = CreateObject("WScript.Shell")
         'link = "https://www.heavynight.com/changelog/categories/4" Reemplaza con tu enlace deseado
         
         'objShell.Run link
     
         result = msgbox("!Hay una actualizacion pendiente!. ¿Quiero actualizarlo?",4+48, "HeavyNiht - " & carpeta & "")
         If result=6 then
         
           '//// Comprueba si tiene el java 8////
             Set fso = CreateObject("Scripting.FileSystemObject")
             If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     
             Set oShell = WScript.CreateObject ("WScript.Shell") 
             oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
             '
             ' Llamar a la subrutina "DownloadFile"
             DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

             ' Leer el contenido del archivo descargado
             Set fso = CreateObject("Scripting.FileSystemObject")
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
             configContent = configFile.ReadAll
             configFile.Close
             
             ' Realizar las sustituciones en el contenido del archivo
             configContent = Replace(configContent, "{category-name}", carpeta)
             
             ' Guardar el contenido modificado de vuelta al archivo
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
             configFile.Write configContent
             configFile.Close
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("launcher\server_sync.exe c1serversync", 1, True)
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("HeavyNight.exe", 1, false)
             '
             texto = "El parche ha terminado."
             MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
             '//// Si no tiene java 8////
     
             Else
     
             MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la integracion de java del launcher. Por favor, contacta con nuestro soporte o reinstale el launcher.", vbCritical + vbSystemModal, "Error de inicio"
             respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
             
             If respuesta = vbYes Then
             CreateObject("WScript.Shell").Run "http://heavynight.com/"
     
             Else
     
             '///DIJISTE QUE NO AL CONTACTAR AL SOPORTE Y CIERRA EL PROCESO///'
     
             End if
     
             End if
     
         Else
         
         '///DIJISTE QUE NO Y CIERRA EL PROCESO///'
         
         End If
         
         End If
     ' ////Final de la comprovacion en la instalacion de java 17.////
       Else
       MsgBox "" & carpeta & " necesita Java 17 y parece que algo ha fallado en la integracion de java. Por favor, contacta con nuestro soporte o vuelva a reinstalar el launcher.", vbCritical + vbSystemModal, "Error de inicio"
       respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
       If respuesta = vbYes Then
       CreateObject("WScript.Shell").Run "http://heavynight.com/"
       end If
       End if
     ' ////Fianal de la comprovacion en la instalacion de la categoria.////
       Else
       texto = "Aun no tienes descargado " & carpeta & "."
       MyBox = MsgBox(texto,266304,"HeavyNight!")
       end if
 End Sub
 
 ' NOTIFICACION DE ACTUALIZACIONES DEL MODPACK 1
 Sub SeccionB4()
  url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
  lineNumber = 0 ' La primera línea
  
  Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
  xmlhttp.Open "GET", url, False
  xmlhttp.Send
  
  If xmlhttp.Status = 200 Then
      responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
      If UBound(responseLines) >= lineNumber Then
          carpeta = responseLines(lineNumber)
      Else
          MsgBox "La línea solicitada no existe en la respuesta."
          WScript.Quit ' Sale del script si ocurre un error en la obtención de la carpeta
      End If
  Else
      MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
      WScript.Quit ' Sale del script si ocurre un error en la obtención de la URL
  End If
  
  ' Nombre y ruta del archivo de destino
  destPath = "launcher\" & carpeta & "\version.txt"
  
  ' Crea un objeto FileSystemObject para comprobar si el archivo existe
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(destPath) Then
      WScript.Quit ' Sale del script si el archivo version.txt no existe
  End If
  
  ' Obtener la ruta actual del directorio donde se está ejecutando el script
  Set fso = CreateObject("Scripting.FileSystemObject")
  currentFolder = fso.GetAbsolutePathName(".")
  
  ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
  versionFolderPath = currentFolder & "\launcher\" & carpeta & "\version.txt"
  
  ' Leer el contenido del archivo version.txt
  Set versionFile = fso.OpenTextFile(versionFolderPath, 1)
  version = versionFile.ReadLine
  versionFile.Close
  
  ' Especificar la URL de la versión del archivo PHP
  urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Modpack/" & carpeta & "/version.txt"
  
  ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
  Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
  winHttpReq.Open "GET", urlRemota, False
  winHttpReq.Send
  
  ' Obtener el contenido del archivo de versión desde la URL remota
  remoteVersion = winHttpReq.responseText
  
  ' Comparar la versión obtenida con la versión actual
  If version = remoteVersion Then
      WScript.Quit ' Sale del script si las versiones coinciden
  Else
      respuesta = MsgBox("Hay una nueva actualización del modpack " & carpeta & ". " & vbCrLf & "" & vbCrLf & "Quieres ver los cambios que se han hecho?", vbYesNo + vbQuestion, "HeavyNight")
      If respuesta = vbYes Then
          CreateObject("WScript.Shell").Run "https://www.heavynight.com/changelog/categories/4"
      End If
  End If
 
 End Sub
 
 ' PARCHA LA EL MODPACK DE LA CATEGORIA 1
 Sub SeccionB5()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
 
     ' Nombre y ruta del archivo de destino
     destPath = "launcher\" & carpeta & "\version.txt"
     
     ' Contenido del archivo
     fileContent = "1.0.0"
     
     ' Crea un objeto FileSystemObject para comprobar si el archivo existe
     Set fs = CreateObject("Scripting.FileSystemObject")
     If Not fs.FileExists(destPath) Then
     
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
     
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
     
     End If
     
     ' Obtener la ruta actual del directorio donde se está ejecutando el script
     Set fso = CreateObject("Scripting.FileSystemObject")
     currentFolder = fso.GetAbsolutePathName(".")
     
     ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
     versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
     
     ' Especificar la ruta completa del archivo version.txt
     versionPath = versionFolderPath & "version.txt"
     
     ' Leer el contenido del archivo version.txt
     Set versionFile = fso.OpenTextFile(versionPath, 1)
     version = versionFile.ReadLine
     versionFile.Close
     
     ' Especificar la URL de la versión del archivo PHP
     urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Modpack/" & carpeta & "/version.txt"
     
 
     ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
     Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
     winHttpReq.Open "GET", urlRemota, False
     winHttpReq.Send
     
     ' Obtener el contenido del archivo de versión desde la URL remota
     remoteVersion = winHttpReq.responseText
     
     ' Comparar la versión obtenida con la versión actual
     If version = remoteVersion Then
     
     ' Si la versión coincide, continuar con el codigo.
     
     result = msgbox("!Ya tienes la ultima actualizacion!. ¿Quiero actualizarlo igualmente?",4+48, "HeavyNiht - " & carpeta & "")
     If result=6 then
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c1serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     else
     
     end if
     else
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"
     
     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close

     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c1serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     
     End If
 End Sub
 
 ' UPDATE DE LA INSTANCIA CATEGORIA 1
 Sub SeccionB6()
     ' Cambiar esta ruta al nombre del archivo de texto local
     strLocalFilePath = "data/categorias.txt"
     
     ' Crear un objeto FileSystemObject
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     
     ' Verificar si el archivo local existe
     If objFSO.FileExists(strLocalFilePath) Then
         ' Abrir el archivo y leer su contenido
         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
         categoriavieja = objFile.ReadLine
         
         ' No olvides cerrar el archivo cuando hayas terminado de usarlo
         objFile.Close
     End If
     
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             nuevacategoria = responseLines(lineNumber) ' Obtener la nueva categoría de la cuarta línea
     
                 ' Convertir los nombres de las carpetas a minúsculas antes de comparar
                 If LCase(categoriavieja) <> LCase(nuevacategoria) Then
                     ' Aquí puede agregar el código que desea ejecutar cuando los nombres no coinciden
                         carpetaViejaPath = "launcher\" & categoriavieja
                         If objFSO.FolderExists(carpetaViejaPath) Then
                             result = MsgBox("Hemos marcado la categoria " & categoriavieja & " como 'CERRADA' ya que hay una nueva disponible actualmente llamada " & nuevacategoria & "." & vbCrLf & "" & vbCrLf & "Quieres actualizar a la nueva categoria?", 4+48, "HeavyNight - Categorias")
                             If result = 6 Then
                                 result = MsgBox("Quieres hacer una copia de seguridad de tus archivos guardados en " & categoriavieja & " antes de actualizar?", 4+48, "HeavyNight - Categorias")
                                 If result = 6 Then
                                     Set oShell = CreateObject("WScript.Shell")
                                     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                       
                                     strFolder = "launcher\" & categoriavieja & ""
                                     strDestFolder = "launcher\zboveda"
                                     strNewFolderName = nuevacategoria
                       
                                     If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                         objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                     End If
                       
                                     If Not objFSO.FolderExists(strDestFolder) Then
                                         objFSO.CreateFolder strDestFolder
                                     End If
                       
                                     If objFSO.FolderExists(strFolder) Then
                                         sourceFolderName = objFSO.GetFolder(strFolder).Name
                                         destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                         objFSO.CreateFolder destFolder
                       
                                         ' Mover el contenido de la carpeta de origen a la carpeta de destino
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "config")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "config"), objFSO.BuildPath(destFolder, "config")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "mods")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "mods"), objFSO.BuildPath(destFolder, "mods")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "saves")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "saves"), objFSO.BuildPath(destFolder, "saves")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "scripts")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "scripts"), objFSO.BuildPath(destFolder, "scripts")
                                         End If
                                         If objFSO.FileExists(objFSO.BuildPath(strFolder, "version.txt")) Then
                                             objFSO.MoveFile objFSO.BuildPath(strFolder, "version.txt"), objFSO.BuildPath(destFolder, "version.txt")
                                         End If
                       
                                         ' Renombrar la carpeta de origen al nuevo nombre
                                         objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                       
                                         'Editar el archivo "launchcfg"
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_12]", "//[p_1_img_button_12]" '(b_parches_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_11]", "//[p_1_img_button_11]" '(b_delete_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_4]", "//[p_1_img_button_4]" '(b_jugar_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "//[p_1_img_button_13]", "[p_1_img_button_13]" '(b_instalar1_a.png)
 
                                         ' Descargar los archivos de instalar nueva categoría
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/logo.png", "data\logo-C1.png")
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/titulo.png", "data\titulo-C1.png")
                       
                                         ' Crear un nuevo ArrayList
                                         Set lines = CreateObject("System.Collections.ArrayList")
                       
                                         ' Abrir el archivo para leer
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                                         ' Leer todas las líneas en el ArrayList
                                         Do Until objFile.AtEndOfStream
                                             lines.Add objFile.ReadLine
                                         Loop
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         ' Cambia la segunda línea al nombre de la nueva categoría
                                         lines.Item(0) = nuevacategoria
                       
                                         ' Abre el archivo para escribir
                                         Set objFSO = CreateObject("Scripting.FileSystemObject")
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                                         ' Escribe todas las líneas en el archivo
                                         For Each line in lines
                                             objFile.WriteLine line
                                         Next
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         Set WshShell = CreateObject("WScript.Shell")
                                         Return = WshShell.Run("HeavyNight.exe", 1, False)
                       
                                         MsgBox "La categoria " & categoriavieja & " ha cambiado y se ha creado una copia en su bóveda para presentarte a nuestra nueva categoría " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                     Else
                                         ' No hacer nada en caso de responder a no.
                                     End If
                       
                                     Set objFSO = Nothing
                                 Else
                                     Set oShell = CreateObject("WScript.Shell")
                                     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                       
                                     strFolder = "launcher\" & categoriavieja & ""
                                     strDestFolder = "launcher\zboveda"
                                     strNewFolderName = nuevacategoria
                       
                                     If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                         objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                     End If
                       
                                     If Not objFSO.FolderExists(strDestFolder) Then
                                         objFSO.CreateFolder strDestFolder
                                     End If
                       
                                     If objFSO.FolderExists(strFolder) Then
                                         sourceFolderName = objFSO.GetFolder(strFolder).Name
                                         destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                         objFSO.CreateFolder destFolder
                       
                                         arrFolders = Array("config", "mods", "saves", "scripts")
                       
                                         For Each subFolder In arrFolders
                                             FolderDel = "launcher\" & categoriavieja & "\" & subFolder
                                             ' Verificar si la carpeta existe antes de intentar eliminarla
                                             If objFSO.FolderExists(FolderDel) Then
                                                 objFSO.DeleteFolder(FolderDel)
                                             End If
                                         Next
                       
                                         ' Renombrar la carpeta de origen al nuevo nombre
                                         objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                       
                                         'Editar el archivo "launchcfg"
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_12]", "//[p_1_img_button_12]" '(b_parches_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_11]", "//[p_1_img_button_11]" '(b_delete_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_1_img_button_4]", "//[p_1_img_button_4]" '(b_jugar_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "//[p_1_img_button_13]", "[p_1_img_button_13]" '(b_instalar1_a.png)

                                          ' Descargar los archivos de instalar nueva categoría
                                          Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/logo.png", "data\logo-C1.png")
                                          Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/titulo.png", "data\titulo-C1.png")
                       
                                         ' Crear un nuevo ArrayList
                                         Set lines = CreateObject("System.Collections.ArrayList")
                       
                                         ' Abrir el archivo para leer
                                         Set objFSO = CreateObject("Scripting.FileSystemObject")
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                                         ' Leer todas las líneas en el ArrayList
                                         Do Until objFile.AtEndOfStream
                                             lines.Add objFile.ReadLine
                                         Loop
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         ' Cambia la segunda línea al nombre de la nueva categoría
                                         lines.Item(0) = nuevacategoria
                       
                                         ' Abre el archivo para escribir
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                                         ' Escribe todas las líneas en el archivo
                                         For Each line in lines
                                             objFile.WriteLine line
                                         Next
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         Set WshShell = CreateObject("WScript.Shell")
                                         Return = WshShell.Run("HeavyNight.exe", 1, False)
                       
                                         MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                     End If
                                 End If
                             End If
                         Else
                             ' Descargar los archivos de instalar nueva categoría
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/logo.png", "data\logo-C1.png")
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria1/imagenes/titulo.png", "data\titulo-C1.png")
 
                             ' Crear un nuevo ArrayList
                             Set lines = CreateObject("System.Collections.ArrayList")
                       
                             ' Abrir el archivo para leer
                             Set objFSO = CreateObject("Scripting.FileSystemObject")
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                             ' Leer todas las líneas en el ArrayList
                             Do Until objFile.AtEndOfStream
                                 lines.Add objFile.ReadLine
                             Loop
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             ' Cambia la segunda línea al nombre de la nueva categoría
                             lines.Item(0) = nuevacategoria
                       
                             ' Abre el archivo para escribir
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                             ' Escribe todas las líneas en el archivo
                             For Each line in lines
                                 objFile.WriteLine line
                             Next
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                         End If
                     End If
                 End If
         End If
 End Sub
 
 ' ABRE LA CARPETA MODS DE LA CATEGORIA 1
 Sub SeccionB7()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La tercera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la carpeta de la categoria.////
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "") Then
         Set objShell = CreateObject("WScript.Shell")
         objShell.Run "explorer.exe """ & "launcher\" & carpeta & """", 1, False
     Else
         MsgBox "Parece que aún no tienes instalada la categoría o no existe la carpeta " & carpeta & "."
     End If
 End Sub
 
 ' ABRE LA WEB INFO DE LA CATEGORIA 1
 Sub SeccionB8()
     url = "https://raw.githubusercontent.com/heavysproject/Categoria-1-Modpack/main/Category-Name%20(1).php"
     lineNumber = 2 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
     Dim responseText
         responseLines = Split(responseText, vbNewLine) ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If

     MsgBox "respuesta: [" & paginaweb & "]"
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/news/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub
 
 ' ABRE LA WEB TIENDA DE LA CATEGORIA 1
 Sub SeccionB9()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/shop/categories/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub

 ' COPIAR IP DE LA CATEGORIA 1
 Sub SeccionB10()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria1/Category-Name.php"
     lineNumber = 3 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             ipserver = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If

     ' Crea un objeto Shell para acceder al portapapeles
     Set objShell = CreateObject("WScript.Shell")
     
     ' Copia el texto al portapapeles
     objShell.Run "cmd /c echo " & ipserver & "| clip", 2, True
     
     ' Muestra un mensaje para indicar que se ha copiado el texto
     texto = "La IP ha sido copiado en tu portapapeles"
     MyBox = MsgBox(texto,266304,"HeavyNight!")
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ SECCIÓN DE FUNCIONES DEL MODPACK A LA CATEGORIA 2 DEL LAUNCHER \\\\\\\\\\\

 ' INSTALA LA CATEGORIA 2
 Sub SeccionC1()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La segunda linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
     '
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("data/instancia.zip")
     obj.DeleteFile("data/mods.zip")

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_5]", "//[p_2_img_button_5]" '(b_descargar_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_4]", "[p_2_img_button_4]" '(b_jugar2_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_10]", "[p_2_img_button_10]" '(b_delete2_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_1]", "[p_2_img_button_1]" '(b_parches2_a.png)
     '
     texto = "!La instalacion fue exitosa!, Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     Else
     '
     respuesta = MsgBox("Algo salio mal porque no se reconocio la carpeta " & carpeta & ". " & vbCrLf & "" & vbCrLf & "Quieres reportarlo con nuestro soporte?!", vbYesNo + vbQuestion, "Instalacion - " & carpeta & "!")
     If respuesta = vbYes Then
     CreateObject("WScript.Shell").Run "http://heavynight.com/"
     end If
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     End If
 End Sub
 
 ' DESINSTALA LA CATEGORIA 2
 Sub SeccionC2()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La segunda linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     If WScript.Arguments.length = 0 Then
         Set objShell = CreateObject("Shell.Application")
         objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
 
     result = msgbox("Esta accion eliminara por completo la instancia y no habra vuelta atras. Tardara unos segundos y cuando haya terminado se abrira el launcher nuevamente." & vbCrLf & "" & vbCrLf & "¿Estas seguro?",4+48, "HeavyNight - Desinstalador C2")
     If result=6 then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
     '
     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_5]", "[p_2_img_button_5]" '(b_descargar_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_4]", "//[p_2_img_button_4]" '(b_jugar2_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_10]", "//[p_2_img_button_10]" '(b_delete2_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_1]", "//[p_2_img_button_1]" '(b_parches2_a.png)
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\" & carpeta & ""
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     texto = "Se eliminaron los archivos con exito!. Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight! - " & carpeta & "")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     else
     
     end if
     else
     End If
     
     End If
 End Sub
 
 ' INICIA LA CATEGORIA 2
 Sub SeccionC3()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
             ip = responseLines(3)
             forge = responseLines(4)
             cjava = responseLines(5)
             cversion = responseLines(1)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la instalacion de la categoria.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FolderExists("launcher\" & carpeta & "") Then
     ' ////Comprovacion en la instalacion de java 17.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists("C:\Program Files\java\jdk-17.0.6\bin\javaw.exe") Then
     
     ' ////Comprovacion de la version de parches.////
       ' Nombre y ruta del archivo de destino
       destPath = "launcher\" & carpeta & "\version.txt"
       
       ' Contenido del archivo
       fileContent = "1.0.0"
       
       ' Crea un objeto FileSystemObject para comprobar si el archivo existe
       Set fs = CreateObject("Scripting.FileSystemObject")
       If Not fs.FileExists(destPath) Then
       
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
       
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
       
       End If
       
       ' Obtener la ruta actual del directorio donde se está ejecutando el script
       Set fso = CreateObject("Scripting.FileSystemObject")
       currentFolder = fso.GetAbsolutePathName(".")
       
       ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
       versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
       
       ' Especificar la ruta completa del archivo version.txt
       versionPath = versionFolderPath & "version.txt"
       
       ' Leer el contenido del archivo version.txt
       Set versionFile = fso.OpenTextFile(versionPath, 1)
       version = versionFile.ReadLine
       versionFile.Close
       
       ' Especificar la URL de la versión del archivo PHP
       urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Modpack/" & carpeta & "/version.txt"
       
       ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
       Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
       winHttpReq.Open "GET", urlRemota, False
       winHttpReq.Send
       
       ' Obtener el contenido del archivo de versión desde la URL remota
       remoteVersion = winHttpReq.responseText
       
       ' Comparar la versión obtenida con la versión actual
       If version = remoteVersion Then
       
       ' ////Si la versión coincide, ejecuta la instancia de juego////
         
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
         '
         ' Llamar a la subrutina "DownloadFile"
         DownloadFile "https://www.heavynight.com/launcherV5/launcher_configs.js", "launcher\resources\app\launcher_config.js"

         ' Leer el contenido del archivo descargado
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 1)
         configContent = configFile.ReadAll
         configFile.Close
         
         ' Realizar las sustituciones en el contenido del archivo
         configContent = Replace(configContent, "{category-ip}", ip)
         configContent = Replace(configContent, "{category-name}", carpeta)
         configContent = Replace(configContent, "{category-version}", cversion)
         configContent = Replace(configContent, "{category-forge}", forge)
         configContent = Replace(configContent, "{category-java}", cjava)
         
         ' Guardar el contenido modificado de vuelta al archivo
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 2)
         configFile.Write configContent
         configFile.Close
         '
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c cd launcher & login.exe", 0, False
         
       ' ////Si la versión NO coincide, mostrar una alerta para actualizar el parche////
         Else
     
     
         result = msgbox("!Hay una actualizacion pendiente!. ¿Quiero actualizarlo?",4+48, "HeavyNiht - " & carpeta & "")
         If result=6 then
         
           '//// Comprueba si tiene el java 8////
             Set fso = CreateObject("Scripting.FileSystemObject")
             If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     
             Set oShell = WScript.CreateObject ("WScript.Shell") 
             oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
             '
             ' Llamar a la subrutina "DownloadFile"
             DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

             ' Leer el contenido del archivo descargado
             Set fso = CreateObject("Scripting.FileSystemObject")
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
             configContent = configFile.ReadAll
             configFile.Close
             
             ' Realizar las sustituciones en el contenido del archivo
             configContent = Replace(configContent, "{category-name}", carpeta)
             
             ' Guardar el contenido modificado de vuelta al archivo
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
             configFile.Write configContent
             configFile.Close
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("launcher\server_sync.exe c2serversync", 1, True)
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("HeavyNight.exe", 1, false)
             '
             texto = "El parche ha terminado."
             MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
             
             '//// Si no tiene java 8////
     
             Else
     
             MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la integracion de java del launcher. Por favor, contacta con nuestro soporte o reinstale el launcher.", vbCritical + vbSystemModal, "Error de inicio"
             respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
             
             If respuesta = vbYes Then
             CreateObject("WScript.Shell").Run "http://heavynight.com/"
     
             Else
     
             '///DIJISTE QUE NO AL CONTACTAR AL SOPORTE Y CIERRA EL PROCESO///'
     
             End if
     
             End if
     
         Else
         
         '///DIJISTE QUE NO Y CIERRA EL PROCESO///'
         
         End If
         
         End If
     ' ////Final de la comprovacion en la instalacion de java 17.////
       Else
       MsgBox "" & carpeta & " necesita Java 17 y parece que algo ha fallado en la integracion de java. Por favor, contacta con nuestro soporte o vuelva a reinstalar el launcher.", vbCritical + vbSystemModal, "Error de inicio"
       respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
       If respuesta = vbYes Then
       CreateObject("WScript.Shell").Run "http://heavynight.com/"
       end If
       End if
     ' ////Fianal de la comprovacion en la instalacion de la categoria.////
       Else
       texto = "Aun no tienes descargado " & carpeta & "."
       MyBox = MsgBox(texto,266304,"HeavyNight!")
       end if
 End Sub
 
 ' NOTIFICACION DE ACTUALIZACIONES DEL MODPACK 2
 Sub SeccionC4()
  url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
  lineNumber = 0 ' La primera línea
  
  Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
  xmlhttp.Open "GET", url, False
  xmlhttp.Send
  
  If xmlhttp.Status = 200 Then
      responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
      If UBound(responseLines) >= lineNumber Then
          carpeta = responseLines(lineNumber)
      Else
          MsgBox "La línea solicitada no existe en la respuesta."
          WScript.Quit ' Sale del script si ocurre un error en la obtención de la carpeta
      End If
  Else
      MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
      WScript.Quit ' Sale del script si ocurre un error en la obtención de la URL
  End If
  
  ' Nombre y ruta del archivo de destino
  destPath = "launcher\" & carpeta & "\version.txt"
  
  ' Crea un objeto FileSystemObject para comprobar si el archivo existe
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(destPath) Then
      MsgBox "El archivo version.txt no existe. No se hace nada."
      WScript.Quit ' Sale del script si el archivo version.txt no existe
  End If
  
  ' Obtener la ruta actual del directorio donde se está ejecutando el script
  Set fso = CreateObject("Scripting.FileSystemObject")
  currentFolder = fso.GetAbsolutePathName(".")
  
  ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
  versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
  
  ' Especificar la ruta completa del archivo version.txt
  versionPath = versionFolderPath & "version.txt"
  
  ' Leer el contenido del archivo version.txt
  Set versionFile = fso.OpenTextFile(versionPath, 1)
  version = versionFile.ReadLine
  versionFile.Close
  
  ' Especificar la URL de la versión del archivo PHP
  urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Modpack/" & carpeta & "/version.txt"
  
  ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
  Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
  winHttpReq.Open "GET", urlRemota, False
  winHttpReq.Send
  
  ' Obtener el contenido del archivo de versión desde la URL remota
  remoteVersion = winHttpReq.responseText
  
  ' Comparar la versión obtenida con la versión actual
  If version = remoteVersion Then
      WScript.Quit ' Sale del script si las versiones coinciden
  Else
      respuesta = MsgBox("Hay una nueva actualización del modpack " & carpeta & ". " & vbCrLf & "" & vbCrLf & "¿Quieres ver los cambios que se han hecho?", vbYesNo + vbQuestion, "Instalación - " & carpeta & "!")
      If respuesta = vbYes Then
          CreateObject("WScript.Shell").Run "https://www.heavynight.com/changelog/categories/4"
      End If
  End If
 
 End Sub
 
 ' PARCHA EL MODPACK DE LA CATEGORIA 2
 Sub SeccionC5()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La segunda linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' Nombre y ruta del archivo de destino
     destPath = "launcher\" & carpeta & "\version.txt"
     
     ' Contenido del archivo
     fileContent = "1.0.0"
     
     ' Crea un objeto FileSystemObject para comprobar si el archivo existe
     Set fs = CreateObject("Scripting.FileSystemObject")
     If Not fs.FileExists(destPath) Then
     
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
     
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
     
     End If
     
     ' Obtener la ruta actual del directorio donde se está ejecutando el script
     Set fso = CreateObject("Scripting.FileSystemObject")
     currentFolder = fso.GetAbsolutePathName(".")
     
     ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
     versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
     
     ' Especificar la ruta completa del archivo version.txt
     versionPath = versionFolderPath & "version.txt"
     
     ' Leer el contenido del archivo version.txt
     Set versionFile = fso.OpenTextFile(versionPath, 1)
     version = versionFile.ReadLine
     versionFile.Close
     
     ' Especificar la URL de la versión del archivo PHP
     urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Modpack/" & carpeta & "/version.txt"
     
     ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
     Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
     winHttpReq.Open "GET", urlRemota, False
     winHttpReq.Send
     
     ' Obtener el contenido del archivo de versión desde la URL remota
     remoteVersion = winHttpReq.responseText
     
     ' Comparar la versión obtenida con la versión actual
     If version = remoteVersion Then
     
     ' Si la versión coincide, continuar con el codigo.
 
     result = msgbox("!Ya tienes la ultima actualizacion!. ¿Quiero actualizarlo igualmente?",4+48, "HeavyNiht - " & carpeta & "")
     If result=6 then
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c2serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     else
     
     end if
     else
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c2serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     
     End If
 End Sub
 
 ' UPDATE DE LA INSTANCIA CATEGORIA 2
 Sub SeccionC6()
     ' Cambiar esta ruta al nombre del archivo de texto local
     strLocalFilePath = "data/categorias.txt"
     
     ' Crear un objeto FileSystemObject
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     
     ' Verificar si el archivo local existe
     If objFSO.FileExists(strLocalFilePath) Then
         ' Abrir el archivo y leer su contenido
         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
         objFile.ReadLine' descarta la primera línea
         categoriavieja = objFile.ReadLine
         
         ' No olvides cerrar el archivo cuando hayas terminado de usarlo
         objFile.Close
     End If
     
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La primera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             nuevacategoria = responseLines(lineNumber) ' Obtener la nueva categoría de la cuarta línea
     
                 ' Convertir los nombres de las carpetas a minúsculas antes de comparar
                 If LCase(categoriavieja) <> LCase(nuevacategoria) Then
                     ' Aquí puede agregar el código que desea ejecutar cuando los nombres no coinciden
                         carpetaViejaPath = "launcher\" & categoriavieja
                         If objFSO.FolderExists(carpetaViejaPath) Then
                             result = MsgBox("Hemos marcado la categoria " & categoriavieja & " como 'CERRADA' ya que hay una nueva disponible actualmente llamada " & nuevacategoria & "." & vbCrLf & "" & vbCrLf & "Quieres actualizar a la nueva categoria?", 4+48, "HeavyNight - Categorias")
                             If result = 6 Then
                                 result = MsgBox("Quieres hacer una copia de seguridad de tus archivos guardados en " & categoriavieja & " antes de actualizar?", 4+48, "HeavyNight - Categorias")
                                 If result = 6 Then
                                     Set oShell = CreateObject("WScript.Shell")
                                     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                       
                                     strFolder = "launcher\" & categoriavieja & ""
                                     strDestFolder = "launcher\zboveda"
                                     strNewFolderName = nuevacategoria
                       
                                     If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                         objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                     End If
                       
                                     If Not objFSO.FolderExists(strDestFolder) Then
                                         objFSO.CreateFolder strDestFolder
                                     End If
                       
                                     If objFSO.FolderExists(strFolder) Then
                                         sourceFolderName = objFSO.GetFolder(strFolder).Name
                                         destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                         objFSO.CreateFolder destFolder
                       
                                         ' Mover el contenido de la carpeta de origen a la carpeta de destino
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "config")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "config"), objFSO.BuildPath(destFolder, "config")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "mods")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "mods"), objFSO.BuildPath(destFolder, "mods")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "saves")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "saves"), objFSO.BuildPath(destFolder, "saves")
                                         End If
                                         If objFSO.FolderExists(objFSO.BuildPath(strFolder, "scripts")) Then
                                             objFSO.MoveFolder objFSO.BuildPath(strFolder, "scripts"), objFSO.BuildPath(destFolder, "scripts")
                                         End If
                                         If objFSO.FileExists(objFSO.BuildPath(strFolder, "version.txt")) Then
                                             objFSO.MoveFile objFSO.BuildPath(strFolder, "version.txt"), objFSO.BuildPath(destFolder, "version.txt")
                                         End If
                       
                                         ' Renombrar la carpeta de origen al nuevo nombre
                                         objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                       
                                         'Editar el archivo "launchcfg"
                                          EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_1]", "//[p_2_img_button_1]" '(b_parches2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "p_2_img_button_10]", "//p_2_img_button_10]" '(b_delete2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_4]", "//[p_2_img_button_4]" '(b_jugar2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_11]", "[p_2_img_button_11]" '(b_instalar2_a.png)
 
                                         ' Descargar los archivos de instalar nueva categoría
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/logo.png", "data\logo-C2.png")
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/titulo.png", "data\titulo-C2.png")
                       
                                         ' Crear un nuevo ArrayList
                                         Set lines = CreateObject("System.Collections.ArrayList")
                       
                                         ' Abrir el archivo para leer
                                         Set objFSO = CreateObject("Scripting.FileSystemObject")
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                                         ' Leer todas las líneas en el ArrayList
                                         Do Until objFile.AtEndOfStream
                                             lines.Add objFile.ReadLine
                                         Loop
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         ' Cambia la segunda línea al nombre de la nueva categoría
                                         lines.Item(1) = nuevacategoria
                       
                                         ' Abre el archivo para escribir
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                                         ' Escribe todas las líneas en el archivo
                                         For Each line in lines
                                             objFile.WriteLine line
                                         Next
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         Set WshShell = CreateObject("WScript.Shell")
                                         Return = WshShell.Run("HeavyNight.exe", 1, False)
                       
                                         MsgBox "La categoria " & categoriavieja & " ha cambiado y se ha creado una copia en su bóveda para presentarte a nuestra nueva categoría " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                     Else
                                         ' No hacer nada en caso de responder a no.
                                     End If
                       
                                     Set objFSO = Nothing
                                 Else
                                     Set oShell = CreateObject("WScript.Shell")
                                     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                       
                                     strFolder = "launcher\" & categoriavieja & ""
                                     strDestFolder = "launcher\zboveda"
                                     strNewFolderName = nuevacategoria
                       
                                     If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                         objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                     End If
                       
                                     If Not objFSO.FolderExists(strDestFolder) Then
                                         objFSO.CreateFolder strDestFolder
                                     End If
                       
                                     If objFSO.FolderExists(strFolder) Then
                                         sourceFolderName = objFSO.GetFolder(strFolder).Name
                                         destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                         objFSO.CreateFolder destFolder
                       
                                         arrFolders = Array("config", "mods", "saves", "scripts")
                       
                                         For Each subFolder In arrFolders
                                             FolderDel = "launcher\" & categoriavieja & "\" & subFolder
                                             ' Verificar si la carpeta existe antes de intentar eliminarla
                                             If objFSO.FolderExists(FolderDel) Then
                                                 objFSO.DeleteFolder(FolderDel)
                                             End If
                                         Next
                       
                                         ' Renombrar la carpeta de origen al nuevo nombre
                                         objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                       
                                         'Editar el archivo "launchcfg"
                                          EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_1]", "//[p_2_img_button_1]" '(b_parches2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "p_2_img_button_10]", "//p_2_img_button_10]" '(b_delete2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "[p_2_img_button_4]", "//[p_2_img_button_4]" '(b_jugar2_a.png)
                                          EditLaunchCfgFile "data\launchcfg", "//[p_2_img_button_11]", "[p_2_img_button_11]" '(b_instalar2_a.png)
 
                                         ' Descargar los archivos de instalar nueva categoría
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/logo.png", "data\logo-C2.png")
                                         Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/titulo.png", "data\titulo-C2.png")
                       
                                         ' Crear un nuevo ArrayList
                                         Set lines = CreateObject("System.Collections.ArrayList")
                       
                                         ' Abrir el archivo para leer
                                         Set objFSO = CreateObject("Scripting.FileSystemObject")
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                                         ' Leer todas las líneas en el ArrayList
                                         Do Until objFile.AtEndOfStream
                                             lines.Add objFile.ReadLine
                                         Loop
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         ' Cambia la segunda línea al nombre de la nueva categoría
                                         lines.Item(1) = nuevacategoria
                       
                                         ' Abre el archivo para escribir
                                         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                                         ' Escribe todas las líneas en el archivo
                                         For Each line in lines
                                             objFile.WriteLine line
                                         Next
                       
                                         ' Cierra el archivo
                                         objFile.Close
                       
                                         Set WshShell = CreateObject("WScript.Shell")
                                         Return = WshShell.Run("HeavyNight.exe", 1, False)
                       
                                         MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                     End If
                                 End If
                             End If
                         Else
                             ' Descargar los archivos de instalar nueva categoría
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/logo.png", "data\logo-C2.png")
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria2/imagenes/titulo.png", "data\titulo-C2.png")
                             ' Crear un nuevo ArrayList
                             Set lines = CreateObject("System.Collections.ArrayList")
                       
                             ' Abrir el archivo para leer
                             Set objFSO = CreateObject("Scripting.FileSystemObject")
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                             ' Leer todas las líneas en el ArrayList
                             Do Until objFile.AtEndOfStream
                                 lines.Add objFile.ReadLine
                             Loop
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             ' Cambia la segunda línea al nombre de la nueva categoría
                             lines.Item(1) = nuevacategoria
                       
                             ' Abre el archivo para escribir
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                             ' Escribe todas las líneas en el archivo
                             For Each line in lines
                                 objFile.WriteLine line
                             Next
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                         End If
                     End If
                 End If
             End If
 End Sub
 
 ' ABRE LA CARPETA MODS DE LA CATEGORIA 2
 Sub SeccionC7()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La tercera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la carpeta de la categoria.////
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "") Then
         Set objShell = CreateObject("WScript.Shell")
         objShell.Run "explorer.exe """ & "launcher\" & carpeta & """", 1, False
     Else
         MsgBox "Parece que aún no tienes instalada la categoría o no existe la carpeta " & carpeta & "."
     End If
 End Sub
 
 ' ABRE LA WEB INFO DE LA CATEGORIA 2
 Sub SeccionC8()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/news/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub
 
 ' ABRE LA WEB TIENDA DE LA CATEGORIA 2
 Sub SeccionC9()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/shop/categories/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub

 ' COPIAR IP DE LA CATEGORIA 2
 Sub SeccionC10()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria2/Category-Name.php"
     lineNumber = 3 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             ipserver = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If

     ' Crea un objeto Shell para acceder al portapapeles
     Set objShell = CreateObject("WScript.Shell")
     
     ' Copia el texto al portapapeles
     objShell.Run "cmd /c echo " & ipserver & "| clip", 2, True
     
     ' Muestra un mensaje para indicar que se ha copiado el texto
     texto = "La IP ha sido copiado en tu portapapeles"
     MyBox = MsgBox(texto,266304,"HeavyNight!")
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ SECCIÓN DE FUNCIONES DEL MODPACK A LA CATEGORIA 3 DEL LAUNCHER \\\\\\\\\\\

 ' INSTALA LA CATEGORIA 3
 Sub SeccionD1()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
     '
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("data/instancia.zip")
     obj.DeleteFile("data/mods.zip")

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_7]", "//[p_3_img_button_7]" '(b_descargar3_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_8]", "[p_3_img_button_8]" '(b_jugar3_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_10]", "[p_3_img_button_10]" '(b_delete3_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_4]", "[p_3_img_button_4]" '(b_parches3_a.png)
     '
     texto = "!La instalacion fue exitosa!, Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     Else
     '
     respuesta = MsgBox("Algo salio mal porque no se reconocio la carpeta " & carpeta & ". " & vbCrLf & "" & vbCrLf & "Quieres reportarlo con nuestro soporte?!", vbYesNo + vbQuestion, "Instalacion - " & carpeta & "!")
     If respuesta = vbYes Then
     CreateObject("WScript.Shell").Run "http://heavynight.com/"
     end If
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     End If
 End Sub
 
 ' DESINSTALA LA CATEGORIA 3
 Sub SeccionD2()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     If WScript.Arguments.length = 0 Then
         Set objShell = CreateObject("Shell.Application")
         objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
 
     result = msgbox("Esta accion eliminara por completo la instancia y no habra vuelta atras. Tardara unos segundos y cuando haya terminado se abrira el launcher nuevamente." & vbCrLf & "" & vbCrLf & "¿Estas seguro?",4+48, "HeavyNiht - Desinstalador")
     If result=6 then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
     '
     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_7]", "[p_3_img_button_7]" '(b_descargar3_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_8]", "//[p_3_img_button_8]" '(b_jugar3_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_10]", "//[p_3_img_button_10]" '(b_delete3_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_4]", "//[p_3_img_button_4]" '(b_parches3_a.png)
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\" & carpeta & ""
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     texto = "Se eliminaron los archivos con exito!. Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight - " & carpeta & "!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     else
     
     end if
     else
     End If
     
     End If
 End Sub
 
 ' INICIA LA CATEGORIA 3
 Sub SeccionD3()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
             ip = responseLines(3)
             forge = responseLines(4)
             cjava = responseLines(5)
             cversion = responseLines(1)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la instalacion de la categoria.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FolderExists("launcher\" & carpeta & "") Then
     ' ////Comprovacion en la instalacion de java 17.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists("C:\Program Files\java\jdk-17.0.6\bin\javaw.exe") Then
     
     ' ////Comprovacion de la version de parches.////
       ' Nombre y ruta del archivo de destino
       destPath = "launcher\" & carpeta & "\version.txt"
       
       ' Contenido del archivo
       fileContent = "1.0.0"
       
       ' Crea un objeto FileSystemObject para comprobar si el archivo existe
       Set fs = CreateObject("Scripting.FileSystemObject")
       If Not fs.FileExists(destPath) Then
       
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
       
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
       
       End If
       
       ' Obtener la ruta actual del directorio donde se está ejecutando el script
       Set fso = CreateObject("Scripting.FileSystemObject")
       currentFolder = fso.GetAbsolutePathName(".")
       
       ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
       versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
       
       ' Especificar la ruta completa del archivo version.txt
       versionPath = versionFolderPath & "version.txt"
       
       ' Leer el contenido del archivo version.txt
       Set versionFile = fso.OpenTextFile(versionPath, 1)
       version = versionFile.ReadLine
       versionFile.Close
       
       ' Especificar la URL de la versión del archivo PHP
       urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Modpack/" & carpeta & "/version.txt"
       
       ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
       Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
       winHttpReq.Open "GET", urlRemota, False
       winHttpReq.Send
       
       ' Obtener el contenido del archivo de versión desde la URL remota
       remoteVersion = winHttpReq.responseText
       
       ' Comparar la versión obtenida con la versión actual
       If version = remoteVersion Then
       
       ' ////Si la versión coincide, ejecuta la instancia de juego////
         
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
         '
         ' Llamar a la subrutina "DownloadFile"
         DownloadFile "https://www.heavynight.com/launcherV5/launcher_configs.js", "launcher\resources\app\launcher_config.js"

         ' Leer el contenido del archivo descargado
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 1)
         configContent = configFile.ReadAll
         configFile.Close
         
         ' Realizar las sustituciones en el contenido del archivo
         configContent = Replace(configContent, "{category-ip}", ip)
         configContent = Replace(configContent, "{category-name}", carpeta)
         configContent = Replace(configContent, "{category-version}", cversion)
         configContent = Replace(configContent, "{category-forge}", forge)
         configContent = Replace(configContent, "{category-java}", cjava)
         
         ' Guardar el contenido modificado de vuelta al archivo
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 2)
         configFile.Write configContent
         configFile.Close
         '
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c cd launcher & login.exe", 0, False
         
       ' ////Si la versión NO coincide, mostrar una alerta para actualizar el parche////
         Else
     
         result = msgbox("!Hay una actualizacion pendiente!. ¿Quiero actualizarlo?",4+48, "HeavyNiht - " & carpeta & "")
         If result=6 then
         
           '//// Comprueba si tiene el java 8////
             Set fso = CreateObject("Scripting.FileSystemObject")
             If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     
             Set oShell = WScript.CreateObject ("WScript.Shell") 
             oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
             '
             ' Llamar a la subrutina "DownloadFile"
             DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

             ' Leer el contenido del archivo descargado
             Set fso = CreateObject("Scripting.FileSystemObject")
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
             configContent = configFile.ReadAll
             configFile.Close
             
             ' Realizar las sustituciones en el contenido del archivo
             configContent = Replace(configContent, "{category-name}", carpeta)
             
             ' Guardar el contenido modificado de vuelta al archivo
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
             configFile.Write configContent
             configFile.Close
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("launcher\server_sync.exe c3serversync", 1, True)
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("HeavyNight.exe", 1, false)
             '
             texto = "El parche ha terminado."
             MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
             
             '//// Si no tiene java 8////
     
             Else
     
             MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la integracion de java del launcher. Por favor, contacta con nuestro soporte o reinstale el launcher.", vbCritical + vbSystemModal, "Error de inicio"
             respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
             
             If respuesta = vbYes Then
             CreateObject("WScript.Shell").Run "http://heavynight.com/"
     
             Else
     
             '///DIJISTE QUE NO AL CONTACTAR AL SOPORTE Y CIERRA EL PROCESO///'
     
             End if
     
             End if
     
         Else
         
         '///DIJISTE QUE NO Y CIERRA EL PROCESO///'
         
         End If
         
         End If
     ' ////Final de la comprovacion en la instalacion de java 17.////
       Else
       MsgBox "" & carpeta & " necesita Java 17 y parece que algo ha fallado en la integracion de java. Por favor, contacta con nuestro soporte o vuelva a reinstalar el launcher.", vbCritical + vbSystemModal, "Error de inicio"
       respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
       If respuesta = vbYes Then
       CreateObject("WScript.Shell").Run "http://heavynight.com/"
       end If
       End if
     ' ////Fianal de la comprovacion en la instalacion de la categoria.////
       Else
       texto = "Aun no tienes descargado " & carpeta & "."
       MyBox = MsgBox(texto,266304,"HeavyNight!")
       end if
 End Sub
 
 ' NOTIFICACION DE ACTUALIZACIONES DEL MODPACK 3
 Sub SeccionD4()
  url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
  lineNumber = 0 ' La primera línea
  
  Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
  xmlhttp.Open "GET", url, False
  xmlhttp.Send
  
  If xmlhttp.Status = 200 Then
      responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
      If UBound(responseLines) >= lineNumber Then
          carpeta = responseLines(lineNumber)
      Else
          MsgBox "La línea solicitada no existe en la respuesta."
          WScript.Quit ' Sale del script si ocurre un error en la obtención de la carpeta
      End If
  Else
      MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
      WScript.Quit ' Sale del script si ocurre un error en la obtención de la URL
  End If
  
  ' Nombre y ruta del archivo de destino
  destPath = "launcher\" & carpeta & "\version.txt"
  
  ' Crea un objeto FileSystemObject para comprobar si el archivo existe
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(destPath) Then
      MsgBox "El archivo version.txt no existe. No se hace nada."
      WScript.Quit ' Sale del script si el archivo version.txt no existe
  End If
  
  ' Obtener la ruta actual del directorio donde se está ejecutando el script
  Set fso = CreateObject("Scripting.FileSystemObject")
  currentFolder = fso.GetAbsolutePathName(".")
  
  ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
  versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
  
  ' Especificar la ruta completa del archivo version.txt
  versionPath = versionFolderPath & "version.txt"
  
  ' Leer el contenido del archivo version.txt
  Set versionFile = fso.OpenTextFile(versionPath, 1)
  version = versionFile.ReadLine
  versionFile.Close
  
  ' Especificar la URL de la versión del archivo PHP
  urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Modpack/" & carpeta & "/version.txt"
  
  ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
  Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
  winHttpReq.Open "GET", urlRemota, False
  winHttpReq.Send
  
  ' Obtener el contenido del archivo de versión desde la URL remota
  remoteVersion = winHttpReq.responseText
  
  ' Comparar la versión obtenida con la versión actual
  If version = remoteVersion Then
      WScript.Quit ' Sale del script si las versiones coinciden
  Else
      respuesta = MsgBox("Hay una nueva actualización del modpack " & carpeta & ". " & vbCrLf & "" & vbCrLf & "¿Quieres ver los cambios que se han hecho?", vbYesNo + vbQuestion, "Instalación - " & carpeta & "!")
      If respuesta = vbYes Then
          CreateObject("WScript.Shell").Run "https://www.heavynight.com/changelog/categories/4"
      End If
  End If
 
 End Sub
 
 ' PARCHA EL MOPDPACK DE LA CATEGORIA 3
 Sub SeccionD5()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
 
     ' Nombre y ruta del archivo de destino
     destPath = "launcher\" & carpeta & "\version.txt"
     
     ' Contenido del archivo
     fileContent = "1.0.0"
     
     ' Crea un objeto FileSystemObject para comprobar si el archivo existe
     Set fs = CreateObject("Scripting.FileSystemObject")
     If Not fs.FileExists(destPath) Then
     
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
     
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
     
     End If
     
     ' Obtener la ruta actual del directorio donde se está ejecutando el script
     Set fso = CreateObject("Scripting.FileSystemObject")
     currentFolder = fso.GetAbsolutePathName(".")
     
     ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
     versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
     
     ' Especificar la ruta completa del archivo version.txt
     versionPath = versionFolderPath & "version.txt"
     
     ' Leer el contenido del archivo version.txt
     Set versionFile = fso.OpenTextFile(versionPath, 1)
     version = versionFile.ReadLine
     versionFile.Close
     
     ' Especificar la URL de la versión del archivo PHP
     urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Modpack/" & carpeta & "/version.txt"
     
     ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
     Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
     winHttpReq.Open "GET", urlRemota, False
     winHttpReq.Send
     
     ' Obtener el contenido del archivo de versión desde la URL remota
     remoteVersion = winHttpReq.responseText
     
     ' Comparar la versión obtenida con la versión actual
     If version = remoteVersion Then
     
     ' Si la versión coincide, continuar con el codigo.
 
     result = msgbox("!Ya tienes la ultima actualizacion!. ¿Quiero actualizarlo igualmente?",4+48, "HeavyNiht - " & carpeta & "")
     If result=6 then
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c3serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     else
     
     end if
     else
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c3serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     
     End If
 End Sub
 
 ' UPDATE DE LA INSTANCIA CATEGORIA 3
 Sub SeccionD6()
     ' Cambiar esta ruta al nombre del archivo de texto local
     strLocalFilePath = "data/categorias.txt"
     
     ' Crear un objeto FileSystemObject
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     
     ' Verificar si el archivo local existe
     If objFSO.FileExists(strLocalFilePath) Then
         ' Abrir el archivo y leer su contenido
         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
         objFile.ReadLine' descarta la primera línea
         objFile.ReadLine' descarta la segunda línea
         categoriavieja = objFile.ReadLine
         
         ' No olvides cerrar el archivo cuando hayas terminado de usarlo
         objFile.Close
 
     End If
     
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La primera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             nuevacategoria = responseLines(lineNumber) ' Obtener la nueva categoría de la cuarta línea
     
                 ' Convertir los nombres de las carpetas a minúsculas antes de comparar
                 If LCase(categoriavieja) <> LCase(nuevacategoria) Then
                     ' Aquí puede agregar el código que desea ejecutar cuando los nombres no coinciden
                         carpetaViejaPath = "launcher\" & categoriavieja
                         If objFSO.FolderExists(carpetaViejaPath) Then
                             result = MsgBox("Hemos marcado la categoria " & categoriavieja & " como 'CERRADA' ya que hay una nueva disponible actualmente llamada " & nuevacategoria & "." & vbCrLf & "" & vbCrLf & "Quieres actualizar a la nueva categoria?", 4+48, "HeavyNight - Categorias")
                             If result = 6 Then
                                     result = MsgBox("Quieres hacer una copia de seguridad de tus archivos guardados en " & categoriavieja & " antes de actualizar?", 4+48, "HeavyNight - Categorias")
                                     If result = 6 Then
                                         Set oShell = CreateObject("WScript.Shell")
                                         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                           
                                         strFolder = "launcher\" & categoriavieja & ""
                                         strDestFolder = "launcher\zboveda"
                                         strNewFolderName = nuevacategoria
                           
                                         If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                             objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                         End If
                           
                                         If Not objFSO.FolderExists(strDestFolder) Then
                                             objFSO.CreateFolder strDestFolder
                                         End If
                           
                                         If objFSO.FolderExists(strFolder) Then
                                             sourceFolderName = objFSO.GetFolder(strFolder).Name
                                             destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                             objFSO.CreateFolder destFolder
                           
                                             ' Mover el contenido de la carpeta de origen a la carpeta de destino
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "config")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "config"), objFSO.BuildPath(destFolder, "config")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "mods")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "mods"), objFSO.BuildPath(destFolder, "mods")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "saves")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "saves"), objFSO.BuildPath(destFolder, "saves")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "scripts")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "scripts"), objFSO.BuildPath(destFolder, "scripts")
                                             End If
                                             If objFSO.FileExists(objFSO.BuildPath(strFolder, "version.txt")) Then
                                                 objFSO.MoveFile objFSO.BuildPath(strFolder, "version.txt"), objFSO.BuildPath(destFolder, "version.txt")
                                             End If
                           
                                             ' Renombrar la carpeta de origen al nuevo nombre
                                             objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                           
                                             'Editar el archivo "launchcfg"
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_4]", "//[p_3_img_button_4]" '(b_parches3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_10]", "//[p_3_img_button_10]" '(b_delete3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_8]", "//[p_3_img_button_8]" '(b_jugar3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_9]", "[p_3_img_button_9]" '(b_instalar3_a.png)
    
                                             ' Descargar los archivos de instalar nueva categoría
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/logo.png", "data\logo-C3.png")
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/titulo.png", "data\titulo-C3.png")
                           
                                             ' Crear un nuevo ArrayList
                                             Set lines = CreateObject("System.Collections.ArrayList")
                           
                                             ' Abrir el archivo para leer
                                             Set objFSO = CreateObject("Scripting.FileSystemObject"
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                           
                                             ' Leer todas las líneas en el ArrayList
                                             Do Until objFile.AtEndOfStream
                                                 lines.Add objFile.ReadLine
                                             Loop
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             ' Cambia la tercera línea al nombre de la nueva categoría
                                             lines.Item(2) = nuevacategoria
                           
                                             ' Abre el archivo para escribir
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                           
                                             ' Escribe todas las líneas en el archivo
                                             For Each line in lines
                                                 objFile.WriteLine line
                                             Next
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             Set WshShell = CreateObject("WScript.Shell")
                                             Return = WshShell.Run("HeavyNight.exe", 1, False)
                           
                                             MsgBox "La categoria " & categoriavieja & " ha cambiado y se ha creado una copia en su bóveda para presentarte a nuestra nueva categoría " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                         Else
                                             ' No hacer nada en caso de responder a no.
                                         End If
                                     Else
                                         Set oShell = CreateObject("WScript.Shell")
                                         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                           
                                         strFolder = "launcher\" & categoriavieja & ""
                                         strDestFolder = "launcher\zboveda"
                                         strNewFolderName = nuevacategoria
                           
                                         If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                             objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                         End If
                           
                                         If Not objFSO.FolderExists(strDestFolder) Then
                                             objFSO.CreateFolder strDestFolder
                                         End If
                           
                                         If objFSO.FolderExists(strFolder) Then
                                             sourceFolderName = objFSO.GetFolder(strFolder).Name
                                             destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                             objFSO.CreateFolder destFolder
                           
                                             arrFolders = Array("config", "mods", "saves", "scripts")
                           
                                             For Each subFolder In arrFolders
                                                 FolderDel = "launcher\" & categoriavieja & "\" & subFolder
                                                 ' Verificar si la carpeta existe antes de intentar eliminarla
                                                 If objFSO.FolderExists(FolderDel) Then
                                                     objFSO.DeleteFolder(FolderDel)
                                                 End If
                                             Next
                           
                                             ' Renombrar la carpeta de origen al nuevo nombre
                                             objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                           
                                             'Editar el archivo "launchcfg"
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_4]", "//[p_3_img_button_4]" '(b_parches3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_10]", "//[p_3_img_button_10]" '(b_delete3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_3_img_button_8]", "//[p_3_img_button_8]" '(b_jugar3_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "//[p_3_img_button_9]", "[p_3_img_button_9]" '(b_instalar3_a.png)
    
                                             ' Descargar los archivos de instalar nueva categoría
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/logo.png", "data\logo-C3.png")
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/titulo.png", "data\titulo-C3.png")
                           
                                             ' Crear un nuevo ArrayList
                                             Set lines = CreateObject("System.Collections.ArrayList")
                           
                                             ' Abrir el archivo para leer
                                             Set objFSO = CreateObject("Scripting.FileSystemObject"
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                           
                                             ' Leer todas las líneas en el ArrayList
                                             Do Until objFile.AtEndOfStream
                                                 lines.Add objFile.ReadLine
                                             Loop
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             ' Cambia la tercera línea al nombre de la nueva categoría
                                             lines.Item(2) = nuevacategoria
                           
                                             ' Abre el archivo para escribir
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                           
                                             ' Escribe todas las líneas en el archivo
                                             For Each line in lines
                                                 objFile.WriteLine line
                                             Next
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             Set WshShell = CreateObject("WScript.Shell")
                                             Return = WshShell.Run("HeavyNight.exe", 1, False)
                           
                                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                         End If
                                     End If
                                 End If
                         Else
                             ' Descargar los archivos de instalar nueva categoría
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/logo.png", "data\logo-C3.png")
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria3/imagenes/titulo.png", "data\titulo-C3.png")
 
                             ' Crear un nuevo ArrayList
                             Set lines = CreateObject("System.Collections.ArrayList")
                       
                             ' Abrir el archivo para leer
                             Set objFSO = CreateObject("Scripting.FileSystemObject"
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                             ' Leer todas las líneas en el ArrayList
                             Do Until objFile.AtEndOfStream
                                 lines.Add objFile.ReadLine
                             Loop
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             ' Cambia la tercera línea al nombre de la nueva categoría
                             lines.Item(2) = nuevacategoria
                       
                             ' Abre el archivo para escribir
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                             ' Escribe todas las líneas en el archivo
                             For Each line in lines
                                 objFile.WriteLine line
                             Next
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                         End If
                     End If
                 End If
             End If
 End Sub
 
 ' ABRE LA CARPETA MODS DE LA CATEGORIA 3
 Sub SeccionD7()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La tercera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la carpeta de la categoria.////
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "") Then
         Set objShell = CreateObject("WScript.Shell")
         objShell.Run "explorer.exe """ & "launcher\" & carpeta & """", 1, False
     Else
         MsgBox "Parece que aún no tienes instalada la categoría o no existe la carpeta " & carpeta & "."
     End If
 End Sub
 
 ' ABRE LA WEB INFO DE LA CATEGORIA 3
 Sub SeccionD8()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/news/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub
 
 ' ABRE LA WEB TIENDA DE LA CATEGORIA 3
 Sub SeccionD9()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/shop/categories/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub

 ' COPIAR IP DE LA CATEGORIA 3
 Sub SeccionD10()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria3/Category-Name.php"
     lineNumber = 3 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             ipserver = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If

     ' Crea un objeto Shell para acceder al portapapeles
     Set objShell = CreateObject("WScript.Shell")
     
     ' Copia el texto al portapapeles
     objShell.Run "cmd /c echo " & ipserver & "| clip", 2, True
     
     ' Muestra un mensaje para indicar que se ha copiado el texto
     texto = "La IP ha sido copiado en tu portapapeles"
     MyBox = MsgBox(texto,266304,"HeavyNight!")
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ SECCIÓN DE FUNCIONES DEL MODPACK A LA CATEGORIA 4 DEL LAUNCHER \\\\\\\\\\\

 ' INSTALA LA CATEGORIA 4
 Sub SeccionE1()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
     '
     Set obj = CreateObject("Scripting.FileSystemObject")
     obj.DeleteFile("data/instancia.zip")
     obj.DeleteFile("data/mods.zip")

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_3]", "//[p_6_img_button_3]" '(b_descargar4_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_11]", "[p_6_img_button_11]" '(b_jugar4_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_10]", "[p_6_img_button_10]" '(b_parches4_a.png)
     EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_6]", "[p_6_img_button_6]" '(b_delete4_a.png)
     
     '
     texto = "!La instalacion fue exitosa!, Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     Else
     '
     respuesta = MsgBox("Algo salio mal porque no se reconocio la carpeta " & carpeta & ". " & vbCrLf & "" & vbCrLf & "Quieres reportarlo con nuestro soporte?!", vbYesNo + vbQuestion, "Instalacion - " & carpeta & "!")
     If respuesta = vbYes Then
     CreateObject("WScript.Shell").Run "http://heavynight.com/"
     end If
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, False)
     End If
 End Sub
 
 ' DESINSTALA LA CATEGORIA 4
 Sub SeccionE2()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     If WScript.Arguments.length = 0 Then
         Set objShell = CreateObject("Shell.Application")
         objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
     Else
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "\assets") Then
 
     result = msgbox("Esta accion eliminara por completo la instancia y no habra vuelta atras. Tardara unos segundos y cuando haya terminado se abrira el launcher nuevamente." & vbCrLf & "" & vbCrLf & "Estas seguro?",4+48, "HeavyNight - Desinstalador")
     If result=6 then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True

     'Editar el archivo "launchcfg"
     EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_3]", "[p_6_img_button_3]" '(b_descargar4_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_11]", "//[p_6_img_button_11]" '(b_jugar4_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_10]", "//[p_6_img_button_10]" '(b_parches4_a.png)
     EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_6]", "//[p_6_img_button_6]" '(b_delete4_a.png)
     '
     Set fso=createobject("Scripting.FileSystemObject")
     FolderDel="launcher\" & carpeta & ""
     fso.DeleteFolder(FolderDel)
     Set fso=nothing
     '
     texto = "Se eliminaron los archivos con exito!. Abriendo launcher..."
     MyBox = MsgBox(texto,266304,"HeavyNight - " & carpeta & "!")
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     else
     
     end if
     else
     End If
     
     End If
 End Sub
 
 ' INICIA LA CATEGORIA 4
 Sub SeccionE3()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
             ip = responseLines(3)
             forge = responseLines(4)
             cjava = responseLines(5)
             cversion = responseLines(1)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la instalacion de la categoria.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FolderExists("launcher\" & carpeta & "") Then
     ' ////Comprovacion en la instalacion de java 17.////
       Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists("C:\Program Files\java\jdk-17.0.6\bin\javaw.exe") Then
     
     ' ////Comprovacion de la version de parches.////
       ' Nombre y ruta del archivo de destino
       destPath = "launcher\" & carpeta & "\version.txt"
       
       ' Contenido del archivo
       fileContent = "1.0.0"
       
       ' Crea un objeto FileSystemObject para comprobar si el archivo existe
       Set fs = CreateObject("Scripting.FileSystemObject")
       If Not fs.FileExists(destPath) Then
       
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
       
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
       
       End If
       
       ' Obtener la ruta actual del directorio donde se está ejecutando el script
       Set fso = CreateObject("Scripting.FileSystemObject")
       currentFolder = fso.GetAbsolutePathName(".")
       
       ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
       versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
       
       ' Especificar la ruta completa del archivo version.txt
       versionPath = versionFolderPath & "version.txt"
       
       ' Leer el contenido del archivo version.txt
       Set versionFile = fso.OpenTextFile(versionPath, 1)
       version = versionFile.ReadLine
       versionFile.Close
       
       ' Especificar la URL de la versión del archivo PHP
       urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Modpack/" & carpeta & "/version.txt"
       
       ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
       Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
       winHttpReq.Open "GET", urlRemota, False
       winHttpReq.Send
       
       ' Obtener el contenido del archivo de versión desde la URL remota
       remoteVersion = winHttpReq.responseText
       
       ' Comparar la versión obtenida con la versión actual
       If version = remoteVersion Then
       
       ' ////Si la versión coincide, ejecuta la instancia de juego////
         
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
         '
         ' Llamar a la subrutina "DownloadFile"
         DownloadFile "https://www.heavynight.com/launcherV5/launcher_configs.js", "launcher\resources\app\launcher_config.js"

         ' Leer el contenido del archivo descargado
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 1)
         configContent = configFile.ReadAll
         configFile.Close
         
         ' Realizar las sustituciones en el contenido del archivo
         configContent = Replace(configContent, "{category-ip}", ip)
         configContent = Replace(configContent, "{category-name}", carpeta)
         configContent = Replace(configContent, "{category-version}", cversion)
         configContent = Replace(configContent, "{category-java}", cjava)
         configContent = Replace(configContent, "{category-forge}", forge)
         
         ' Guardar el contenido modificado de vuelta al archivo
         Set configFile = fso.OpenTextFile("launcher\resources\app\launcher_config.js", 2)
         configFile.Write configContent
         configFile.Close
         '
         Set oShell = WScript.CreateObject ("WScript.Shell") 
         oShell.Run "cmd /c cd launcher & login.exe", 0, False
         
       ' ////Si la versión NO coincide, mostrar una alerta para actualizar el parche////
         Else
     
         result = msgbox("!Hay una actualizacion pendiente!. ¿Quiero actualizarlo?",4+48, "HeavyNiht - " & carpeta & "")
         If result=6 then
         
           '//// Comprueba si tiene el java 8////
             Set fso = CreateObject("Scripting.FileSystemObject")
             If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     
             Set oShell = WScript.CreateObject ("WScript.Shell") 
             oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
             '
             ' Llamar a la subrutina "DownloadFile"
             DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

             ' Leer el contenido del archivo descargado
             Set fso = CreateObject("Scripting.FileSystemObject")
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
             configContent = configFile.ReadAll
             configFile.Close
             
             ' Realizar las sustituciones en el contenido del archivo
             configContent = Replace(configContent, "{category-name}", carpeta)
             
             ' Guardar el contenido modificado de vuelta al archivo
             Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
             configFile.Write configContent
             configFile.Close
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("launcher\server_sync.exe c4serversync", 1, True)
             '
             Set WshShell = WScript.CreateObject("WScript.Shell")
             Return = WshShell.Run("HeavyNight.exe", 1, false)
             '
             texto = "El parche ha terminado."
             MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
             
             '//// Si no tiene java 8////
     
             Else
     
             MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la integracion de java del launcher. Por favor, contacta con nuestro soporte o reinstale el launcher.", vbCritical + vbSystemModal, "Error de inicio"
             respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
             
             If respuesta = vbYes Then
             CreateObject("WScript.Shell").Run "http://heavynight.com/"
     
             Else
     
             '///DIJISTE QUE NO AL CONTACTAR AL SOPORTE Y CIERRA EL PROCESO///'
     
             End if
     
             End if
     
         Else
         
         '///DIJISTE QUE NO Y CIERRA EL PROCESO///'
         
         End If
         
         End If
     ' ////Final de la comprovacion en la instalacion de java 17.////
       Else
       MsgBox "" & carpeta & " necesita Java 17 y parece que algo ha fallado en la integracion de java. Por favor, contacta con nuestro soporte o vuelva a reinstalar el launcher.", vbCritical + vbSystemModal, "Error de inicio"
       respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
       If respuesta = vbYes Then
       CreateObject("WScript.Shell").Run "http://heavynight.com/"
       end If
       End if
     ' ////Fianal de la comprovacion en la instalacion de la categoria.////
       Else
       texto = "Aun no tienes descargado " & carpeta & "."
       MyBox = MsgBox(texto,266304,"HeavyNight!")
       end if
 End Sub
 
 ' NOTIFICACION DE ACTUALIZACIONES DEL MODPACK 4
 Sub SeccionE4()
  url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
  lineNumber = 0 ' La primera línea
  
  Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
  xmlhttp.Open "GET", url, False
  xmlhttp.Send
  
  If xmlhttp.Status = 200 Then
      responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
      If UBound(responseLines) >= lineNumber Then
          carpeta = responseLines(lineNumber)
      Else
          MsgBox "La línea solicitada no existe en la respuesta."
          WScript.Quit ' Sale del script si ocurre un error en la obtención de la carpeta
      End If
  Else
      MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
      WScript.Quit ' Sale del script si ocurre un error en la obtención de la URL
  End If
  
  ' Nombre y ruta del archivo de destino
  destPath = "launcher\" & carpeta & "\version.txt"
  
  ' Crea un objeto FileSystemObject para comprobar si el archivo existe
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(destPath) Then
      MsgBox "El archivo version.txt no existe. No se hace nada."
      WScript.Quit ' Sale del script si el archivo version.txt no existe
  End If
  
  ' Obtener la ruta actual del directorio donde se está ejecutando el script
  Set fso = CreateObject("Scripting.FileSystemObject")
  currentFolder = fso.GetAbsolutePathName(".")
  
  ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
  versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
  
  ' Especificar la ruta completa del archivo version.txt
  versionPath = versionFolderPath & "version.txt"
  
  ' Leer el contenido del archivo version.txt
  Set versionFile = fso.OpenTextFile(versionPath, 1)
  version = versionFile.ReadLine
  versionFile.Close
  
  ' Especificar la URL de la versión del archivo PHP
  urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Modpack/" & carpeta & "/version.txt"
  
  ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
  Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
  winHttpReq.Open "GET", urlRemota, False
  winHttpReq.Send
  
  ' Obtener el contenido del archivo de versión desde la URL remota
  remoteVersion = winHttpReq.responseText
  
  ' Comparar la versión obtenida con la versión actual
  If version = remoteVersion Then
      WScript.Quit ' Sale del script si las versiones coinciden
  Else
      respuesta = MsgBox("Hay una nueva actualización del modpack " & carpeta & ". " & vbCrLf & "" & vbCrLf & "¿Quieres ver los cambios que se han hecho?", vbYesNo + vbQuestion, "Instalación - " & carpeta & "!")
      If respuesta = vbYes Then
          CreateObject("WScript.Shell").Run "https://www.heavynight.com/changelog/categories/4"
      End If
  End If
 
 End Sub
 
 ' PARCHA EL MOPDPACK DE LA CATEGORIA 4
 Sub SeccionE5()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La tercera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
 
     ' Nombre y ruta del archivo de destino
     destPath = "launcher\" & carpeta & "\version.txt"
     
     ' Contenido del archivo
     fileContent = "1.0.0"
     
     ' Crea un objeto FileSystemObject para comprobar si el archivo existe
     Set fs = CreateObject("Scripting.FileSystemObject")
     If Not fs.FileExists(destPath) Then
     
       ' Crea un objeto FileSystemObject para crear el archivo
       Set file = fs.CreateTextFile(destPath, True)
     
       ' Escribe el contenido en el archivo
       file.Write fileContent
       file.Close
     
     End If
     
     ' Obtener la ruta actual del directorio donde se está ejecutando el script
     Set fso = CreateObject("Scripting.FileSystemObject")
     currentFolder = fso.GetAbsolutePathName(".")
     
     ' Especificar la ruta de la carpeta donde se encuentra el archivo version.txt
     versionFolderPath = currentFolder & "\launcher\" & carpeta & "\"
     
     ' Especificar la ruta completa del archivo version.txt
     versionPath = versionFolderPath & "version.txt"
     
     ' Leer el contenido del archivo version.txt
     Set versionFile = fso.OpenTextFile(versionPath, 1)
     version = versionFile.ReadLine
     versionFile.Close
     
     ' Especificar la URL de la versión del archivo PHP
     urlRemota = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Modpack/" & carpeta & "/version.txt"
     
     ' Crear un objeto WinHttpRequest para hacer la solicitud a la URL remota
     Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
     winHttpReq.Open "GET", urlRemota, False
     winHttpReq.Send
     
     ' Obtener el contenido del archivo de versión desde la URL remota
     remoteVersion = winHttpReq.responseText
     
     ' Comparar la versión obtenida con la versión actual
     If version = remoteVersion Then
     
     ' Si la versión coincide, continuar con el codigo.
 
     result = msgbox("!Ya tienes la ultima actualizacion!. ¿Quiero actualizarlo igualmente?",4+48, "HeavyNiht - " & carpeta & "")
     If result=6 then
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c3serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     else
     
     end if
     else
     
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists("C:\Program Files\java\Jre_8\bin\javaw.exe") Then
     '
     Set oShell = WScript.CreateObject ("WScript.Shell") 
     oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, False
     '
     ' Llamar a la subrutina "DownloadFile"
     DownloadFile "https://www.heavynight.com/launcherV5/config_sync.json", "launcher\config_sync.json"

     ' Leer el contenido del archivo descargado
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 1)
     configContent = configFile.ReadAll
     configFile.Close
     
     ' Realizar las sustituciones en el contenido del archivo
     configContent = Replace(configContent, "{category-name}", carpeta)
     
     ' Guardar el contenido modificado de vuelta al archivo
     Set configFile = fso.OpenTextFile("launcher\config_sync.json", 2)
     configFile.Write configContent
     configFile.Close
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("launcher\server_sync.exe c4serversync", 1, True)
     '
     Set WshShell = WScript.CreateObject("WScript.Shell")
     Return = WshShell.Run("HeavyNight.exe", 1, false)
     '
     texto = "El parche ha terminado."
     MyBox = MsgBox(texto,266304,"HeavyNight - Parches")
     '
     Else
     MsgBox "Esta función necesita Java 8 y parece que algo ha fallado en la instalación de las categorías. Por favor, contacta con nuestro soporte.", vbCritical + vbSystemModal, "Error de inicio"
     
     respuesta = MsgBox("¿Quieres contactar con nuestro soporte?", vbYesNo + vbQuestion, "HeavyNight - Soporte")
     
     If respuesta = vbYes Then
         CreateObject("WScript.Shell").Run "http://heavynight.com/"
     End If
     End If
     
     End If
 End Sub
 
 ' UPDATE DE LA INSTANCIA CATEGORIA 4
 Sub SeccionE6()
     ' Cambiar esta ruta al nombre del archivo de texto local
     strLocalFilePath = "data/categorias.txt"
     
     ' Crear un objeto FileSystemObject
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     
     ' Verificar si el archivo local existe
     If objFSO.FileExists(strLocalFilePath) Then
         ' Abrir el archivo y leer su contenido
         Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
         objFile.ReadLine' descarta la primera línea
         objFile.ReadLine' descarta la segunda línea
         objFile.ReadLine' descarta la tercera línea
         categoriavieja = objFile.ReadLine
         
         ' No olvides cerrar el archivo cuando hayas terminado de usarlo
         objFile.Close
 
     End If
     
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La primera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             nuevacategoria = responseLines(lineNumber) ' Obtener la nueva categoría de la cuarta línea
     
                 ' Convertir los nombres de las carpetas a minúsculas antes de comparar
                 If LCase(categoriavieja) <> LCase(nuevacategoria) Then
                     ' Aquí puede agregar el código que desea ejecutar cuando los nombres no coinciden
                         carpetaViejaPath = "launcher\" & categoriavieja
                         If objFSO.FolderExists(carpetaViejaPath) Then
                             result = MsgBox("Hemos marcado la categoria " & categoriavieja & " como 'CERRADA' ya que hay una nueva disponible actualmente llamada " & nuevacategoria & "." & vbCrLf & "" & vbCrLf & "Quieres actualizar a la nueva categoria?", 4+48, "HeavyNight - Categorias")
                             If result = 6 Then
                                     result = MsgBox("Quieres hacer una copia de seguridad de tus archivos guardados en " & categoriavieja & " antes de actualizar?", 4+48, "HeavyNight - Categorias")
                                     If result = 6 Then
                                         Set oShell = CreateObject("WScript.Shell")
                                         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                           
                                         strFolder = "launcher\" & categoriavieja & ""
                                         strDestFolder = "launcher\zboveda"
                                         strNewFolderName = nuevacategoria
                           
                                         If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                             objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                         End If
                           
                                         If Not objFSO.FolderExists(strDestFolder) Then
                                             objFSO.CreateFolder strDestFolder
                                         End If
                           
                                         If objFSO.FolderExists(strFolder) Then
                                             sourceFolderName = objFSO.GetFolder(strFolder).Name
                                             destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                             objFSO.CreateFolder destFolder
                           
                                             ' Mover el contenido de la carpeta de origen a la carpeta de destino
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "config")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "config"), objFSO.BuildPath(destFolder, "config")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "mods")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "mods"), objFSO.BuildPath(destFolder, "mods")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "saves")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "saves"), objFSO.BuildPath(destFolder, "saves")
                                             End If
                                             If objFSO.FolderExists(objFSO.BuildPath(strFolder, "scripts")) Then
                                                 objFSO.MoveFolder objFSO.BuildPath(strFolder, "scripts"), objFSO.BuildPath(destFolder, "scripts")
                                             End If
                                             If objFSO.FileExists(objFSO.BuildPath(strFolder, "version.txt")) Then
                                                 objFSO.MoveFile objFSO.BuildPath(strFolder, "version.txt"), objFSO.BuildPath(destFolder, "version.txt")
                                             End If
                           
                                             ' Renombrar la carpeta de origen al nuevo nombre
                                             objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)

                                             'Editar el archivo "launchcfg"
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_10]", "//[p_6_img_button_10]" '(b_parches4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_6]", "//[p_6_img_button_6]" '(b_delete4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_11]", "//[p_6_img_button_11]" '(b_jugar4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_12]", "[p_6_img_button_12]" '(b_instalar4_a.png)

                                             ' Descargar los archivos de instalar nueva categoría
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/logo.png", "data\logo-C4.png")
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/titulo.png", "data\titulo-C4.png")
                           
                                             ' Crear un nuevo ArrayList
                                             Set lines = CreateObject("System.Collections.ArrayList")
                           
                                             ' Abrir el archivo para leer
                                             Set objFSO = CreateObject("Scripting.FileSystemObject")
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                           
                                             ' Leer todas las líneas en el ArrayList
                                             Do Until objFile.AtEndOfStream
                                                 lines.Add objFile.ReadLine
                                             Loop
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             ' Cambia la tercera línea al nombre de la nueva categoría
                                             lines.Item(3) = nuevacategoria
                           
                                             ' Abre el archivo para escribir
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                           
                                             ' Escribe todas las líneas en el archivo
                                             For Each line in lines
                                                 objFile.WriteLine line
                                             Next
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             Set WshShell = CreateObject("WScript.Shell")
                                             Return = WshShell.Run("HeavyNight.exe", 1, False)
                           
                                             MsgBox "La categoria " & categoriavieja & " ha cambiado y se ha creado una copia en su bóveda para presentarte a nuestra nueva categoría " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                         Else
                                             ' No hacer nada en caso de responder a no.
                                         End If
                                     Else
                                         Set oShell = CreateObject("WScript.Shell")
                                         oShell.Run "cmd /c taskkill /IM HeavyNight.exe", 0, True
                           
                                         strFolder = "launcher\" & categoriavieja & ""
                                         strDestFolder = "launcher\zboveda"
                                         strNewFolderName = nuevacategoria
                           
                                         If Not objFSO.FolderExists(objFSO.GetParentFolderName(strDestFolder)) Then
                                             objFSO.CreateFolder objFSO.GetParentFolderName(strDestFolder)
                                         End If
                           
                                         If Not objFSO.FolderExists(strDestFolder) Then
                                             objFSO.CreateFolder strDestFolder
                                         End If
                           
                                         If objFSO.FolderExists(strFolder) Then
                                             sourceFolderName = objFSO.GetFolder(strFolder).Name
                                             destFolder = objFSO.BuildPath(strDestFolder, sourceFolderName)
                                             objFSO.CreateFolder destFolder
                           
                                             arrFolders = Array("config", "mods", "saves", "scripts")
                           
                                             For Each subFolder In arrFolders
                                                 FolderDel = "launcher\" & categoriavieja & "\" & subFolder
                                                 ' Verificar si la carpeta existe antes de intentar eliminarla
                                                 If objFSO.FolderExists(FolderDel) Then
                                                     objFSO.DeleteFolder(FolderDel)
                                                 End If
                                             Next
                           
                                             ' Renombrar la carpeta de origen al nuevo nombre
                                             objFSO.MoveFolder strFolder, objFSO.BuildPath(objFSO.GetParentFolderName(strFolder), strNewFolderName)
                           
                                             'Editar el archivo "launchcfg"
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_10]", "//[p_6_img_button_10]" '(b_parches4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_6]", "//[p_6_img_button_6]" '(b_delete4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "[p_6_img_button_11]", "//[p_6_img_button_11]" '(b_jugar4_a.png)
                                             EditLaunchCfgFile "data\launchcfg", "//[p_6_img_button_12]", "[p_6_img_button_12]" '(b_instalar4_a.png)
                                             
                                             ' Descargar los archivos de instalar nueva categoría
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/logo.png", "data\logo-C4.png")
                                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/titulo.png", "data\titulo-C4.png")
                           
                                             ' Crear un nuevo ArrayList
                                             Set lines = CreateObject("System.Collections.ArrayList")
                           
                                             ' Abrir el archivo para leer
                                             Set objFSO = CreateObject("Scripting.FileSystemObject")
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                           
                                             ' Leer todas las líneas en el ArrayList
                                             Do Until objFile.AtEndOfStream
                                                 lines.Add objFile.ReadLine
                                             Loop
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             ' Cambia la tercera línea al nombre de la nueva categoría
                                             lines.Item(3) = nuevacategoria
                           
                                             ' Abre el archivo para escribir
                                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                           
                                             ' Escribe todas las líneas en el archivo
                                             For Each line in lines
                                                 objFile.WriteLine line
                                             Next
                           
                                             ' Cierra el archivo
                                             objFile.Close
                           
                                             Set WshShell = CreateObject("WScript.Shell")
                                             Return = WshShell.Run("HeavyNight.exe", 1, False)
                           
                                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                                         End If
                                     End If
                                 End If
                         Else
                             ' Descargar los archivos de instalar nueva categoría
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/logo.png", "data\logo-C4.png")
                             Call DownloadFile("https://heavynightlauncher.com/Launcher-Categorias/Categoria4/imagenes/titulo.png", "data\titulo-C4.png")
 
                             ' Crear un nuevo ArrayList
                             Set lines = CreateObject("System.Collections.ArrayList")
                       
                             ' Abrir el archivo para leer
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 1)
                       
                             ' Leer todas las líneas en el ArrayList
                             Do Until objFile.AtEndOfStream
                                 lines.Add objFile.ReadLine
                             Loop
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             ' Cambia la tercera línea al nombre de la nueva categoría
                             lines.Item(3) = nuevacategoria
                       
                             ' Abre el archivo para escribir
                             Set objFSO = CreateObject("Scripting.FileSystemObject")
                             Set objFile = objFSO.OpenTextFile(strLocalFilePath, 2)
                       
                             ' Escribe todas las líneas en el archivo
                             For Each line in lines
                                 objFile.WriteLine line
                             Next
                       
                             ' Cierra el archivo
                             objFile.Close
                       
                             MsgBox "La categoria " & categoriavieja & " ha cambiado y ahora te presentamos " & nuevacategoria & ".", vbInformation, "HeavyNight"
                         End If
                     End If
                 End If
             End If
 End Sub
 
 ' ABRE LA CARPETA MODS DE LA CATEGORIA 4
 Sub SeccionE7()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La tercera línea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             carpeta = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     
     ' ////Comprovacion en la carpeta de la categoria.////
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FolderExists("launcher\" & carpeta & "") Then
         Set objShell = CreateObject("WScript.Shell")
         objShell.Run "explorer.exe """ & "launcher\" & carpeta & """", 1, False
     Else
         MsgBox "Parece que aún no tienes instalada la categoría o no existe la carpeta " & carpeta & "."
     End If
 End Sub
 
 ' ABRE LA WEB INFO DE LA CATEGORIA 4
 Sub SeccionE8()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/news/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub
 
 ' ABRE LA WEB TIENDA DE LA CATEGORIA 4
 Sub SeccionE9()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 0 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             paginaweb = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If
     Set objShell = CreateObject("WScript.Shell")
     link = "https://www.heavynight.com/shop/categories/" & paginaweb & ""  ' Reemplaza con tu enlace deseado
     
     objShell.Run link
 End Sub

 ' COPIAR IP DE LA CATEGORIA 4
 Sub SeccionE10()
     url = "https://www.heavynightlauncher.com/Launcher-Categorias/Categoria4/Category-Name.php"
     lineNumber = 3 ' La primera linea
     
     Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
     xmlhttp.Open "GET", url, False
     xmlhttp.Send
     
     If xmlhttp.Status = 200 Then
         responseLines = Split(xmlhttp.responseText, "<br>") ' Divide la respuesta por la etiqueta HTML <br>
         If UBound(responseLines) >= lineNumber Then
             ipserver = responseLines(lineNumber)
         Else
             MsgBox "La línea solicitada no existe en la respuesta."
         End If
     Else
         MsgBox "No se pudo obtener el valor de la URL. Código de estado: " & xmlhttp.Status
     End If

     ' Crea un objeto Shell para acceder al portapapeles
     Set objShell = CreateObject("WScript.Shell")
     
     ' Copia el texto al portapapeles
     objShell.Run "cmd /c echo " & ipserver & "| clip", 2, True
     
     ' Muestra un mensaje para indicar que se ha copiado el texto
     texto = "La IP ha sido copiado en tu portapapeles"
     MyBox = MsgBox(texto,266304,"HeavyNight!")
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' \\\\\\\\\\\ DEFINICION DE CADA SECCION CREADA \\\\\\\\\\\

 ' Función para ejecutar una sección y manejar errores
 Sub EjecutarSeccionConManejoDeErrores(Seccion)
     On Error Resume Next
     Select Case hn
        ' LAUNCHER
         Case "categoryinstall"
             SeccionA1()
         Case "categorydelete"
             SeccionA2()
         Case "errorversion"
             SeccionA3()
         Case "mantenimiento"
             SeccionA4()
         Case "launcherupdate"
             SeccionA5()
        ' MODPACK 1
         Case "c1install"
             SeccionB1()
         Case "c1delete"
             SeccionB2()
         Case "c1launch"
             SeccionB3()
         Case "c1notifiupdate"
             SeccionB4()
         Case "c1parche"
             SeccionB5()
         Case "c1nuevacategoria"
             SeccionB6()
         Case "c1mods"
             SeccionB7()
         Case "c1web"
             SeccionB8()
         Case "c1tienda"
             SeccionB9()
         Case "c1ipcopy"
             SeccionB10()
        ' MODPACK 2
         Case "c2install"
             SeccionC1()
         Case "c2delete"
             SeccionC2()
         Case "c2launch"
             SeccionC3()
         Case "c2notifiupdate"
             SeccionC4()
         Case "c2parche"
             SeccionC5()
         Case "c2nuevacategoria"
             SeccionC6()
         Case "c2mods"
             SeccionC7()
         Case "c2web"
             SeccionC8()
         Case "c2tienda"
             SeccionC9()
         Case "c2ipcopy"
             SeccionC10()
        ' MODPACK 3
         Case "c3install"
             SeccionD1()
         Case "c3delete"
             SeccionD2()
         Case "c3launch"
             SeccionD3()
         Case "c3notifiupdate"
             SeccionD4()
         Case "c3parche"
             SeccionD5()
         Case "c3nuevacategoria"
             SeccionD6()
         Case "c3mods"
             SeccionD7()
         Case "c3web"
             SeccionD8()
         Case "c3tienda"
             SeccionD9()
         Case "c3ipcopy"
             SeccionD10()
        ' MODPACK 4
         Case "c4install"
             SeccionE1()
         Case "c4delete"
             SeccionE2()
         Case "c4launch"
             SeccionE3()
         Case "c4notifiupdate"
             SeccionE4()
         Case "c4parche"
             SeccionE5()
         Case "c4nuevacategoria"
             SeccionE6()
         Case "c4mods"
             SeccionE7()
         Case "c4web"
             SeccionE8()
         Case "c4tienda"
             SeccionE9()
         Case "c4ipcopy"
             SeccionE10()
         Case Else
             WScript.Echo "Valor no válido para hn. Use una 'case'."
     End Select

     ' Mostrar ventana emergente de error predeterminada si hubo un error en la sección
     If Err.Number <> 0 Then
         MsgBox "Error al ejecutar la sección: " & Err.Description, vbExclamation, "Error"
     End If
     On Error GoTo 0 ' Restaurar el manejo de errores normal
 End Sub
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' Ejecuta la sección de código correspondiente en función del argumento "hn"
EjecutarSeccion hn