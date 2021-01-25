Attribute VB_Name = "Herramientas"
' Procesos internos

' Abrir un Formulario
' CarpetaProyecto       --> Devuelve el PATH completo de la carpeta del programa,
'                           si se indica una subcarpeta, devuelve el PATH incluyéndola.
'                           Si no existe, crea dicha carpeta.
' CreaCarpeta           --> Crea una carpeta con el nombre indicado.
' MostrarEstado          -> Permite mostrar en un label un mensaje definido en un color especifico (estandard para este proyecto)
' Regards                -> Devuelve el saludo dependiendo de la hora del día.
' Extraer               --> Extrae numero o caracteres de un string dependiendo de la seleccion.
' RutaCarpetasEspeciales -> Devuelve la ruta de una carpeta especial del sistema dependiendo
'                           del valor numerico introducido. Mas notas en la Funcion.
' RandomNumbers         --> Genera numeros Random
' ArchivoExiste         --> Devuelve un valor 'true' si el archivo indicado existe
' Columna                -> Devuelve el valor numerico de una columna o viceversa
' UltimoDiaDelMes       --> Devuelve el último día del mes de un mes determinado en un año determinado.
' CrearArchivoPlano      -> Crea archivo plano introduciendo la ruta+archivo+extension y el texto que quieras grabar.

' FormCargado            -> Devuelve un Verdadero o Falso si el formulario del que viene el activo es el indicado en la funcion
' LimpiarNombreArchivo   -> Limpia el nombre de un arhivo de caracteres no permitidos sustituyéndolo por uno elegido por el usuario.
'                           Por defecto: Un espacio (Codigo chr(32)
' Idioma                --> Devuelve el Idioma en el que está la aplicación (Inglés o Españo) o el código de lenguaje si no reconoce los anteriores.
' AbrirArchivo          --> Permite abrir un archivo o vínculo a internet.

Option Explicit
Private Const NombreCarpetaTrabajo As String = "fbExcel"   ' Nombre de la carpeta donde se guardará cualquier archivo generado en esta app.

''' Creamos un acceso directo a la carpeta del programa de manera que
'   siempre podamos acceder a ella a través del nombre de función CarpetaProyecto
'   y sea cual sea el nombre que haya definido el Usuario.
'   Si no ponemos NomCarpeta solo genera la carpeta indicada en ruta
'   Función personalizada - J.G.Arvidsson
Function CarpetaProyecto(Optional subcarpeta As String = "", Optional CrearSiNoExiste As Boolean = True) As String
''' Declaramos la variables
    Dim Dir_Raiz As String
    Dim Dir_Main As String
    Dim RutaCompleta As String
    Dim SoloDir As String           ' Sí solo queremos el Path completo y no ponemos subcarpeta
    
    Dir_Raiz = Application.DefaultFilePath
    Dir_Main = NombreCarpetaTrabajo         ' Definido por el ususario en la parte superior de este Módulo

    SoloDir = subcarpeta & "\"
    If subcarpeta = "" Then SoloDir = ""
    CarpetaProyecto = Dir_Raiz & "\" & Dir_Main & "\" & SoloDir
    If CrearSiNoExiste Then _
        If Dir(CarpetaProyecto, vbDirectory) = "" Then MkDir CarpetaProyecto

End Function


''' Creamos una Carpeta Nueva, si existe no pasa nada.
'   Si no ponemos NomCarpeta solo genera la carpeta indicada en ruta
'   Función personalizada - J.G.Arvidsson
Sub CreaCarpeta(ruta As String, Optional NomCarpeta As String = "")

    Dim RutaCompleta As String
    
    If NomCarpeta = "" Then
        RutaCompleta = ruta
    Else
        RutaCompleta = ruta & "\" & NomCarpeta
    End If
    
    If Dir(RutaCompleta, vbDirectory) = "" Then MkDir RutaCompleta

End Sub

''' Prioridades:
'   1 -> Normal (Color Negro)
'   2 -> Conexion OK (Color Verde?)
'   3 -> Conexion KO (Color Naranja?)
'   4 -> En Local OK (Color Azul)
'   5 -> Mal, muy mal (ROJO negrita)
Function MostrarEstado(DondeLoMuestro As MSForms.Label, QueMensajeMuestro As String, Prioridad As Single)

    If Prioridad = 1 Then
        DondeLoMuestro.ForeColor = &H80000012
        DondeLoMuestro.Font.Bold = False
    ElseIf Prioridad = 2 Then
        DondeLoMuestro.ForeColor = &H8000&
        DondeLoMuestro.Font.Bold = False
    ElseIf Prioridad = 3 Then
        DondeLoMuestro.ForeColor = &H40C0&
        DondeLoMuestro.Font.Bold = False
    ElseIf Prioridad = 4 Then
        DondeLoMuestro.ForeColor = &H800000
        DondeLoMuestro.Font.Bold = False
    ElseIf Prioridad = 5 Then
        DondeLoMuestro.ForeColor = &HFF&
        DondeLoMuestro.Font.Bold = True
    End If
    
    
    DondeLoMuestro.Caption = QueMensajeMuestro


End Function

Function Regards(Idioma_Es_o_In As String) As String
''' Declaramos la variable
    Dim HoraActual As Date
    Dim dia As String
    Dim tarde As String
    Dim noche As String

''' Asignamos el valor de la hora a la variable
    HoraActual = Time
    
''' Según el Idioma elegido, cargamos las variables con los datos
    If Idioma_Es_o_In = "Es" Then
        dia = "Buenos días"
        tarde = "Buenas tardes"
        noche = "Buenas noches"
    ElseIf Idioma_Es_o_In = "In" Then
        dia = "Good Morning"
        tarde = "Good Afternoon"
        noche = "Good Evening"
    End If
    
''' Miramos en qué tramo del día estamos y devolvemos una respuesta
    Select Case HoraActual
        Case "00:00" To "11:59:59"
            Regards = dia
        Case "12:00" To "18:59:59"
            Regards = tarde
        Case "19:00" To "23:59:59"
            Regards = noche
    End Select
End Function



''' Extraemos los caracteres o números de una cadena dependiendo de nuestra eleccion buleana.
'   Función personalizada - J.G.Arvidsson
Function Extraer(cadena, Caracteres As Boolean) As String

    Dim i As Double
    Dim dimension As Double
    Dim Caracter As String
    Dim Numerico As String
    
    dimension = Len(cadena) ' Primero vemos cuanto mide la cadena
    For i = 1 To dimension
        If IsNumeric(Mid(cadena, i, 1)) Then
            Numerico = Numerico + Mid(cadena, i, 1)
        Else
            Caracter = Caracter + Mid(cadena, i, 1)
        End If
    
    Next i

''' Para finalizar, dependiendo de nuestra elección, mostraremos una u otra
    If Caracteres = True Then Extraer = Caracter
    If Caracteres = False Then Extraer = Numerico

End Function

''' Devuelve el 'Path' de una carpeta específica del sistema.
'   Función personalizada - J.G.Arvidsson
Function RutaCarpetasEspeciales(Numero As Single)
''' Valores devueltos:
'   0.- Donde se encuentre este Documento
'   1.- Mis documentos
'   2.- Escritorio
'   3.- Escritorio para todos los usuarios
'   4.- Archivos recientes
'   5.- Mis favoritos
'   6.- Archivos de programa
'   7.- Menú de Inicio
'   8.- Enviar "archivo" a...
'   9.- Directorio AppData
'  10.- C:/

    Dim objFolders As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders
    
    If Numero = 0 Then                                          ' 0 = Ruta del Libro de Excel
        RutaCarpetasEspeciales = ThisWorkbook.path
    ElseIf Numero = 1 Then                                      ' 1 = Mis Documentos
        RutaCarpetasEspeciales = objFolders("mydocuments")
    ElseIf Numero = 2 Then                                      ' 2 = Escritorio
        RutaCarpetasEspeciales = objFolders("desktop")
    ElseIf Numero = 3 Then                                      ' 3 = Escritorio para todos los usuarios
        RutaCarpetasEspeciales = objFolders("alluserdesktop")
    ElseIf Numero = 4 Then                                      ' 4 = Archivos Recientes
        RutaCarpetasEspeciales = objFolders("recent")
    ElseIf Numero = 5 Then                                      ' 5 = Mis Favoritos
        RutaCarpetasEspeciales = objFolders("favorites")
    ElseIf Numero = 6 Then                                      ' 6 = Programs
        RutaCarpetasEspeciales = objFolders("programs")
    ElseIf Numero = 7 Then                                      ' 7 = Menu de Inicio
        RutaCarpetasEspeciales = objFolders("startmenu")
    ElseIf Numero = 8 Then                                      ' 8 = Enviar a...
        RutaCarpetasEspeciales = objFolders("SendTo")
    ElseIf Numero = 9 Then                                      ' 9 = AppData (donde está BinPar)
        RutaCarpetasEspeciales = objFolders("AppData")
    ElseIf Numero = 10 Then                                     ' 10 = C:\
        RutaCarpetasEspeciales = "C:\"
    End If
End Function

''' Genera un número Random
'   Encontrado en internet, sin atribución.
Function RandomNumbers(Lowest As Long, Highest As Long, Optional Decimals As Integer)
    Application.Volatile  'Remove this line to “freeze” the numbers
    
    If IsMissing(Decimals) Or Decimals = 0 Then
        Randomize
        RandomNumbers = Int((Highest + 1 - Lowest) * Rnd + Lowest)
    Else
        Randomize
        RandomNumbers = Round((Highest - Lowest) * Rnd + Lowest, Decimals)
    End If
End Function

''' Devuelve un verdadero si el archivo indicado existe
'   Función personalizada - J.G.Arvidsson
Function ArchivoExiste(ArhivoconRuta As String) As Boolean
    Dim Archivo As String
    ArchivoExiste = True
        If Dir(ArhivoconRuta) = "" Then
            ArchivoExiste = False
            Exit Function
        End If
End Function

''' Devuelve la letra correspondiente a una columna indicando su posición numérica:
'   a = 1, b = 2, c = 3,...
'   Y viceversa.
'   Encontrado en internet, sin atribución.
Function columna(col)
    On Error Resume Next
    If IsNumeric(col) = True Then
        columna = Mid(Split(Columns(col).address, ":")(1), 2)
    Else
        columna = Columns(col).Column
    End If
End Function


''' Devuelve el último día del mes de un mes determinado en un año determinado.
'   Encontrado en internet, sin atribución.
Function UltimoDiaDelMes(Mes As Single, year As Single) As Single
    Dim bisiesto As Single
    
    If year >= 0 Then
        If year Mod 4 = 0 And (year Mod 100 <> 0 Or year Mod 400 = 0) Then ' Verificando que el año sea bisiesto
            bisiesto = 1
        Else
            bisiesto = 0
        End If
        
        If Mes <= 12 Or Mes >= 1 Then       ' Verificando que el mes se encuentre entre 1 y 12
            If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
                UltimoDiaDelMes = 31
            End If
            If Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
                UltimoDiaDelMes = 30
            End If
            If bisiesto = 1 And Mes = 2 Then
                UltimoDiaDelMes = 29
            End If
            If bisiesto = 0 And Mes = 2 Then
                UltimoDiaDelMes = 28
            End If
        End If
        If Mes > 12 Or Mes < 1 Then
            MsgBox "Corrige el número del mes"
        End If
    Else
        MsgBox "Ingrese un año válido"
    End If
End Function

''' Genera un archivo plano (del tipo txt) con la extensión elegida por el usuario.
'   Función personalizada - J.G.Arvidsson
Function CrearArchivoPlano(rutayarchivocompletoconextension As String, cadenadetextoparainsertar As String, Optional BorrarSiYaExiste As Boolean = True)
    On Error Resume Next
    If BorrarSiYaExiste = True Then Kill rutayarchivocompletoconextension
    Open rutayarchivocompletoconextension For Append As #1 ' change the path to reflect your path
    Print #1, cadenadetextoparainsertar
    Close #1
End Function

''' Limpia el nombre de un arhivo de caracteres no permitidos sustituyéndolo por uno elegido por el usuario.
'   Por defecto: Un espacio (Codigo chr(32)
'   Encontrado en internet, sin atribución + modificación para adaptar a necesidades -> J.G.Arvidsson.
Function LimpiarNombreArchivo(NombreArchivo As String, Optional ValorASCII As Double = 32) As String
    Dim Permitido As String
    Dim Chequeo As Object

''' Asignamos valor a la variable con los caracteres permitidos
    Permitido = "[^a-z0-9-_]"

''' Instanciamos el objeto
    Set Chequeo = CreateObject("VBScript.RegExp")

''' Comprobamos la cadena de caracteres del nombre de archivo y remplazamos los valores no validos por el caracter deseado
    With Chequeo
        .Global = True
        .IgnoreCase = True
        .Pattern = Permitido
        
        If .Test(NombreArchivo) Then LimpiarNombreArchivo = .Replace(NombreArchivo, Chr(ValorASCII))
    End With

End Function


''' Devuelve el idioma en el que se encuentra la aplicación.
'   Encontrado en internet, sin atribución.
Function Idioma() As String
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
        Case 1033, 3081, 10249, 4105, 9225, 14345, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297  ' Revisa si está en inglés
            Idioma = "English"
    
        Case 1034, 2058, 3082, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 14346, 8202 ' Revisa si está en español
            Idioma = "Spanish"
            
        Case Else
            MsgBox ("Lenguaje no reconocido como inglés ni español. Puede que la macro no despliegue el año correctamente.")
            Idioma = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    End Select
End Function


''' Abre el archivo indicado o vínculo a Internet
'   Encontrado en internet, sin atribución + modificación para adaptar a necesidades -> J.G.Arvidsson.
Function AbrirArchivo(RutaArchivo As String, Optional NombreArchivo As String = "")
''' Esta funcion nos permite abrir vinculos a archivos usando la funcion de hipervínculo,
'   si se corta el proceso, pueden aparecer errores, por lo que se ha añadido un descriptor
'   de errores a la función.
    Dim Tipo As String
    
    If InStr(1, RutaArchivo & NombreArchivo, "www") > 0 Or _
       InStr(1, RutaArchivo & NombreArchivo, "http") > 0 Then
        Tipo = "vínculo a la web indicada"
    Else
        Tipo = "archivo indicado"
    End If
    
    On Error GoTo handler
    ThisWorkbook.FollowHyperlink RutaArchivo & NombreArchivo
    Exit Function
    
handler:
        'Usamos un Select Case para identificar los números de error
    Select Case Err.Number

    Case Is = -2147467260
        MsgBox "El proceso de acceso fue cancelado.", vbInformation, "Proceso interrumpido por el usuario"
    Case Is = -2146697210
        MsgBox "No se encontró el " & Tipo & ".", vbInformation, "Destino no localizado"
    Case Else
        MsgBox Err.Number & " " & Err.Description
    End Select
End Function

''' Lee archivos localizados en el computador y devuelve el contenido como cadena de texto.
'   Encontrado en internet, sin atribución.
Function ReadTextFile(strPath As String, lngFormat As Long) As String
       ' lngFormat -2 - System default, -1 - Unicode, 0 - ASCII
       On Error Resume Next
       With CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath, 1, False, lngFormat)
           ReadTextFile = ""
           If Not .AtEndOfStream Then ReadTextFile = .ReadAll
           .Close
       End With
End Function

