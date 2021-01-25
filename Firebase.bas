Attribute VB_Name = "Firebase"
' ========================== Proyecto iMelin <---> Firebase Connection ==========================
'
'            Desarrollado por J.G.Arvidsson para White Noise Solution como herramienta
'    para la conexión carga y descarga de datos de Firebase y manejo del Registro de Usuarios
'                                        version 1.0 @2020
'
'         Más informacion en https://github.com/jgarvidsson/Conectar-Excel-con-Firebase
'                                  www.whitenoisesolutions.com
'                                      jgarvidsson@gmail.com
'                                           @jgarvidsson
'
' ===============================================================================================
'
'   Índice de Funciones:
'   FirebaseDB:             Resuelve la operación requerida con Firebase dependiendo del modo indicado:
'                           - POST     - Postear (añade un ID al JSON posteado y retorna dicho ID)
'                           - PATCH    - Agregar mensaje JSON una dirección sin borrar el resto
'                           - PUT      - Añadir valor JSON en una dirección
'                           - GET      - Recibir - En este caso Mensaje irá vacío = ""
'                           - DELETE   - Borra la direccion indicada - En este caso Mensaje irá vacío = ""
'                           - DOWNLOAD - Descarga en un archivo el JSON - En este caso en Mensaje irá la carpeta + nombre de archivo + extensión de destino
'                           - BACKUP   - Copia (y duplica) el contenido de la dirección indicada en la dirección indicada en Menaje en modo "POST".
'                           - MOVE     - Mueve el contenido de la dirección indicada a la dirección indicada en Mensaje.
'                           - COPY     - Copia (y duplica) el contenido de la dirección indicada en la dirección indicada en Mensaje.
'   FirebasePC:             Devuelve Matriz de datos desde un archivo JSON descargado
'   DevolverValorEspecificoDeFirebase: Devuelve un valor específico de una base de datos de Firebase
'   DevolverValorAutorizacion: Devuelve el contenido de un parámetro de la Autorización de Firebase
'   GenerarJSONError:       Genera y envía un JSON con el contenido de un error a una dirección de Firebase
'   RegistrarUso:           Registra el uso del programa en Firebase.
'   arrayLength:            Da el numero de elementos de un Array. Usado para comprobar si los Array redibidos continen información.
'   ComprobarConexion:      Devuelve True o False si hay o no conexión a la red
'   AlmacenarJSON:          Guarda un JSON como string en el una dirección del PC en concreto
'   ParametroDB:            Devuelve un parámetro de acuerdo a las constantes privadas indicadas al incicio
'   AccionConUsuario:       Permite realizar direfentes acciones sobre el usuario de la base de datos:
'                           - NEW:    Crea nuevo usuario a partir del correo electronico y un password.
'                           - INFO:   Recupera información referente al Usuario.
'                           - ANONIMUS: Crea un usuario anonimo. Necesita configurar Autentificación en Firebase.
'                           - AUTH: Recupera los datos de autorización para manejar la BD.
'                           - UPDATE: Actualiza la información de usuario.
'                           - REMOVE: Borra el registro completo de Usuario en el servidor.
'                           - RESETPASSWORD: Manda una petición de recuperación al correo electronico indicado.
'                           - CHANGEMAIL: Cambiar el eMail del Usuario activo.
'                           - CHANGEPASSWORD: Cambiar el Password del Usuario activo.

Option Explicit
Private Const dbNAME As String = "<sustituya por el nombre de la base de datos>"                ' Nombre de la base de datos (sin cabecera http ni servidor).
Private Const dbURL As String = "https://" & dbNAME & ".firebaseio.com/"                        ' Direccion de la base de datos.    ' Servidor en USA
'Private Const dbURL As String = "https://" & dbNAME & ".europe-west1.firebasedatabase.app/"    ' Direccion de la base de datos.    ' Servidor en Europa
Private Const dbAPI As String = "<sustituya por la API de la base de datos>"                    ' API de la base de datos

'Private PathError As String                         ' <-- aquí irán los registros de error en el PC (esto es informativo, no util. Ver mas abajo GenerarJSONError, para más información).
Private Const cURLError As String = "ErrorLog/"      ' <-- aquí irán los registros de error en Firebase (no usado en este ejemplo).
Private TokenAutorizacion As String                  ' Variable que contendrá la autorización de la base de datos


' ===================================================================================================
'   TRABAJAR CON LA BASE DE DATOS FIREBASE REALTIME
' ===================================================================================================

''' Acciones con FireBase

'   iMelinDB es la más básica lectura de un JSON que genera un Array en el que la
'   direccion 0,0 es la cabecera, y el resto los valores.
'   Modes:
'   POST     - Postear (añade un ID al JSON posteado y retorna dicho ID)
'   PATCH    - Agregar mensaje JSON una dirección sin borrar el resto
'   PUT      - Añadir valor JSON en una dirección
'   GET      - Recibir - En este caso Mensaje irá vacío = ""
'   DELETE   - Borra la direccion indicada - En este caso Mensaje irá vacío = ""
'   DOWNLOAD - Descarga en un archivo el JSON - En este caso en Mensaje irá la carpeta + nombre de archivo + extensión de destino
'   BACKUP   - Copia (y duplica) el contenido de la dirección indicada en la dirección indicada en Menaje en modo "POST".
'   MOVE     - Mueve el contenido de la dirección indicada a la dirección indicada en Mensaje.
'   COPY     - Copia (y duplica) el contenido de la dirección indicada en la dirección indicada en Mensaje.

Public Function FirebaseDB(Mode As String, Direccion As String, Mensaje As String, Optional claveautorizacion As String = "", Optional SoloContenidoIndice As Boolean = False) As Variant
' submits a JSON message object to Firebase list
' return value is handle on created object
''' Declaramos las variables
    Dim oHTTP As Object ' HTTP object for connection to database
    Dim sMessage As String ' JSON message sent to database
    Dim sResponseText As String ' response from database: ID of created object
    Dim Dominio As String
    Dim Datos()                 ' Matriz para recolección de datos desde iMerlin (directamente)
    Dim Cabecera()              ' Matriz para recolección de cabeceras desde iMerlin (no es util, pero necesario)
    Dim sState As String
    Dim vJSON
    Dim vFlat
    Dim Salida() As String
    Dim x As Double, y As Double
    Dim n As Double, m As Double
    Dim i As Long
    Dim Modo As String
    Dim URLLicencia As String
    Dim CopiarEn As String
    
    'On Error GoTo SinConexion

''' Dependiendo del Modo, hacemos selecciones de entorno
    If Mode = "GET" Or Mode = "DELETE" Then Mensaje = ""
    Dominio = ParametroDB(2) & Direccion & ".json"
    If claveautorizacion <> "" Then Dominio = ParametroDB(2) & Direccion & ".json?auth=" & claveautorizacion
    If claveautorizacion = "Empty" Then GoTo SinConexion
    If Mode = "DOWNLOAD" Or Mode = "BACKUP" Or Mode = "MOVE" Or Mode = "COPY" Then
        Modo = "GET"
    Else
        Modo = Mode
    End If

''' Creamos la conexion
    Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        oHTTP.Open Modo, Dominio, False
        oHTTP.setRequestHeader "Content-type", "application/json"
        oHTTP.setRequestHeader "Accept", "application/json"
        If Mode = "DOWNLOAD" Or Mode = "DELETE" Or Mode = "BACKUP" Or Mode = "MOVE" Or Mode = "COPY" Then
            oHTTP.send
        Else
            oHTTP.send Mensaje
        End If
        
''' Pasamos la cadena redibida a una variable
    sResponseText = oHTTP.ResponseText

    If sResponseText = "null" And Mode = "DELETE" Then GoTo SalidaAqui

''' Si el estatus devuelto es No Localizada la URL saltamos a SinConexion
    If oHTTP.Status <> 200 Then GoTo SinConexion
    
''' Chequea del archivo JSON
    JSON.Parse sResponseText, vJSON, sState   ' Parse JSON response
    If sState = "Error" Then GoTo NoHayNadaEniMerlin
    
    If Mode = "BACKUP" Or Mode = "MOVE" Or Mode = "COPY" Then GoTo RealizarBackup
    
    If vJSON.Count = 1 Then
        ReDim Salida(0)
        If vJSON.Keys()(0) = "error" Then
            Salida(0) = "error"
            FirebaseDB = Salida
        Else
            Salida(0) = sResponseText
            FirebaseDB = Salida
        End If
    ElseIf vJSON.Count > 1 Then
        ReDim Salida(0)
        Salida(0) = sResponseText
        FirebaseDB = Salida
    End If
    If Mode = "PUT" Or Mode = "POST" Or Mode = "PATCH" Then GoTo SalidaAqui
    If Mode = "DOWNLOAD" Then GoTo SoloDescargar
    JSON.Flatten vJSON, vFlat               ' Flatten JSON
    JSON.ToArray vJSON, Datos, Cabecera     ' Convertimos en Matriz bidimensional

''' Hacemos Recuento de la profuncidad y largo del Array de Salida
    x = vJSON.Count
    y = UBound(Cabecera())

    If SoloContenidoIndice = True Then
    ''' Componemos el Array de Salida
        x = x - 1
        ReDim Salida(x)
        For i = 0 To x
            Salida(i) = vJSON.Keys()(i)
        Next i
    Else
    
    ''' Componemos el Array de Salida
        ReDim Salida(x, y)
        
    ''' Rellenamos el Array de Salida con los datos extraidos del JSON
        For n = 0 To x
        If n = 20 Then
            n = n
        End If
            For m = 0 To y
                If n = 0 Then
                    Salida(n, m) = Cabecera(m)
                Else
                    Salida(n, m) = Datos(n - 1, m)  ' -1 porque hemos usado el 0,0 de Salida para la cabecera
                End If
            Next m
        Next n
    End If
    FirebaseDB = Salida
    Set oHTTP = Nothing
    Exit Function
    
SalidaAqui:
    ReDim Salida(0)
    Salida(0) = sResponseText
    FirebaseDB = Salida
    Set oHTTP = Nothing
    Exit Function

SoloDescargar:                                              ' Solo descargarmos y guardamos donde indiquemos el archivo JSON
    If Dir(Mensaje) <> "" Then Kill Mensaje
    Open Mensaje For Append As #1
    Print #1, sResponseText
    Close #1
    Salida(0) = sResponseText
    FirebaseDB = Salida
    Set oHTTP = Nothing
Exit Function

NoHayNadaEniMerlin:                                         ' Si no hay información recuperada en modo JSON, lanzamos un EMPTY o el contenido
    ReDim Salida(0)
    If Len(sResponseText) > 0 Then
        Salida(0) = sResponseText
    Else
        Salida(0) = "Empty"
    End If
        FirebaseDB = Salida
        Set oHTTP = Nothing
    Exit Function

RealizarBackup:                                             ' Realizamos la copia a otra parte de la base de datos
    URLLicencia = ParametroDB(2)
    CopiarEn = URLLicencia & Mensaje & ".json?auth=" & claveautorizacion
    oHTTP.Open IIf(Mode = "BACKUP", "POST", "PUT"), CopiarEn, False
    'oHTTP.Open "PUT", CopiarEn, False
    oHTTP.setRequestHeader "Content-type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.send sResponseText
    If Mode = "MOVE" Then
        oHTTP.Open "DELETE", Dominio, False
        oHTTP.setRequestHeader "Content-type", "application/json"
        oHTTP.setRequestHeader "Accept", "application/json"
        oHTTP.send
    End If
    ReDim Salida(0)
    Salida(0) = sResponseText
    FirebaseDB = Salida
    Set oHTTP = Nothing
    Exit Function

SinConexion:                                                ' No hay conexion, devolvemos variable con un aviso
    ReDim Salida(0)
    Salida(0) = "Disconnected"
    FirebaseDB = Salida
    Set oHTTP = Nothing
err1:
fin:
End Function

''' Usando la Función FirebaseDB buscamos un valor específico devolviendo el contenido.

Function DevolverValorEspecificoDeFirebase(Direccion As String, Valor As String, Optional claveautorizacion As String = "") As String
''' Declaramos las variables
    Dim matriz As Variant
    Dim nSubMatriz As Double
    Dim ValorUpdate As String
    Dim nMatriz As Long
    Dim n As Single, m  As Single
    
    On Error Resume Next
    matriz = FirebaseDB("GET", Direccion, "", claveautorizacion)
    nMatriz = arrayLength(matriz)
    If nMatriz = 0 Then
        GenerarJSONError "0", "Information not located", "Search in " & Direccion & " the value " & Valor & " and it was not located.", claveautorizacion
        DevolverValorEspecificoDeFirebase = ""
        Exit Function
    End If
    nSubMatriz = UBound(matriz, 2)
    For n = 0 To nMatriz
        If matriz(n, 0) = Valor Then
                If matriz(n, 1) = "" Then
                    For m = 1 To nSubMatriz
                        ValorUpdate = matriz(n, m)
                            If ValorUpdate <> "" Then
                                DevolverValorEspecificoDeFirebase = ValorUpdate
                                Exit Function
                            End If
                    Next m
                End If
            
            DevolverValorEspecificoDeFirebase = matriz(n, 1)
            If nMatriz > 0 Then Erase matriz
            Exit Function
        End If
    Next n
End Function

''' Trabajamos como si estuvieramos en la nube pero en una direccion específica enviada por Direccion
'   GET     - Recibir - En este caso Mensaje irá vacío = ""

Function FirebasePC(Mode As String, Direccion As String, Mensaje As String, Optional SoloContenidoIndice As Boolean = False) As Variant

''' Declaramos las variables
    Dim sResponseText As String ' Cadena JSON leida.
    Dim Dominio As String
    Dim Datos()                 ' Matriz para recolección con los datos del JSON
    Dim Cabecera()              ' Matriz para recolección de cabeceras del JSON.
    Dim sState As String
    Dim vJSON
    Dim vFlat
    Dim Salida() As String
    Dim x As Double, y As Double
    Dim n As Double, m As Double
    Dim i As Long
    
    If Mode = "GET" Or Mode = "DELETE" Then Mensaje = ""
    Dominio = IIf(InStr(1, Direccion, ".json") > 1, Direccion, Direccion & ".json")

    sResponseText = ReadTextFile(Dominio, -2)

''' Chequea del archivo JSON
    JSON.Parse sResponseText, vJSON, sState
    If sState = "Error" Then GoTo NoHayNada
    JSON.Flatten vJSON, vFlat
    JSON.ToArray vJSON, Datos, Cabecera             ' Convertimos en Matriz bidimensional

''' Hacemos Recuento de la profuncidad y largo del Array de Salida
    x = vJSON.Count
    y = UBound(Cabecera())

    If SoloContenidoIndice = True Then             ' Componemos el Array de Salida
        x = x - 1
        ReDim Salida(x)
        For i = 0 To x
            Salida(i) = vJSON.Keys()(i)
        Next i
    Else
        ReDim Salida(x, y)                          ' Componemos el Array de Salida
    
        For n = 0 To x                              ' Rellenamos el Array de Salida con los datos extraidos del JSON
        If n = 20 Then
            n = n
        End If
            For m = 0 To y
                If n = 0 Then
                    Salida(n, m) = Cabecera(m)
                Else
                    Salida(n, m) = Datos(n - 1, m)  ' -1 porque hemos usado el 0,0 de Salida para la cabecera
                End If
            Next m
        Next n
    End If

FirebasePC = Salida

NoHayNada:
err1:
End Function

''' Usando la Función FirebasePC buscamos un valor específico de un archivo JSON localizado en el
'   PC devolviendo el contenido.

Function DevolverValorEspecificoDeJSONLocal(DirectorioYArchivo As String, Valor As String) As String
    Dim matriz As Variant
    Dim nSubMatriz As Double
    Dim ValorUpdate As String
    Dim nMatriz As Long
    Dim n As Single, m  As Single
    On Error Resume Next
    matriz = FirebasePC("GET", DirectorioYArchivo, "", False)
    nMatriz = arrayLength(matriz)
    If nMatriz = 0 Then
        GenerarJSONError "0", "Information not located", "The value " & Valor & " and it was not located.", ""
        DevolverValorEspecificoDeJSONLocal = ""
        Exit Function
    End If
     nSubMatriz = UBound(matriz, 2)
    If nMatriz - 1 < 2 Then
        For n = 0 To nSubMatriz
            If matriz(0, n) = Valor Then
                If matriz(1, n) <> "" Then
                    DevolverValorEspecificoDeJSONLocal = matriz(1, n)
                    Exit Function
                End If
            End If
        Next n
    Else
        For n = 0 To nMatriz
            For m = 0 To nSubMatriz
                If matriz(n, m) = Valor Then
                    ValorUpdate = matriz(n, m + 1)
                        DevolverValorEspecificoDeJSONLocal = ValorUpdate
                        Exit Function
                End If
            Next m
        Next n
    End If
End Function

''' Devuelve un Valor específico de la cadena con la autorización enviada por Firebase
'   Si Valor = "IdToken" devuelve el Token de Autorización

Function DevolverValorAutorizacion(Valor As String, IDUsereMail As String, IDUserPassword As String) As String
''' Declaramos las variables
    Dim matriz As Variant
    Dim nMatriz As Double
    Dim n As Single

    matriz = AccionConUsuario("AUTH", IDUsereMail, IDUserPassword)
    nMatriz = arrayLength(matriz)
    For n = 0 To nMatriz - 1
        If matriz(n, 0) = Valor Then
            DevolverValorAutorizacion = matriz(n, 1)
            GoTo fin
        End If
    Next n
    
fin:
    If nMatriz > 0 Then Erase matriz
    Exit Function
End Function

''' Esta función recoge los errores y genera un JSON para:
'   - Ser enviado al servidor (ayuda al desarrollador), y si no puede porque no hay conexión a internet
'   - Generar un archivo JSON en Registers para ser enviado en otro momento.

Function GenerarJSONError(NumeroError, descripcionerror, Mensaje, IdToken As String) As Variant
''' Declaramos las variables
    Dim CadenaText As String
    Dim rutaYnombreDeSalida As String               ' En caso de fallo de conexion o de procesamiento en el servidor
    Dim Respuesta As Variant

''' Asignamos valores
    rutaYnombreDeSalida = RutaCarpetasEspeciales(1) & "ERROR_" & NumeroError & "_" & Format(Now(), "yyyymmdd_hhnnss") & ".json"
    
On Error GoTo NaHayConexion
''' Registramos el último uso
    CadenaText = "{" & _
    """Device"":""" & Environ("Userdomain") & """," & _
    """Workbook"":""" & ThisWorkbook.Name & """," & _
    """Error_Number"":""" & NumeroError & """," & _
    """Error_Description"":""" & descripcionerror & """," & _
    """Error_Mensaje"":""" & Mensaje & """," & _
    """TSL"":""" & Now() & """," & _
    """Timestamp"":{"".sv"":""timestamp""}" & _
    "}"

    Respuesta = FirebaseDB("POST", cURLError, CadenaText, IdToken)
    If UBound(Respuesta) > 0 And UBound(Respuesta, 2) = 1 Then _
        If Respuesta(1, 0) = "error" Then GoTo NaHayConexion

    If arrayLength(Respuesta) = 0 Then Erase Respuesta
    On Error GoTo 0
    Exit Function
    
NaHayConexion:          ' SI no hay conexión, se almacena en un archivo
    AlmacenarJSON rutaYnombreDeSalida, CadenaText
    If arrayLength(Respuesta) = 0 Then Erase Respuesta
    On Error GoTo 0

End Function

Function RegistrarUso(Optional qModulo As String = "", Optional ElToken As String = "")

    Dim CadenaText As String
    Dim Token As String

    If qModulo = "" Then qModulo = "App"
    
''' Registramos el último uso
    CadenaText = "{" & _
    """Domain"":""" & Environ("Userdomain") & """," & _
    """Workbook"":""" & ThisWorkbook.Name & """," & _
    """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
    """TSS"":{"".sv"":""timestamp""}" & _
    "}"
    
    FirebaseDB "POST", "AHG/Logs/", CadenaText, ElToken

End Function


''' Para comprobar si hay conexión con la base de datos de Firebase creo una
'   carpeta llamada conexion y la dejo en sólo lectura.
'   En esta carpeta dejo un Valor llamado Permitida con el contenido 'true'

''' Comprobamos la conexion y devolvemos un valor buleano.
'   De esta manera vamos chequeando cada intento de conexión y dejando un mensaje.
Function ComprobarConexion() As Boolean
    Dim Valor As Variant
    On Error GoTo ErrorConexion

    Valor = FirebaseDB("GET", "Conexion/Permitida/", "")
    ComprobarConexion = Valor(0)
    If arrayLength(Valor) > 0 Then Erase Valor
    On Error GoTo 0
    Exit Function
ErrorConexion:
    ComprobarConexion = False
End Function


''' Guarda el contenido de una cadena JSON en un archivo de texto.

Function AlmacenarJSON(RutaYArchivo As String, Contenido As String)
    If Err.Number = 55 Then Close #1
    Open RutaYArchivo For Append As #1
    Print #1, Contenido
    Close #1
End Function



''' Módulo desarrollado a partir de la información contenida en:
'   https://developers.google.com/identity/toolkit/web/reference/relyingparty/
'   https://firebase.google.com/docs/reference/rest/auth#section-send-password-reset-email
'   Antes de empezar hay que hacer una petición AUTH para requerir el IDToken,
'   una vez pedida y almacenada en una Variable Privada, ya podrá operarse con
'   esta función.
'   Función personalizada - J.G.Arvidsson
Function AccionConUsuario(Accion As String, IDUsereMail As String, IDUserPassword As String, _
                            Optional IDTokenUser As String = "", Optional IDNameUser As String = "", _
                            Optional IDURLFoto As String = "") As Variant
''' Declaramos las variables
    Dim URLdb As String
    Dim CargaUtil As String
    Dim sResponseText As String
    Dim TextoLimpio As String
    Dim ArrayAuth() As String
    Dim Cadenas As Variant
    Dim nCadenas As Double
    Dim n As Single                     ' Conteo de For
    Dim a As Single                     ' Almacena el numero de valores de providerUserInfo si los hay
    Dim LoginRequest As Object
    Dim ValorAccion As String
    Dim LocalIdUsuario As String
    Dim Parsed As Dictionary            ' Almacenará los datos JSON antes de convertirlo en Array
    Dim Salida As String                ' Contendrá la carga de información filtrada
    Dim CabeceraHTTP As String          ' Contiene el formato de envío HTTP
''' Recuperamos valores realizando una petición de usuario a Firebase
    'TokenAutorizacion = DevolverValorFirebase("idToken", IDUsereMail, IDUserPassword)
    'LocalIdUsuario = DevolverValorFirebase("localID", IDUsereMail, IDUserPassword)
    CabeceraHTTP = "application/json"
    
    If Accion = "NEW" Then
        ValorAccion = "signupNewUser"
        CargaUtil = "{" & _
                IIf(IDNameUser <> "", """" & "displayName" & """:""" & IDNameUser & """,", "") & _
                IIf(IDURLFoto <> "", """" & "photoUrl" & """:""" & IDURLFoto & """,", "") & _
                """" & "email" & """:""" & IDUsereMail & """," & _
                """" & "emailVerified" & """:""" & "true" & """," & _
                """" & "password" & """:""" & IDUserPassword & """" & _
                "}"
    ElseIf Accion = "ANONIMUS" Then
        ValorAccion = "signupNewUser"
        CargaUtil = "{" & _
                """" & "returnSecureToken" & """:""" & "true" & """" & _
                "}"
    ElseIf Accion = "INFO" Then
        ValorAccion = "getAccountInfo"
        CargaUtil = "{" & _
                """" & "idToken" & """:""" & IDTokenUser & """," & _
                """" & "email" & """:""" & IDUsereMail & """," & _
                """" & "localID" & """:""" & LocalIdUsuario & """" & _
                "}"
    ElseIf Accion = "UPDATE" Then
        ValorAccion = "setAccountInfo"
        CargaUtil = "{" & _
                """" & "idToken" & """:""" & IDTokenUser & """," & _
                """" & "email" & """:""" & IDUsereMail & """," & _
                IIf(IDNameUser <> "", """" & "displayName" & """:""" & IDNameUser & """,", "") & _
                IIf(IDURLFoto <> "", """" & "photoUrl" & """:""" & IDURLFoto & """,", "") & _
                """" & "password" & """:""" & IDUserPassword & """" & _
                "}"
    ElseIf Accion = "AUTH" Then
        ValorAccion = "verifyPassword"
        CargaUtil = "{" & _
                """" & "email" & """:""" & IDUsereMail & """," & _
                """" & "password" & """:""" & IDUserPassword & """," & _
                """" & "returnSecureToken" & """:""" & "true" & """" & _
                "}"
    ElseIf Accion = "REMOVE" Then
        ValorAccion = "deleteAccount"
        CargaUtil = "{" & _
                """" & "idToken" & """:""" & IDTokenUser & """" & _
                "}"
    ElseIf Accion = "RESETPASSWORD" Then
        ValorAccion = "https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode"
        CargaUtil = "{" & _
                """" & "requestType" & """:""" & "PASSWORD_RESET" & """," & _
                """" & "email" & """:""" & IDUsereMail & """" & _
                "}"
    ElseIf Accion = "CHANGEMAIL" Then
        ValorAccion = "https://identitytoolkit.googleapis.com/v1/accounts:update"
        CargaUtil = "{" & _
                """" & "email" & """:""" & IDUsereMail & """," & _
                """" & "idToken" & """:""" & IDTokenUser & """," & _
                """" & "returnSecureToken" & """:""" & "true" & """" & _
                "}"
    ElseIf Accion = "CHANGEPASSWORD" Then
        ValorAccion = "https://identitytoolkit.googleapis.com/v1/accounts:update"
        CargaUtil = "{" & _
                """" & "password" & """:""" & IDUserPassword & """," & _
                """" & "idToken" & """:""" & IDTokenUser & """," & _
                """" & "returnSecureToken" & """:""" & "true" & """" & _
                "}"
    End If
   Debug.Print CargaUtil

    URLdb = "https://www.googleapis.com/identitytoolkit/v3/relyingparty/" & ValorAccion & "?key=" & dbAPI
    If InStr(1, ValorAccion, "https") > 0 Then URLdb = ValorAccion & "?key=" & dbAPI
    
   ' On Error GoTo SinConexion
    Set LoginRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    LoginRequest.Open "POST", URLdb, False
    LoginRequest.setRequestHeader "Content-type", CabeceraHTTP
    LoginRequest.send (CargaUtil)

''' Capturamos la descarga y transerimos
    sResponseText = LoginRequest.ResponseText
Debug.Print sResponseText
    Salida = sResponseText
'GoTo continua

' ====================================================================================
'                             Nota sobre este procedimiento
' ====================================================================================
'   Sé que es mejorable usando el módulo de JSON, pero me ha dado tantos problemas
'   a la hora de trabajar con matrices dentro del archivo JSON "[ ... ]", que para
'   una solución de urgencia he decidirlo hacerlo a mano.
'   El problema surgirá si esta matriz contine múltiples usuarios, no serán mostrados.
' ====================================================================================

''' El resultado lo convierto en una matriz para poder extraer los datos requeridos
'   Si la respuesta del servidor está basada en la petición de datos de usuario
'   aparecerán sub-informaciones con las cabeceras: users, providerUserInfo

If Accion = "INFO" Or Accion = "UPDATE" Then
    Dim Matrices As Double
    Dim Caracteres As Double
    Dim Caracter As String
    Dim i As Long
    Dim BuscarIndex  As Long
    Dim BuscarMatriz As Long
    Dim providerUserInfo As String
    Dim ProviderInfo As Variant
    Caracteres = Len(sResponseText)

    BuscarIndex = InStr(1, sResponseText, "providerUserInfo")
    
    If BuscarIndex > 0 Then
        Salida = Replace(sResponseText, Mid(sResponseText, BuscarIndex - 1, InStr(BuscarIndex, sResponseText, "]") - BuscarIndex), "")
        BuscarIndex = InStr(BuscarIndex, sResponseText, "[")
        BuscarMatriz = InStr(BuscarIndex, sResponseText, "]")
        providerUserInfo = Mid(sResponseText, BuscarIndex + 1, BuscarMatriz - BuscarIndex)
        TextoLimpio = Replace(providerUserInfo, "providerUserInfo", "")
        TextoLimpio = Replace(Replace(Replace(TextoLimpio, ": [", ""), "],", ""), "]", "")
        TextoLimpio = Replace(Replace(Replace(Replace(TextoLimpio, "{", ""), "}", ""), Chr(10), ""), Chr(34), "")
        ProviderInfo = Split(TextoLimpio, ",")
        nCadenas = UBound(ProviderInfo)
        ReDim ArrayAuth(nCadenas, 1)
        For n = 0 To nCadenas
            ArrayAuth(n, 0) = "puI" & Trim(Left(ProviderInfo(n), InStr(1, ProviderInfo(n), ":") - 1))
            ArrayAuth(n, 1) = Trim(Mid(ProviderInfo(n), InStr(1, ProviderInfo(n), ":") + 1))
        Next n
        a = n                           ' Almacenamos el numerod de valores guardados para usarlo más adelante.
        If nCadenas > 0 Then Erase ProviderInfo
    End If
    
    
    'Debug.Print providerUserInfo
    Debug.Print "================================================"
    Debug.Print Salida
    Debug.Print "================================================"

End If
' ====================================================================================
'   FIN DEL ARREGLO SIN USAR EL MÓDULO DE JSON
' ====================================================================================


    If InStr(1, Salida, "[") > 0 Then
        TextoLimpio = Replace(Replace(Salida, "users", ""), "providerUserInfo", "")
        TextoLimpio = Replace(Replace(TextoLimpio, ": [", ""), "]", "")
        TextoLimpio = Replace(Replace(Replace(Replace(TextoLimpio, "{", ""), "}", ""), Chr(10), ""), Chr(34), "")
        TextoLimpio = Replace(TextoLimpio, " , ", "")
    Else
        TextoLimpio = Replace(Replace(Replace(Replace(sResponseText, "{", ""), "}", ""), Chr(10), ""), Chr(34), "")
    End If
    
    'Debug.Print providerUserInfo
    Debug.Print "===============================================--------"
    Debug.Print TextoLimpio
    
continua:
    If InStr(1, TextoLimpio, "html") > 0 Then GoTo SinConexion
    Cadenas = Split(TextoLimpio, ",")
    nCadenas = arrayLength(Cadenas)
    If nCadenas = 0 Then GoTo SinConexion
    ReDim ProviderInfo((nCadenas - 1) + a, 1)
    For n = 0 To nCadenas - 1                   ' Añadimos los valores de Usuario
        ProviderInfo(n, 0) = Trim(Left(Cadenas(n), InStr(1, Cadenas(n), ":") - 1))
        ProviderInfo(n, 1) = Trim(Mid(Cadenas(n), InStr(1, Cadenas(n), ":") + 1))
    Next n
    If a > 0 Then                               ' Añadimos los valores de Provider
        For n = (nCadenas - 1) To (nCadenas - 1) + a - 1
            ProviderInfo(n, 0) = ArrayAuth(n - (nCadenas - 1), 0)
            ProviderInfo(n, 1) = ArrayAuth(n - (nCadenas - 1), 1)
        Next n
    End If
    AccionConUsuario = ProviderInfo
    Exit Function
    
SinConexion:
    ReDim ArrayAuth(0, 1)
        ArrayAuth(n, 0) = "idToken"
        ArrayAuth(n, 1) = "Empty"
        AccionConUsuario = ArrayAuth

End Function

' ===================================================================================================
'   HERRAMIENTAS COMUNES
' ===================================================================================================
'   ParametroDB: Devuelve parámetros de configuración de Firebase.
'   arrayLength: Devuelve la longitud de un array desde 0.
'   CheckJSON:   Chequea la cadena JSON. Devuelve True si es correcta.

''' ParametroDB permite extraer en Formularios creados en EXCEL algún
'   dato necesario para realizar alguna operación específica.
'   Usada en algunos caso para insertar en el cadena codificada un valor necesario:
'   1 -> Nombre de la base de datos
'   2 -> Direccion de la base de datos
'   3 -> API de la base de datos
Function ParametroDB(Valor As Single) As String
    If Valor = 1 Then
        ParametroDB = dbNAME
    ElseIf Valor = 2 Then
        ParametroDB = dbURL
    ElseIf Valor = 3 Then
        ParametroDB = dbAPI
    End If

End Function

''' arrayLength es una herramienta que permite obtener el número exacta de valores
'   contenidos en un array desde 0 al máximo (no desde -1 si está vacío).
 Function arrayLength(Arr As Variant) As Long
  On Error GoTo handler

  Dim lngLower As Long
  Dim lngUpper As Long

  lngLower = LBound(Arr)
  lngUpper = UBound(Arr)

  arrayLength = (lngUpper - lngLower) + 1
  Exit Function

handler:
  arrayLength = 0 'error occured.  must be zero length
End Function

Function CheckJSON(cadenadetexto As String) As Boolean
    Dim sState As String
    Dim vJSON
    Dim vFlat
    CheckJSON = True
    JSON.Parse cadenadetexto, vJSON, sState
    If sState = "Error" Then CheckJSON = False
End Function
