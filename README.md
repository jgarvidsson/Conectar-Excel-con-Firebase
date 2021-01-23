# Conectar Excel con Firebase

**Índice**
1. [Antes de Empezar](#id1)
2. [Las Credenciales](#id2)



## Antes de Empezar <div id='id1' />

Antes de empezar es necesario remarcar que para conectar **Firebase** con **EXCEL** se requiere tener creada y configurada una base de datos.

El ejemplo adjunto en este repositorio, crea una carpeta en "Mis Documentos" (Documents) con el nombre de **fbExcel** donde se guardará algún datos que necesite ser descargado (foto de perfil o archivo JSON). Este procedimiento lo uso habitualmente para cuando finalice con las pruebas, poder borrar todos los archivos generados de manera más sencilla. Para cambiar el nombre de esta carpeta, abra el modulo **Herramientas** y cambie el contenido de la variable **NombreCarpetaTrabajo** por el que desee.

<div id='id2' />
## Las Credenciales

Para trabajar con Firebase, además de los datos principales proporcionados por el servidor y que son necesarios para que Firebase reconozca tu aplicación como app con privilegios de Administrador


### Credenciales del Servidor
Se precisa la siguiente información por parte del servidor:

		Private Const dbNAME As String = "<nombre de la base de datos>"                             ' Nombre de la base de datos (sin cabecera http ni servidor).
		'Private Const dbURL As String = "https://" & dbNAME & ".firebaseio.com/"                   ' Direccion de la base de datos.    ' Servidor en USA
		Private Const dbURL As String = "https://" & dbNAME & ".europe-west1.firebasedatabase.app/" ' Direccion de la base de datos.    ' Servidor en Europa
		Private Const dbAPI As String = "<API>"                   									' API de la base de datos

El servidor puede estar localizado en USA o EU, y es posible que aparezcan otras posibilidades en el futuro. Confirmad antes cual es la dirección donde almacenareis vuestra BD (lo selecionais en la configuración) y dejar habilitada una de las dos **dbURL**

#### Las Reglas de Seguridad de Firebase
Es muy importante recordar que, dependiendo de las reglas de Uso configuradas para su base de datos, el registro o visualización de datos requerirá de credenciales específicas. Esta **Reglas** se configuran en la *consola de Firebase*, en el menú *Reglas* dentro de *RealTime Database*.

Sin configuramos las Reglas de las siguiente manera.

		{
		  "rules": {
		    ".read": true, 
		    ".write": true,  
		  }
		}

Cualquiera que tenga las Credenciales del Servidor (API y dirección del Servidor), podrá acceder a los datos.

En cambio, si tenemos en cuenta nuestras necesidades podremos configurar dichas reglas de tal menera que sólo podamos acceder a los datos registrados cuando estemos 'logeados' como **Usuario** de la base de datos.

		{
			"rules": {
		     "Conexion": {".write": "auth != null",".read": true},
		     "Test":     {".write": "auth != null",".read": "auth != null"}
								}
		}

Iniciamos la regla con 'rules' para que el servidor sepa qué estamos configurando. Como ejemplo de este repositorio se ha creado dos directorios dentro de la base de datos: 'Conexion' y 'Test'.

  - 'Conexion' sólo contendrá un valor con un contenido booleano y lo usaremos para comprobar que nuestra aplicación tiene acceso a la base de datos.
  - 'Test' contendrá todas los registros de datos realizados en el ejemplo.

 Si quereis probar con otros nombres o añadir más no habrá problemas siempre y cuando respeteis las **Reglas** de seguridad configuradas.

### Credenciales del Usuario
Las credenciales del **Usuario** son un Correo Electrónico y un Password. Estos datos los suministra el Administrador o puede activarse desde un Formulario creado en EXCEL a través de una de las Funciones que podreis ver más adelante. En el  archivo de ***Test*** contenido en este repositorio, las credenciales pueden integrarse en el programa a través de variables privadas o introducirse a través de *cajas de texto*.

En mis proyectos EXCEL<->FireBase uso un archivo encriptado que contiene la información del Servidor y del Usuario. El Usuario final no necesita conocer esta información porque no va a usarla con fines personales, sólo es una llave de acceso a una información contenida en el servidor y que será como intercambiador de datos y centralización de información. Yo les paso el archivo Excel y un 'permit' y pueden trabajar sin problemas.

## Módulos VBA (.bas)

Este repositorio cuenta con tres módulos necesarios para realizar todas las operaciones de conexión, envío y recepción de datos con **Firebase** usando **EXCEL**.
- **Módulo Firebase**: Contiene las funciones necesarias para realizar la acciones requeridas.
- **Módulo JSON**: Contiene las funciones necesarias para leer, escribir y validar archivos JSON (repositorio original https://github.com/omegastripes/VBA-JSON-parser).
- **Módulo de Herramiemtas**: Contiene funciones extras que ayudan en algunas operaciones dentro de **EXCEL**.

### Módulo Firebase
Es un módulo escrito a partir de variaciones del Módulo JSON y las indicaciones de la Web de Firebase. Contiene los instrumentos necesarios para realizar el conexionado con la Base de Datos (con o sin autorización, dependerá del tipo de configuración del usuario).

Cuando se cree un Formulario de VBA en EXCEL hay que recordar que el primer paso es requerir la autorización de conexión a Firebase, par ello se debe comenzar el módulo con la variable privada que contendrá dicha autorización, al menos, en el formulario principal:

      Activate TokenAutorizacion As String

A continuación, deberá decidir si el Token se requerirá automáticamente o manual.

- Si es **automático** (con la apertura del Formulario) deberá incluir el codigo siguiente:

      Private Sub UserForm_Initialize()
          Dim user As String
          Dim pass As String
          
          emailuser = user
          passuser = pass
          If ComprobarConexion = True Then
              MostrarEstado NetStatus, "Connection with Server OK", 2
              TokenAutorizacion = DevolverValorFirebase("idToken", user, pass)           ' Variable privada - Devuelve el token de conexión a iMerlin.
          Else
              MostrarEstado NetStatus, "There is not Server Connection", 3
          End If
      End Sub

- Si es **manual** (realizando una acción, como presionar un botón), podríamos resolver la petición del *token* de la siguiente creando un botón *Conectar* y usando el código indicado a continuación (este código es parte del archivo adjunto a este repositorio):

      Private Sub Conectar_Click()
          If ComprobarConexion = True Then
              MostrarEstado NetStatus, "Connection with Server OK", 2
              TokenAutorizacion = DevolverValorFirebase("idToken", emailuser, passuser)
          Else
              MostrarEstado NetStatus, "There is not Server Connection", 3
          End If
      End Sub

Además de **DevolverValorFirebase** que especifica el contenido de un valor "idToken" para extraer el ***Token***, se puede usar la siguiente función:

      Function RecibirAutorizacion(IDUsereMail As String, IDUserPassword As String) As Variant

Devuelve un array (matriz) con el contenido de la respuesta del servidor. Con esta función se pueden obtener los detalles de un error y el mensaje devuelto.


#### Enviar información a la base de datos
##### Carga Útil (o PayLoad)
Hay que tener mucho cuidado a la hora de generar las cadenas con los datos en formato JSON. Es un error muy normal no acertar en la composición de una cadena completa, hay que tener especial ojo en las características que tiene, ya que estos errores provocan el rechazo por parte del servidor, por lo que puede marearnos un poco.

Estructura JSON:

      Mensaje = "{" & _
      """Domain"":""" & Environ("Userdomain") & """," & _
      """Workbook"":""" & ThisWorkbook.Name & """," & _
      """" & Valor & """:""" & Contenido & """," & _
      """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
      """TSS"":{"".sv"":""timestamp""}" & _
      "}"

#### Función para trabajar con Firebase Realtime
La Función principal para trabajar con la base de datos de *Fiebase* es **FirebaseDB**. Esta función, y dependiendo del modo de trabajo requerido, permite trabajar con la base de datos en modo online realizando varias tareas especícicas.

    Function FirebaseDB(Mode As String, Direccion As String, Mensaje As String, Optional claveautorizacion As String = "", Optional SoloContenidoIndice As Boolean = False) As Variant
    
-  **Mode** representa el tipo de trabajo que se va a realizar (ver siguiente lista).
-  **Direccion** indica el nombre del árbol principal o Valor que contendrá la información con la que se quiere trabajar en la base de datos.
-  **Mensaje** contendrá el mensaje en formato JSON que será enviado al servidor. En algunos **MODOS** contendrá la dirección de descarga en el PC o la dirección en la base de datos donde los datos serán transferidos (ver descripción de los **MODOS** para más información).
-  **claveautorizacion** contendrá el token de autorización, dependiendo de la seguridad de la base de datos, para trabajar en la base de datos. Si esta no está protegida para lectura y/o escritura no será necesaria.
-  **SoloContenidoIndice** indicará que en la descarga de datos se descargue sólo los valores (sin el contenido). Por defecto es "false".
    
##### Los Modos de Trabajo de la Función FirebaseDB
  - **POST**     - Postear (añade un ID al JSON posteado y retorna dicho ID)
  - **PATCH**    - Agregar mensaje JSON a una ***Dirección*** sin borrar la información existente (actualizar).
  - **PUT**      - Añadir valor JSON a una ***Dirección***.
  - **GET**      - Recibir contenido requerido en la ***Dirección*** indicada. En este caso ***Mensaje*** irá vacío = ""
  - **DELETE**   - Borra el contenido de la ***Dirección*** indicada - En este caso ***Mensaje*** irá vacío = ""
  - **DOWNLOAD** - Descarga en un archivo JSON el contenido de ***Dirección*** (en formato JSON). En este caso en ***Mensaje*** irá la *carpeta de destino + nombre de archivo + extensión de destino*.
  - **BACKUP**   - Copia (y duplica) el contenido de la ***Dirección*** indicada a la dirección indicada en ***Mensaje*** en modo "**POST**". Resuelve el problema de las copias de seguridad cuando un usuario no administrador realiza un cambio en un valor de la BD permitiendo que pueda ser recuperado en otro momento.
  - **MOVE**     - Mueve el contenido de la ***Dirección*** indicada a la dirección indicada en ***Mensaje***.
  - **COPY**     - Copia (y duplica) el contenido de la ***Dirección*** indicada en la dirección indicada en ***Mensaje***.


#### Modos de envío
Hay tres formas de enviar información a la base de datos:
- 1) Con el modo *POST*
- 2) Con el modo *PUT*
- 3) Con el modo *PATCH*

Cada modo funciona de manera similar, pero coloca la información de una manera específica.
- Con **POST** enviaremos una carga util de datos en formato JSON que será localizada en una dirección específica usando un identificador *temporal* que organizará los datos en el orden en el que se suban. Como cuando posteas en una red social.

El botón *POST* del ejemplo adjunto en este repositorio, sería:

      Private Sub dbPost_Click()
          Dim Direccion As String
          Dim Mensaje As String
          Dim Respuesta As Variant

          Direccion = Caminofb.Text            ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

      ''' Creamos la carga util con el mensaje o datos que queremos enviar
          Mensaje = "{" & _
          """Domain"":""" & Environ("Userdomain") & """," & _
          """Workbook"":""" & ThisWorkbook.Name & """," & _
          """" & Valor & """:""" & Contenido & """," & _
          """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
          """TSS"":{"".sv"":""timestamp""}" & _
          "}"

          Respuesta = FirebaseDB("POST", Direccion, Mensaje, TokenAutorizacion)

      End Sub
      
- Con **PUT** enviaremos los datos y se localizarán en la dirección indicada. Si se vuelve a enviar otra carga útil borrará el registro anterior sustituyendo los valores diferentes y eliminando los que ya no están incluidos. 

El botón *PUT* del ejemplo adjunto en este repositorio, sería:

      Private Sub dbPut_Click()
          Dim Direccion As String
          Dim Mensaje As String
          Dim Respuesta As Variant

          Direccion = Caminofb.Text            ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

      ''' Creamos la carga util con el mensaje o datos que queremos enviar
          Mensaje = "{" & _
          """Domain"":""" & Environ("Userdomain") & """," & _
          """Workbook"":""" & ThisWorkbook.Name & """," & _
          """" & Valor & """:""" & Contenido & """," & _
          """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
          """TSS"":{"".sv"":""timestamp""}" & _
          "}"

          Respuesta = FirebaseDB("PUT", Direccion, Mensaje, TokenAutorizacion)
          
      End Sub


- Con **PATCH** enviaremos la carga util, se añadirán nuevos valores sin tocar los que ya estaban y no están incluidos en el envío a no ser que tengan el mismo nombre, lo que hará que se actualice el contenido de ese valor.

El botón *PATCH* del ejemplo adjunto en este repositorio, sería:

      Private Sub dbPatch_Click()
          Dim Direccion As String
          Dim Mensaje As String
          Dim Respuesta As Variant

          Direccion = Caminofb.Text            ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

      ''' Creamos la carga util con el mensaje o datos que queremos enviar
          Mensaje = "{" & _
          """Domain"":""" & Environ("Userdomain") & """," & _
          """Workbook"":""" & ThisWorkbook.Name & """," & _
          """" & Valor & """:""" & Contenido & """," & _
          """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
          """TSS"":{"".sv"":""timestamp""}" & _
          "}"

          Respuesta = FirebaseDB("PATCH", Direccion, Mensaje, TokenAutorizacion)
          
      End Sub


#### Modos de Recepción
Para recibir los datos desde Firebase se debe requerir directamente la dirección que contiene la información o el árbol de datos, recibiendo el dato contenido o los datos (valores y contenidos) respectivamente.

- El modo de recepción oficial es **GET**.


		Private Sub recibirDB_Click()
		    Dim Direccion As String
		    Dim Mensaje As String
		    Dim Respuesta As Variant
		    Dim nRespuesta As Single
		    Dim ValorConexion As String
		    Dim i As Single
		        
		    BorrarCampos                    ' Borrar los campos TextBox de la acción Recibir (es sólo por limpiar la ventana)
		    
		    Direccion = Caminofb.Text       ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.
		    RespuestaServ2 = "Procensando"
		    DoEvents


		''' Enviamos la petición. Se le ha asignado una variable para soportar la carga de devolución de la base de datos,
		'   con esta información podemos controlar los datos devueltos por el servidor.
		    Respuesta = FirebaseDB("GET", Direccion, Mensaje, TokenAutorizacion)
		    nRespuesta = arrayLength(Respuesta)
		    If nRespuesta = 0 Then Exit Sub
		    
		    On Error Resume Next
		    For i = 0 To nRespuesta - 1
		        If Respuesta(i, 0) = Valor2 Then Contenido2.Text = Respuesta(i, 1)
		        If Respuesta(i, 0) = "TSL" Then TimeLocal.Text = Respuesta(i, 1)
		        If Respuesta(i, 0) = "TSS" Then TimeServer.Text = Respuesta(i, 1)
		        If Respuesta(i) = "Disconnected" Then ValorConexion = Respuesta(i)
		        If Respuesta(i) = "null" Then ValorConexion = Respuesta(i)
		    Next i

		''' Dependiendo de la recepción, podremos definir si todo salió bien, o hubo un problema
		    If ValorConexion = "Disconnected" Then
		        RespuestaServ2 = "No Tiene Permiso"
		    ElseIf ValorConexion = "null" Then
		        RespuestaServ2 = "No hay datos para descargar"
		    Else
		        RespuestaServ2 = "Datos Recibidos"
		    End If
		fin:
		    If nRespuesta > 0 Then Erase Respuesta
		    On Error GoTo 0
		    
		End Sub

Como respuestas tendremos que:
  - Si el **Usuario** no tiene permisos recibirá un Valor=Disconnected.
  - Si la *dirección* de origen no existe recibirá un Valor = null.
  - Si todo va bien recibirá un Array con el contenido. 

- El modo personalizado **DOWNLOAD**, permite descargar en formato JSON la información contenida sobre un valor o árbol de datos de *Firebase*.

El botón *DOWNLOAD* del ejemplo adjunto en este repositorio, sería:

      Private Sub DescargarDB_Click()
          Dim Direccion As String
          Dim ArchivoPC As String
          Dim Respuesta As Variant

          Direccion = Caminofb2.Text                  ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.
          ArchivoPC = RutaCarpetasEspeciales(0) & _
                      "\" & Caminofb2 & ".json"       ' Indicamos que el archivo se descargue en la misma carpeta que el archivo EXCEL

      '   Para borrar, la variable Mensaje va en blanco o simplemente se le añade un "".
          Respuesta = FirebaseDB("DOWNLOAD", Direccion, ArchivoPC, TokenAutorizacion)
      End Sub

#### Modo de borrado
Se puede eliminar un valor y su contenido directamente o un árbol completo de datos usando el modo **DELETE**.

- El botón *BORRAR* del ejemplo adjunto en este repositorio, sería:

      Private Sub BorrarPath_Click()
          Dim Direccion As String
          Dim Mensaje As String
          Dim Respuesta As Variant

          Direccion = Caminofb.Text            ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

          Respuesta = FirebaseDB("DELETE", Direccion, Mensaje, TokenAutorizacion)
      End Sub

#### Otros modos de trabajo (personalizados)
Para realizar distintas tareas dentro del árbol de datos de *Firebase* tuve que ampliar la función de trabajo **FirebaseDB** añadiéndole otros modos para aumentar su funcionalidad.

El modo personalizado **BACKUP** permite realizar una copia de un dato o árbol de datos a otra parte de la base de datos usando el modo **POST** de manera que se guarda una copia de seguidad de los cambios realizados en caso de querer volver a revisarlos más tarde.

- El botón *BACKUP* del ejemplo adjunto en este repositorio, sería:

      Private Sub BackUpDB_Click()
          Dim Direccion As String
          Dim CopiarDB As String
          Dim Respuesta As Variant

          Direccion = Caminofb2.Text          ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.
          CopiarDB = DestinoCopiar            ' Indicamos que el archivo se copie en la dirección de la DB indicada

          Respuesta = FirebaseDB("BACKUP", Direccion, CopiarDB, TokenAutorizacion)

      End Sub

- El modo personalizado **MOVE** permite mover un dato o árbol de datos a otra parte de la base de datos.

El botón *MOVER* del ejemplo adjunto en este repositorio, sería:

      Private Sub MoverDB_Click()
          Dim Direccion As String
          Dim CopiarDB As String
          Dim Respuesta As Variant

          Direccion = Caminofb2.Text          ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.
          CopiarDB = DestinoCopiar            ' Indicamos que el archivo se copie en la dirección de la DB indicada

          Respuesta = FirebaseDB("MOVE", Direccion, CopiarDB, TokenAutorizacion)

      End Sub
      
- El modo personalizado **COPY** permite mover un dato o árbol de datos a otra parte de la base de datos duplicando dicha información.

El botón *COPIAR* del ejemplo adjunto en este repositorio, sería:

      Private Sub CopiarDB_Click()
          Dim Direccion As String
          Dim CopiarDB As String
          Dim Respuesta As Variant

          Direccion = Caminofb2.Text          ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.
          CopiarDB = DestinoCopiar            ' Indicamos que el archivo se copie en la dirección de la DB indicada

          Respuesta = FirebaseDB("COPY", Direccion, CopiarDB, TokenAutorizacion)

      End Sub

## Funciones de Registro
Las funciones de registro son comandos que procesan ciertos datos y los envía al servidor. Estas funciones son:
- **RegistrarUso**: Registra una cadena JSON con datos específicos en modo ***POST***. De esta manera se puede llevar un control del uso de una aplicación.
- **GenerarJSONError**: Registra los errores que puedan aparecer en la aplicación en modo ***POST***. Si por cuaquier motivo aparece un error de conexión, la cadena JSON que contiene el error se guarda en una carpeta especificada. Se puede crear una Función que intente enviar el contenido de archivos generados en otro momento (no incluido en este repositorio).

## FUNCIONES DE USUARIO
Para trabajar con usuarios usaremos la siguiente funcion:

Function AccionConUsuario(Accion As String, IDUsereMail As String, IDUserPassword As String, _
                            Optional IDTokenUser As String = "", Optional IDNameUser As String = "", _
                            Optional IDURLFoto As String = "") As Variant

  - **Accion**: Se indicará que acción se llevará a cabo por el usuario. Para más información, ver la lista de acciones más abajo (Valor Obligatorio).
  - **IDUsereMail**: Se indicará el correo electronico del **Usuario** (Valor Obligatorio).
  - **IDUserPassword**: Se indicará el password de la sesión de **Usuario** (Valor Obligatorio).
  - **IDTokenUser**: Se suministrará el IdToken cuando se realice alguna acción sobre la información del **Usuario** (Valor opcional).
  - **IDNameUser**: Se indicará el Nombre de **Usuario** que será mostrado por el servidor (Valor opcional).
  - **IDURLFoto**: Se indicará la dirección Web de la foto de perfil del **Usuario** (Valor opcional).

### Acciones de Usuario
Para indicar qué acción se llevará a cabo, se indicará la palabra clave correspondiente:
  - **NEW**: Creará un nuevo usuario a través del *Correo Electrónico* y una *clave de acceso*.

  		AccionConUsuario("NEW", eMail, Password)
  - **ANONIMUS**: Permite crear un Usuario Anonimo con las mismas características de un Usuario Registrado. El IdToken generado cadurá pasada una hora. Se puede actualizar de ***ANONIMUS*** a ***Usuario Registrado*** usando el **IdToken** generado para el primero y actualizando los datos con la **Accion UPDATE**.

  		AccionConUsuario("ANONIMUS", "", "")
  - **INFO**: Recupera los datos del Usuario cuyo IdToken esté activo.

  		AccionConUsuario("INFO", eMail, Password)
  - **UPDATE**: Actualiza la información de un Usuario excepto la dirección de Correo Electrónico. Si actualiza también refresca el IdToken, pero tiene que ser antes de que caduque (tienen una vida de 3600 segundos).

  		AccionConUsuario("UPDATE", eMail, Password, IdToken, Nombre, URLfoto)
  - **AUTH**: Recupera el IdToken de un **Usuario Registrado**.

  		AccionConUsuario("AUTH", eMail, Password)
  - **REMOVE**: Borra el registro de un **Usuario**.

		AccionConUsuario("REMOVE", eMail, Password)

La función **AccionConUsuario** devuelve una matriz con los datos extraidos y pueden ser tomados usando un bucle ***For...Next simple***. En el siguiente ejemplo se muestra la acción de Crear Nuevo Usuario del archivo de Test incluido en este repositorio.


		Private Sub CrearUsuarioNuevo_Click()
		''' Declaramos las variables
		    Dim Respuesta As Variant
		    Dim nRespuesta As Single
		    Dim i As Single
		        
		''' Enviamos el requerimiento al servidor y chequeamos el contenido de la respuesta
		    Respuesta = AccionConUsuario("NEW", nuevoMail, nuevoPass, "", nuevoNombre, nuevoFoto)
		    nRespuesta = arrayLength(Respuesta)
		    If nRespuesta = 0 Then Exit Sub
		    
		''' Si hay datos de respuesta, desplegamos la matriz visualmente
		    For i = 0 To nRespuesta - 1
		        If Respuesta(i, 0) = "kind" Then kind = Respuesta(i, 1)
		        If Respuesta(i, 0) = "email" Then email = Respuesta(i, 1)
		        If Respuesta(i, 0) = "error" Then sCodigo = Extraer(Respuesta(i, 1), False)
		        If Respuesta(i, 0) = "message" Then sMensaje = Respuesta(i, 1)
		        If Respuesta(i, 0) = "reason" Then sEstatus = Respuesta(i, 1)
		        If Respuesta(i, 0) = "idToken" Then _
		            sCodigo = "200": TokenAutorizacion = Respuesta(i, 1): _
		            Me.Caption = "Comunicación con FireBase - (ID Token para " & nuevoMail & ") -"
		        If Respuesta(i, 0) = "registered" Then sEstatus = Respuesta(i, 1)
		        If Respuesta(i, 0) = "kind" Then sMensaje = "USER CREATED!!!"
		    Next i
    		If nRespuesta > 0 Then Erase Respuesta
		End Sub


## ERRORES
Cuando enviamos información al servidor para realizar una petición, en respuesta se recibe un código, este código puede significar que la petición fue aceptada o presentó un error:
  - **200**: Petición aceptada.
  - **400**: El servidor no ha podido procesar la petición porque hay un error.

 | Error | Descripción | Acción |
 | :---: | --- | --- |
 | **EMAIL_NOT_FOUND** | El eMail introducido por el usuario no está registrado. | Comprueba la sintaxis de los datos introducidos o ponte en contacto con el **Administrador** para registrar tu **Usuario**. |
 | **INVALID_PASSWORD** | El password correspondiente al eMail introducido no es correcto | Compruebe la sintaxis de los datos introducidos o realice una petición de recuperación al Administrador. |
 | **INVALID_ID_TOKEN** | Está intentando realizar una acción sin estar identificado o con un IdToken diferente al **Usuario** indicado. | Inicie sesión con sus credenciales para realizar la acción deseada. |
 | **CREDENTIAL_TOO_OLD_LOGIN_AGAIN** | Ha intentado realizar alguna acción en el servidor con una credencial caducada. | Vuelva a conextarse para actualizar el IdToken de **Usuario**. |
