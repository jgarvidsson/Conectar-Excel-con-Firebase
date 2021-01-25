# Conectar Excel con Firebase

**Índice**
1. [Antes de Empezar](#id1)
2. [Las Credenciales](#id2)
   - [Credenciales del Servidor](#id3)
   - [Credenciales del Usuario](#id4)
   - [Las Reglas de Seguridad de Firebase](#id5)
3. [Los Archivos Adjuntos](#id6)
   - [Módulos VBA (.bas)](#id7)
   - [Archivo EXCEL (.xlsm)](#id8)
4. [Conexión con Firebase](#id9)
   - [IdToken](#id10)
   - [Carga Util (Payload)](#id11)
5. [Funciones](#id12)
   - [FirebaseDB](#id13)
     - [Modo POST](#id14)
     - [Modo PATCH](#id15)
     - [Modo PUT](#id16)
     - [Modo GET](#id17)
     - [Modo DELETE](#id18)
     - [Modo DOWNLOAD](#id19)
     - [Modo BACKUP](#id20)
     - [Modo MOVE](#id21)
     - [Modo COPY](#id22)
   - [DevolverValorEspecificoDeFirebase](#id23)
   - [DevolverValorAutorizacion](#id24)
   - [FirebasePC](#id25)
   - [DevolverValorEspecificoDeJSONLocal](#id26)
   - [Funciones de Registro](#id27)
     - [GenerarJSONError](#id28)
     - [RegistrarUso](#id29)
     - [AlmacenarJSON](#id30)
   - [Acciones para USUARIO](#id31)
     - [Crear Nuevo Usuario](#id32)
     - [Crear Usuario Anonimo](#id33)
     - [Requerir Información de un Usuario](#id34)
     - [Actualizar Información de un Usuario](#id35)
     - [Activar IdToken de un Usuario](#id36)
     - [Borrar a un Usuario](#id37)
     - [Petición de Nuevo Password de Usuario](#id38)
     - [Cambiar eMail](#id39)
     - [Cambiar Password](#id40)
     - [Qué responde el servidor](#id41)
   - [ComprobarConexion](#id42)
6. [ERRORES](#id43)
   - [Lista de Errores Comunes](#id44)



<div id='id1' />

## Antes de empezar

Antes de empezar es necesario remarcar que para conectar **Firebase** con **EXCEL** se requiere tener creada y configurada una base de datos.

El ejemplo adjunto en este repositorio (archivo EXCEL), crea una carpeta en "Mis Documentos" (Documents) con el nombre de **fbExcel** donde se guardarán datos que necesiten ser descargado (foto de perfil o archivo JSON). Este procedimiento lo uso habitualmente para cuando al finalizar con las pruebas, poder borrar todos los archivos generados de manera más sencilla. Para cambiar el nombre de esta carpeta, abra el modulo **Herramientas** y cambie el contenido de la variable **NombreCarpetaTrabajo** localizado en la cabecera por el que desee.

		Private Const NombreCarpetaTrabajo As String = "fbExcel"


**Por favor, tened en cuenta que hay que tener un conocimiento medio del manejo de VBA para poder tocar las funciones**.

<div id='id2' />

## Las Credenciales
Para trabajar con Firebase, se requieren credenciales de conexión proporcionadas por el **Servidor** para permitir la comunicación y las credenciales de **Usuario** para poder trabajar con los datos.

<div id='id3' />

### Credenciales del Servidor
Se precisa la siguiente información por parte del servidor:

		Private Const dbNAME As String = "<nombre de la base de datos>"                             ' Nombre de la base de datos (sin cabecera http ni servidor).
	   'Private Const dbURL As String = "https://" & dbNAME & ".firebaseio.com/"                    ' Direccion de la base de datos.    ' Servidor en USA
		Private Const dbURL As String = "https://" & dbNAME & ".europe-west1.firebasedatabase.app/" ' Direccion de la base de datos.    ' Servidor en Europa
		Private Const dbAPI As String = "<API>"                   									' API de la base de datos

El servidor puede estar localizado en USA o EU, y es posible que aparezcan otras posibilidades en el futuro. Confirmad antes cual es la dirección donde almacenareis vuestra BD (lo selecionais en la configuración) y dejar habilitada una de las dos **dbURL**

<div id='id4' />

### Credenciales de Usuario
Para poder manejar los datos del servidor, se requierirá que el Usuario esté registrado en la base de datos.
Las credenciales del **Usuario** se componen de un Correo Electrónico y un Password. Estos datos los suministra el Administrador o puede activarse desde un Formulario creado en EXCEL a través de una de las Funciones que podreis ver más adelante. En el  archivo de ***Test*** contenido en este repositorio, las credenciales pueden integrarse en el programa a través de variables privadas o introducirse a través de *cajas de texto*.

En mis proyectos EXCEL<->FireBase uso un archivo encriptado que contiene la información del Servidor y del Usuario. El Usuario final no necesita conocer esta información porque no va a usarla con fines personales, sólo es una llave de acceso a una información contenida en el servidor y que será como intercambiador de datos y centralización de información. Yo les paso el archivo Excel y un 'permit' y pueden trabajar sin problemas.

<div id='id5' />

### Las Reglas de Seguridad de Firebase
Es muy importante recordar que, dependiendo de las reglas de Uso configuradas para su base de datos, el registro o visualización de datos requerirá de credenciales específicas. Esta **Reglas** se configuran en la *consola de Firebase*, en el menú *Reglas* dentro de *RealTime Database*.

Si configuramos las Reglas de las siguiente manera.

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

Iniciamos la regla con 'rules' para que el servidor sepa qué estamos configurando. Como ejemplo de este repositorio se han creado dos directorios dentro de la base de datos: 'Conexion' y 'Test'.

  - 'Conexion' sólo contendrá un valor con un contenido booleano y lo usaremos para comprobar que nuestra aplicación tiene acceso a la base de datos.
  - 'Test' contendrá todas los registros de datos realizados en el ejemplo.

 Si quereis probar con otros nombres o añadir más no habrá problemas siempre y cuando respeteis las **Reglas** de seguridad configuradas.


<div id='id6' />

## Los Archivos Adjuntos

<div id='id7' />

### Módulos VBA (.bas)

Este repositorio cuenta con tres módulos necesarios para realizar todas las operaciones de conexión, envío y recepción de datos con **Firebase** usando **EXCEL** y un ejemplo complementado en un archivo de **EXCEL** con todas las funciones explicadas aquí.Contiene los instrumentos necesarios para realizar el conexionado con la Base de Datos (con o sin autorización, dependerá del tipo de configuración del usuario) y el manejo del registro de **Usuarios**.


- **Módulo Firebase**: Contiene las funciones necesarias para realizar la acciones requeridas. Es el que se desarrollará en este texto.
- **Módulo JSON**: Contiene las funciones necesarias para leer, escribir y validar archivos JSON (repositorio original https://github.com/omegastripes/VBA-JSON-parser).
- **Módulo de Herramiemtas**: Contiene funciones extras que ayudan en algunas operaciones dentro de **EXCEL**.


<div id='id8' />

### Archivo EXCEL (.xlsm)
- Archivo EXCEL **Conectar con Firebase.xlsm**: Complementa todas las funciones explicadas y desarrolladas para un mejor entendimiento.


<div id='id9' />

## Conexión con Firebase

Antes de empezar hay algunos conceptos que tenéis que tener en cuenta. Cuando eres un ***Usuario Registrado en Firebase*** (no *Administrador*), necesitas un código para poder manejar la información de la Base de Datos, el **IdToken**. Estamos suponiendo que tu Base de Datos tiene **Reglas de Seguridad** que protegen la información, [como hemos visto anteriormente](#id5). En todas las explicaciones que se puedan dar en este repositorio, siempre será así.

Por otro lado tenemos la Carga Útil de información que enviaremos al Servidor cada vez que hagamos una petición. Este detalle es importante cuando manejamos información en la Base de Datos, ya que esta *carga útil* de información está estructurada con el protocolo [JSON](https://es.wikipedia.org/wiki/JSON), si la sintaxis del archivo se pierde, los datos no serán enviados correctamente.


<div id='id10' />

### IdToken

Cuando se crea un Formulario de VBA en EXCEL hay que recordar que el primer paso es requerir el código de autorización del **USUARIO**, el llamado **IdToken**, para ello se debe comenzar el módulo con la variable privada que contendrá dicho código:

      Private TokenAutorizacion As String

A continuación, deberá decidir si el Token se requerirá automáticamente o de forma manual.

- Si es **automático** (con la apertura del Formulario) deberá incluir el codigo siguiente:

      Private Sub UserForm_Initialize()
          Dim user As String
          Dim pass As String
          
          emailuser = user
          passuser = pass
          If ComprobarConexion = True Then
              MostrarEstado NetStatus, "Connection with Server OK", 2
              TokenAutorizacion = DevolverValorFirebase("idToken", user, pass)           ' Variable privada - Devuelve el token de conexión a la base de datos.
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

La función buleana **ComprobarConexion** que será explicada más adelante, devuelve un valor verdadero su se ha podido conectar con nuestra base de datos.

Además de **DevolverValorFirebase** que especifica el contenido de un valor "idToken" para extraer el ***Token***, se puede usar la siguiente función:

      Function RecibirAutorizacion(IDUsereMail As String, IDUserPassword As String) As Variant

Devuelve un array (matriz) con el contenido de la respuesta del servidor. Con esta función se pueden obtener los detalles de un error y el mensaje devuelto.


<div id='id11' />

### Carga Util (Payload)

Hay que tener mucho cuidado a la hora de generar las cadenas con los datos en formato JSON. Es un error muy normal no acertar en la composición de una cadena completa, hay que tener especial ojo en las características que tiene (https://es.wikipedia.org/wiki/JSON), ya que estos errores provocan el rechazo por parte del servidor, por lo que puede marearnos un poco.

Estructura JSON usada en este repositorio:

      Mensaje = "{" & _
      """Domain"":""" & Environ("Userdomain") & """," & _
      """Workbook"":""" & ThisWorkbook.Name & """," & _
      """" & Valor & """:""" & Contenido & """," & _
      """TSL"":""" & Format(Now(), "yyyy-MM-dd hh:mm:ss") & """," & _
      """TSS"":{"".sv"":""timestamp""}" & _
      "}"


<div id='id12' />

## Funciones

<div id='id13' />

### FirebaseDB

La Función principal para trabajar con la base de datos de *Fiebase* es **FirebaseDB**. Esta función, y dependiendo del modo de trabajo requerido, permite trabajar con la base de datos en modo online realizando varias tareas especícicas.

    Function FirebaseDB(Mode As String, Direccion As String, Mensaje As String, Optional claveautorizacion As String = "", Optional SoloContenidoIndice As Boolean = False) As Variant
    
-  **Mode** representa el tipo de trabajo que se va a realizar (ver siguiente lista).
-  **Direccion** indica el nombre del árbol principal o Valor que contendrá la información con la que se quiere trabajar en la base de datos.
-  **Mensaje** contendrá el mensaje en formato JSON que será enviado al servidor. En algunos **MODOS** contendrá la dirección de descarga en el PC o la dirección en la base de datos donde los datos serán transferidos (ver descripción de los **MODOS** para más información).
-  **claveautorizacion** contendrá el token de autorización, dependiendo de la seguridad de la base de datos, para trabajar en la base de datos. Si esta no está protegida para lectura y/o escritura no será necesaria.
-  **SoloContenidoIndice** indicará que en la descarga de datos se descargue sólo los valores (sin el contenido). Por defecto es "false".
    
#### Los Modos de Trabajo de la Función FirebaseDB
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


<div id='id14' />

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


<div id='id15' />
      
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


<div id='id16' />
      
- Con **PUT** enviaremos los datos y se localizarán en la dirección indicada. Si se vuelve a enviar otra carga útil con la misma dirección borrará el registro anterior sustituyendo los valores diferentes y eliminando los que ya no están incluidos. 

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

#### Modos de Recepción
Para recibir los datos desde Firebase se debe requerir directamente la dirección que contiene la información o el árbol de datos, recibiendo el dato contenido o los datos (valores y contenidos) respectivamente.


<div id='id17' />
      
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
		    If nRespuesta = 0 Then Goto fin
		    
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
  - Si el **Usuario** no tiene permisos recibirá un Valor = Disconnected.
  - Si la *dirección* de origen no existe recibirá un Valor = null.
  - Si todo va bien recibirá un Array con el contenido. 


<div id='id18' />

Se puede eliminar un valor y su contenido directamente o un árbol completo de datos usando el modo **DELETE**.

- El botón *BORRAR* del ejemplo adjunto en este repositorio, sería:

      Private Sub BorrarPath_Click()
          Dim Direccion As String
          Dim Mensaje As String
          Dim Respuesta As Variant

          Direccion = Caminofb.Text            ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

          Respuesta = FirebaseDB("DELETE", Direccion, Mensaje, TokenAutorizacion)
      End Sub
  

<div id='id19' />

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



<div id='id20' />

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


<div id='id21' />
   
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


<div id='id22' />
         
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


<div id='id23'  />

## Otras funciones

Hay funciones que trabajan de manera concreta como **DevolverValorEspecificoDeFirebase**:

		Function DevolverValorEspecificoDeFirebase(Direccion As String, Valor As String, Optional claveautorizacion As String = "") As String

Nos devuelve el contenido de un **Valor** especificado en el campo *Valor* en una *Direccion* de la *Base De Datos* sin tener que manejar JSON por parte del **Usuario**. Esta función se puede descartar y usar FirebaseDB en modo **GET**, pero la diferencié así por comodidad en algunos casos específicos.

<div id='id24'  />

Otra función de las mismas características que la anterior, pero que trabaja sobre los datos del **Usuario** y no de la Base de Datos es **DevolverValorAutorizacion**:

		Function DevolverValorAutorizacion(Valor As String, IDUsereMail As String, IDUserPassword As String) As String


Nos retorna un valor específico de la cadena devuelta por el servidor cuando se le consulta por los datos del **Usuario**. Se usa para adquirir el IdToken de **Usuario** para poder trabajar con la base de datos.


<div id='id25'  />

## FirebasePC
Función que permite trabajar con los archivos JSON descargados de la Base de Datos. En este contexto no se ha usado, ya que está más centrado en el proceso de datos de **Realtime Database**. :point_right: **Esta función no está desarrollada completamente aún** :point_left:.

		Function FirebasePC(Mode As String, Direccion As String, Mensaje As String, Optional SoloContenidoIndice As Boolean = False) As Variant

La estructura es similar a la ***Función FirebaseDB***, pero los datos se maneja off-line.
Sólo funciona con un modo: **GET** que nos permite abrir un archivo JSON local y chequearlo para poder extraer información.


<div id='id26'  />

## DevolverValorEspecificoDeJSONLocal
Trabaja en conjunción cno FirebasePC y permite extraer el contenido de un valor indicado.

		Function DevolverValorEspecificoDeJSONLocal(DirectorioYArchivo As String, Valor As String) As String


<div id='id27'  />

## Funciones de Registro
Las funciones de registro son funciones que procesan ciertos datos y los envía al servidor sin esperar respuestas, suele usarme para generara 'Logs' de Uso o Registrar Errores que puedan aparecer en el programa. Estas funciones son:

<div id='id28'  />

  - **GenerarJSONError**: Registra los errores que puedan aparecer en la aplicación en modo ***POST***. Si por cuaquier motivo aparece un error de conexión, la cadena JSON que contiene el error se guarda en una carpeta especificada. Se puede crear una Función que intente enviar el contenido de archivos generados en otro momento (no incluido en este repositorio).

		Function GenerarJSONError(NumeroError, descripcionerror, Mensaje) As Variant

Un ejemplo de uso: Cuando se intenta abrir un documento y este no existe se reedirigen los datos del error usando 'on error' y se capturan los valores del error a través de err.number y err.description. La variable *Mensaje* suelo dejarla para textos personalizados.

<div id='id29' 

  - **RegistrarUso**: Registra una cadena JSON con datos específicos en modo ***POST***. De esta manera se puede llevar un control del uso de una aplicación.

		Function RegistrarUso(Optional qModulo As String = "", Optional ElToken As String = "")

Un ejemplo de uso: Lo pongo en 'Thisworkbook' para que se registre cada vez que se abra el archivo y en algunos Formularios para registar cuales se usan más. La variable *qModulo* contendría el nombre del Formulario y *ELToken* el IdToken por si donde se va a registrar está protegido con *Reglas de Seguridad*.

<div id='id30' />

- **AlmacenarJSON** es una pequeña Función que recoge una cadena de texto la guarda en el computador. 

 		Function AlmacenarJSON(RutaYArchivo As String, Contenido As String)

Para este contexto se usa para guardar cadenas con estructura JSON en una ruta indicada.


<div id='id31' />

## Acciones para USUARIO

Para trabajar con usuarios usaremos la función **AccionConUsuario**:

		Function AccionConUsuario(Accion As String, IDUsereMail As String, IDUserPassword As String, _
		                            Optional IDTokenUser As String = "", Optional IDNameUser As String = "", _
		                            Optional IDURLFoto As String = "") As Variant

  - **Accion**: Se indicará que acción se llevará a cabo por el usuario. Para más información, ver la lista de acciones más abajo (Valor Obligatorio).
  - **IDUsereMail**: Se indicará el correo electronico del **Usuario** (Valor Obligatorio).
  - **IDUserPassword**: Se indicará el password de la sesión de **Usuario** (Valor Obligatorio).
  - **IDTokenUser**: Se suministrará el IdToken cuando se realice alguna acción sobre la información del **Usuario** (Valor opcional).
  - **IDNameUser**: Se indicará el Nombre de **Usuario** que será mostrado por el servidor (Valor opcional).
  - **IDURLFoto**: Se indicará la dirección Web de la foto de perfil del **Usuario** (Valor opcional).

Para indicar qué acción se llevará a cabo, se indicará una palabra clave específica. A continuación, vamos a ver que parámetros serán necesarios para llevar a cabo una acción sobre un **Usuario**. Para más información, ver el código de ejemplo integrado en el archivo excel del repositorio.

<div id='id32' />

## Crear nuevo Usuario
Creará un nuevo usuario a través del *Correo Electrónico* y una *clave de acceso*. Se realizará con el parámetro en *Accion* **NEW**.

  		AccionConUsuario("NEW", eMail, Password)


<div id='id33' />

## Crear Usuario Anonimo	
Permite crear un Usuario Anonimo con las mismas características de un Usuario Registrado. El IdToken generado cadurá pasada una hora. Se puede actualizar de ***ANONIMUS*** a ***Usuario Registrado*** usando el **IdToken** generado para el primero y actualizando los datos con la **Accion UPDATE**. Se realizará con el parámetro en *Accion* **ANONIMUS**.

  		AccionConUsuario("ANONIMUS", "", "")


<div id='id34' />

## Requerir información de un Usuario
Recupera los datos del Usuario cuyo IdToken esté activo. Se realizará con el parámetro en *Accion* **INFO**.

  		AccionConUsuario("INFO", eMail, Password)


<div id='id35' />

## Actualizar información de un Usuario
Actualiza la información de un Usuario excepto la dirección de Correo Electrónico y el Password. Esta acción también refresca el IdToken, pero tiene que ser antes de que caduque (tienen una vida de 3600 segundos). Se realizará con el parámetro en *Accion* **UPDATE**.

  		AccionConUsuario("UPDATE", eMail, Password, IdToken, Nombre, URLfoto)


<div id='id36' />

## Activa el IdToken de un Usuario
Recupera un IdToken actualizado de un **Usuario Registrado**. Esto es necesario para realizar tareas en la base de datos, sin este IdToken, no podrán realizarse acciones sobre los datos de la misma si las reglas de seguridad así lo espacifícan. Se realizará con el parámetro en *Accion* **AUTH**.

  		AccionConUsuario("AUTH", eMail, Password)


<div id='id37' />

## Borra un Usuario
Borra al **Usuario** registrado cuyas credenciales son indicadas. Se realizará con el parámetro en *Accion* **REMOVE**.

		AccionConUsuario("REMOVE", eMail, Password)


<div id='id38' />

## Petición de nuevo password de Usuario
Si un **Usuario** registrado ha perdido su **Password**, es posible enviarse un eMail con la acción de recuperación del mismo. Se realizará con el parámetro en *Accion* **RESETPASSWORD**.

		AccionConUsuario("RESETPASSWORD", eMail, "")


<div id='id39' />

## Cambiar eMail

Permite a un **Usuario** registrado cambiar el correo electrónico de sus credenciales. Se realizará con el parámetro en *Accion* **CHANGEMAIL**.

		AccionConUsuario("CHANGEMAIL", eMail, "", IDTokenUser))


<div id='id40' />

## Cambiar Password

Permite a un **Usuario** registrado cambiar el password de sus credenciales. Se realizará con el parámetro en *Accion* **CHANGEPASSWORD**.

  - **CHANGEPASSWORD**: Permite cambiar el password de un **Usuario Registrado**.

		AccionConUsuario("CHANGEPASSWORD", "", Password, IDTokenUser)


<div id='id41' />

## Qué responde el Servidor

Cuando realizamos una petición con la función **AccionConUsuario** recibiremos de vuelta una matriz de datos. Estos datos pueden ser extraidos y desplagados usando un bucle ***For...Next simple***. Para ello es muy importante saber qué palabras clave contendrá dicha matriz.
Algunas palabras claves han sido modificadas (dentro del módulo Firebase) para obtener unos resultados más funcionales. En la tabla siguiente se muestran las palabras clave que obtendremos en la matriz, marcando con un :envelope: los no modificados y con un :love_letter: los que se han modificado.

| Nombre | Descripción | Acciones |
| :--- | --- | --- |
| :envelope: kind | Devuelve el tipo de operación solicitada al **Servidor** | _NEW > _ANONIMUS > REMOVE > AUTH > INFO > CHANGEPASSWORD |
| :envelope: idToken | Devuelve el código de Autorización del **Usuario** actualizado | _NEW > _ANONIMUS > AUTH > CHANGEPASSWORD |
| :envelope: email | Devuelve el eMail de **Usuario** | _NEW > AUTH > INFO > CHANGEPASSWORD |
| :envelope: refreshToken | Devuelve el código de Autorización de refresco del **Usuario** | _NEW > _ANONIMUS > AUTH > CHANGEPASSWORD |
| :envelope: expiresIn | Devuelve el tiempo en segundo en los que el Id Token caducará (*por defecto 3600 s.*) | _NEW > _ANONIMUS > AUTH > CHANGEPASSWORD |
| :envelope: localID | Devuelve el Identificador de **Usuario** | _NEW > _ANONIMUS > AUTH > INFO > CHANGEPASSWORD |
| :envelope: passwordHash | Devuelve la version del HASH del Password | INFO > CHANGEPASSWORD |
| :envelope: displayName | Devuelve el nombre de **Usuario** que será mostrado | AUTH > CHANGEPASSWORD |
| :envelope: registered | Devuelve un valor buleano indicando si el correo electrónico es para una cuenta existente | AUTH |
| :envelope: emailVerified | Devuelve un valor buleano indicando si el correo electrónico de inicio de sesión está verificado | INFO > CHANGEPASSWORD |
| :envelope: passwordUpdateAt | Devuelve el *timestamp* en milisegundos en la que se cambió por última vez la contraseña de la cuenta | INFO  |
| :envelope: validSince | Devuelve la marca de tiempo, en segundos, que marca un límite, antes del cual el token de ID de Firebase se considera revocado | INFO |
| :envelope: lastLoginAt | Devuelve el *timestap* en milisegundos en la que la cuenta inició sesión por última | INFO |
| :envelope: createdAt | Devuelve el *timestap* en milisegundos en la que la cuenta fue creada | INFO |
| :envelope: lastRefreshAt | Devuelve el *timestap* en milisegundos del último refresco de información del **Usuario** | INFO |
| :envelope: photoUrl | Devuelve la dirección de la imagen de **Usuario** |
| errors |
| :envelope: error | Devuelve un código de error si la petición ha sido rechazada por el Servidor. Este dato no aparece si no hay error| En todas las Acciones |
| :envelope: menssage | Devuelve un mensaje de error si la petición ha sido rechazada por el Servidor. Ver [lista de errores comunes](#id44). Este dato no aparece si no hay error | En todas las Acciones |
| :envelope: reason | Devuelve una cadena de error si la petición ha sido rechazada por el Servidor. Este dato no aparece si no hay error| En todas las Acciones |
| providerUserInfo | Devuelve la lista de todos los objetos de proveedor vinculados que contienen "providerId" y "federatedId". | CHANGEPASSWORD |
| :envelope: puIproviderId | Devuelve el ID de proveedor vinculado (por ejemplo, "google.com" para el proveedor de Google). Sino va vinculado a ningún servidio mostrará la palabra clave *password* | INFO |
| :love_letter: puIdisplayName | Devuelve el nombre de **Usuario** que será mostrado |  |
| :love_letter: puIphotoUrl | Devuelve la dirección de la imagen de **Usuario** | INFO |
| :love_letter: puIfederateId | Devuelve el identificador ID único de la cuenta IdP (proveedor de Identidad) | INFO |
| :love_letter: puIrawId | Devuelve el Identificador de Credenciales | INFO |


Con esta información, podemos extraer los datos de la siguiente manera (ejemplo localizado en el archivo EXCEL adjunto a este repositorio):


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


<div id='id42' />

## ComprobarConexion

Con esta función nos conectamos a la base de datos y requerimos un valor específico. Si está conectado o no.

		Function ComprobarConexion() As Boolean

En el apartado [Reglas de Seguridad](#id5) se explica como se deja en Modo sólo lectura un directorio que contiene un valor que es el que revisa esta función.


<div id='id43' />

## ERRORES
Cuando enviamos información al servidor para realizar una petición, en respuesta se recibe un código, este código puede significar que la petición fue aceptada o presentó un error:
  - **200**: Petición aceptada.
  - **400**: El servidor no ha podido procesar la petición porque hay un error.


<div id='id44' />

### Lista de Errores Comunes

 | Error | Descripción | Acción |
 | :---: | --- | --- |
 | **EMAIL_NOT_FOUND** | El eMail introducido por el usuario no está registrado. | Comprueba la sintaxis de los datos introducidos o ponte en contacto con el **Administrador** para registrar tu **Usuario**. |
 | **INVALID_PASSWORD** | El password correspondiente al eMail introducido no es correcto. | Compruebe la sintaxis de los datos introducidos o realice una petición de recuperación al Administrador. |
 | **MISSING_PASSWORD** | El password que está intentando cambiar está vacío. | Revise el campo que contiene el password, o está vacío o el nombre es incorrecto. |
 | **INVALID_ID_TOKEN** | Está intentando realizar una acción sin estar identificado o con un IdToken diferente al **Usuario** indicado. | Inicie sesión con sus credenciales para realizar la acción deseada. |
 | **CREDENTIAL_TOO_OLD_LOGIN_AGAIN** | Ha intentado realizar alguna acción en el servidor con una credencial caducada. | Vuelva a conextarse para actualizar el IdToken de **Usuario**. |
 | EMAIL_EXIST | Intenta crear un **Usuario** con una eMail que ya existe. | Compruebe las credenciales de los usuario existentes o pruebe con otra dirección de correo electrónico. |
 | TOO_MANY_ATTEMPTS_TRY_LATER | Se ha superado el numero de intentos de conexión con un **Usuario** y se ha dehabilitado temporalmente la cuenta. | Póngase encontacto con el Administrador si es necesario o requiera un reseteo del Password. |
 
