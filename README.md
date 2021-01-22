# Conectar Excel con Firebase

## Antes de Empezar
Antes de empezar es necesario remarcar que para conectar **Firebase** con **EXCEL** se requiere tener creada y configurada una base de datos.

El ejemplo adjunto en este repositorio, crea una carpeta en "Mis Documentos" (Documents) con el nombre de **fbExcel** donde se guardará algún datos que necesite ser descargado (foto de perfil o archivo JSON). Este procedimiento lo uso habitualmente para cuando finalice con las pruebas, poder borrar todos los archivos generados de manera más sencilla. Para cambiar el nombre de esta carpeta, abra el modulo **Herramientas** y cambie el contenido de la variable **NombreCarpetaTrabajo** por el que desee.

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
          Dim i As Single

          Direccion = Caminofb.Text       ' Se transfiere el dato del textbox con la dirección seleccionada a la variable dirección.

          Respuesta = FirebaseDB("GET", Direccion, Mensaje, TokenAutorizacion)
          nRespuesta = arrayLength(Respuesta)
          If nRespuesta = 0 Then Exit Sub

          For i = 0 To nRespuesta - 1
              If Respuesta(i, 0) = Valor2 Then Contenido2.Text = Respuesta(i, 1)
              If Respuesta(i, 0) = "TSL" Then TimeLocal.Text = Respuesta(i, 1)
              If Respuesta(i, 0) = "TSS" Then TimeServer.Text = Respuesta(i, 1)
          Next i
      End Sub

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

