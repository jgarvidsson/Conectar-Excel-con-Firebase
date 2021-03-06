En el hilo de **Ejemplo** podéis encontrar el archivo de 'ejemplo' que contine las funciones indicadas en [Readme.md](README.md). Como comento, es un ejemplo de como usar las Funciones, el Usuario podrá modificar o usar estas características para programar de acuerdo a sus propósitos. 

<p align="center">
**RECUERDA QUE PARA COMPRENDER MEJOR LAS ACCIONES DE ESTE ARCHIVO, DEBES HABER LEÍDO ANTES README.MD**
</p>

## Que vamos a ver en este texto...

1. [Cómo usar el archivo de Ejemplo](#id1)
2. [Abrir el archivo](#id2)
   - [Probar Edición de Datos en RealTime Database](#id3)
     - [Conectar con la Base De Datos](#id4)
     - [Enviar registros a las base de datos](#id5)
     - [Manejar registros de la base de datos](#id6)
   - [Probar Administración de Usuario](#id7)
     - [Nuevos Usuarios](#id8)
     - [Datos de Usuarios](#id9)
     - [Acciones de Usuarios](#id10)


<div id='id1' />

# Como usar el archivo de Ejemplo
Es archivo [**Conectar con Firebase.xlsm**](https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/ab9879fed13a3c7f719b4607dabd2b4468ef8b76/Conectar%20Con%20Firebase.xlsm?raw=true) muestra en dos ventanas el funcionamiento de las Funciones contenidas en el Modulo Firebase.

Los ejemplos mostrados permiten manejar una base de datos previamente creada por el Usuario y configurada como se explica en Readme.md.
Por otro lado podrá crear o eliminar Usuario del registro.

<div id='id2' />

# Abrir el archivo
Cuando abres el archivo podrás visualizar las siguiente pantalla.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/01_PantallaInicio.png" />
</p>

Si la conexión con la base de datos ha sido correcta, en la parte inferior de la ventana aparecerá el mensaje 'Conexion with Server OK'. Para que esto ocurra las credenciales de la base de datos deben estar correctamente configuradas y el parámetro de conexión indicada en la base de datos. Esta información está descrita en el archivo README.MD de este repositorio.

<div id='id3' />

## Probar Edición de Datos en RealTime Database
Abre el formulario (ventana) que permite conectar con la Base de Datos para aquirir el IdToken del Usuario y poder realizar varias operaciones.

<div id='id4' />

### Conectar con la Base De Datos
Para conectar con la base de datos debes tener al Usuario regisrado.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/02_PantallaConectar.png" />
</p>

Introducimos el correo electrónico y el password del Usuario y pulsamos 'Conectar'. Si no aparece ningún mensaje de error, en la parte superior de la ventana aparecerá indicado que el 'ID Token para el usuario' está activo.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/03_PantallaConectada.png" />
</p>

<div id='id5' />

### Enviar registros a las base de datos

En la pestaña 'Enviar' podremos ver las tres opciones de envío que nos da Firebase.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/04_PantallaEnviar.png" />
</p>

  - Con **Post**, enviaremos la información agregándola a la base de datos en modo 'posteo' como si fuera una red social. Se generará un directorio con base de tiempo que ordenará las subidas en el mismo orden en el que se vayan haciendo.

  - Con **Patch** actializaremos las información existente o crearemos un registro nuevo. No se borrarán los datos no actualizados.
  
  - Con **Put** enviaremos la información a la dirección indicada. Si se reenvía la información, el contenido anterior será borrado y lo sustituirá el nuevo. Lo contrario que pasaría con ***Patch***.

<div id='id6' />

### Manejar registros de la base de datos

En la pestaña 'Recibir/Editar/Descargar/Mover' podremos manerjar la información alamacenada en la base de datos.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/05_PantallaAcciones.png" />
</p>

  - **Recibir** toma los datos de la dirección indicada y los muestra en pantalla.
  - **Borrar** borra los datos de la dirección indicada.
  - **Descargar** descarga en el computador la informacion contenida en la dirección indicada en formato JSON.
  - **BackUp** realiza una copia de la información contenida en la dirección indicada a una segunda dirección en modo ***POST***. No borra los datos de origen.
  - **Mover** mueve los datos contenidos en una dirección a un destino. Borra la información de la dirección de origen.
  - **Copiar** realiza una copia de la información contenida en la dirección indicada a una segunda dirección en modo ***PUT***. No borra los datos de origen.

Estos son sólo unos ejemplos de las opciones que pueden desarrollarse.


<div id='id7' />

## Probar Administración de Usuario
Abre el formulario (ventana) que permite crear, borrar y modificar datos de Usuario que podrán manejar la información de la base de datos.


<div id='id8' />

### Nuevos Usuarios
En la pestaña **Registro de Usuarios** podremos registrar o borrar a Usuarios.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/10_Usuario_Registro.png" />
</p>

Se pueden **registrar** nuevos usuarios con ***credenciales*** o ***anónimos***, pero no se podrán borrar si no tienes activo un IdToken de **Usuario**.


<div id='id9' />

### Datos de Usuarios
En la pestaña **Datos de Usuarios** podremos ver el perfil del usuario activo o actualizar sus datos (sólo el nombre mostrado y la URL de la imagen). Para implementar otros datos hace falta una tarea específica sobre la base de datos. **Lo veremos en un proyecto posterior**.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/12_Usuarios_Datos.png" />
</p>

  - **Ver Perfil** permite revisar los datos almacenados en el servidor sobre el **Usuario**. No son datos personales.
  - **Actualizar Datos** permite actualizar el nombre y/o la URL donde se encuentra almacenada la *Imagen del Usuario*.
  - **ID Token** realiza una petición al servidor para que envíe al **Usuario** un IdToken actualizado para poder realizar varias tareas: *Borrar Usuario*, *Ver Perfil* o *Actualizar Datos*.

<div id='id10' />

### Acciones de Usuarios
En la pestaña **Acciones de Usuarios** podremos resetear el Password de Usuario o modificar las credenciales.

<p align="center">
  <img src="https://github.com/jgarvidsson/Conectar-Excel-con-Firebase/blob/img/13_Usuario_Acciones.png" />
</p>

  - **Reset Password** enviará al correo indicado un acceso para resetear el Password por si ha sido olvidado.
  - **Cambiar eMail** modifica el eMail del usuario.
  - **Cambiar Pass** modifica el Password del usuario.

Para realizar estas tareas el IdToken de Usuario tiene que estar activo.
