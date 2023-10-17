# Proyecto Excels AFIP

### Utilización de R
   Durante mi trayectoria académica estudiando Licenciatura en Estadística, mi formación principal, relacionada con la programación, se centró en el lenguaje de programación R. Mi experiencia se enfocó en el manejo de bases de datos, análisis estadísticos y manipulación de archivos de Excel utilizando dicho lenguaje. Por lo tanto, inicialmente consideré que utilizar R para escribir este programa sería la elección más adecuada y eficiente debido a mi familiaridad y experiencia con este lenguaje.
   
   Durante la primera semana de desarrollo, logré un avance significativo al crear un programa capaz de procesar múltiples archivos de Excel que contenían facturas emitidas y recibidas de un mismo cliente. Este programa generaba un nuevo archivo de Excel con toda la información esencial cuidadosamente organizada: celdas y columnas correctamente formateadas, ajustes en las fuentes y tamaños de letra, modificaciones en los importes, y proporcionaba una forma mucho más cómoda y eficiente de visualizar los datos:
   
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/00f16c2d-9adc-4018-bb17-77bf0baae165)
   El usuario colocaba en una carpeta predefinida todos los archivos de Excel que contengan la información de interés del cliente

   Una vez ejecutado el programa, se crea un archivo de Excel con 4 Hojas: Ventas (Facturas Emitidas), Ventas Escalera (Hoja para visualizar más fácilmente las Ventas), Compras (Facturas Recibidas), Compras Escalera(Hoja para visualizar más fácilmente las Compras). Para mantener la privacidad del cliente, oculté información confidencial.
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/c5ff2b04-9549-4148-a536-2d86ea3f6076)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/ea8efce9-04b6-4c22-91b4-2a393be91749)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/c2dbfa19-921d-4387-8314-261b5d45caa4)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/e73b7faa-4ebc-4a49-ae56-d1ca2ea83ab0)

   Cuando presenté el programa a mi papa, quedó considerablemente convencido por su utilidad. Sin embargo, surgió un desafío: en la mayoría de los casos, su tarea no implicaba la creación de un nuevo archivo Excel, sino la actualización de archivos ya existentes con nueva información. Por lo tanto, se volvió esencial adaptar el programa para que pudiera cumplir con esta función.
   
   Además, aprendí en el proceso que R es un lenguaje de programación que se ejecuta de manera secuencial. Esto me planteó un obstáculo: no había una forma sencilla de crear un acceso directo para que los usuarios pudieran simplemente hacer doble clic en un ícono y abrir el programa, acceder a un menú de opciones en la consola y realizar las operaciones que desearan. Ante esta limitación, me dí cuenta de que R no era la opción viable para continuar con el proyecto. Fue en ese momento cuando decidí migrar todo el programa a Python, un lenguaje de programación que ofrecía la flexibilidad necesaria para lograr este tipo de interacción, basándome en mi experiencia previa con Python.

### Utilización de Python
   El pasar todo el código a Python fue un desafio, en especial debido a que gran aprte de las librerias que iba a terminar utilizando nunca las habia usado ni sabia como funcionaban, tampoco habia hecho ningun proceso de data cleaning intenso en el pasado, por lo que el proceso de pasar de R a Python fue una experiencia completamente nueva y desafiante.
   
   A pesar de eso, logré pasar el programa con exito, asi que ahora tocaba la parte de hacer que el programa actualice excels ya preexistentes.
   
   La mayor dificultad se presentó con los Excels ya existentes, que eran aproximadamente 90, los cuales como fueron creados, actualizados y modificados manualmente, cada uno tenia un formato distinto, a veces incluso el formato variaba dentro del mismo excel, a veces habia multiples hojas dentro de un mismo archivo con la misma información o con información que deberia estar toda junta en una hoja, como tambien informacion que deberia estar en hojas separadas, o que estaba en una columna pero deberia estar en otra, cadenas de caracteres donde deberian ir numeros, etc. A pesar de que logre programar ciertas funciones que ayudaron en el proceso de formatear todos los excels para que sigan todos una misma plantilla y estructura, gran parte del trabajo tuve que hacerlo manualmente porque no habia forma de automatizarlo con un programa.
   
   Y esto tiene una explicación, los excels no fueron ni van a ser utilzados como bases de datos, sino como archivos de excel, donde se van a cambiar celdas, agregar comentarios, borrar filas, crear columnas, etc. Por lo que todas estas cualidades iban a dificultad la programacion de un código que iba a ser capaz de actualizar los excels dado de que el programa iba a tener que lograr agregar la información sin modificar en ningun aspecto ninguna otra celda del excel mas que las que vaya a utilizar para agregar la nueva información.
   
   Después de varias semanas, logré hacer un programa que lograba lo que se quería desde un principio:

- El usuario puede descargar la cantidad de archivos de excel, de Facturas Emitidas como tambien de Facturas Recibidas, que el quiera, de la cantidad de clientes que desee, y al ejecutar el programa este no solo va a filtrar los excels que le perteneces a cada cliente, los formatea y aplica cualquier modificación necesaria, sino que tambien encuentra el excel del cliente al cual se deseaba actualizar y actualiza dicho excel con la nueva información. Este procedimiento lo repite con todos los clientes a los que se les haya descargado nueva información para actualizar sus respectivos excels. (Además cumple la función que inicialmente programe: crear un excel para un nuevo cliente)

  Para mostrar como funciona el programa, a continuación voy a mostrar el output de ejecutar el programa habiendo descargado archivos de excel de dos clientes totalmente distintos, uno al cual quiero que actualice su excel ya preexistente, y otro que no existe su excel por lo cual voy a crear un neuvo excel con la informacion que acabo de descargar.

![image](https://github.com/marcosziadi/excels_afip/assets/82457357/c14b26e2-0e0b-40ed-89be-2bf034bfe7f1)
  Se abre el menú de opciones, elijo primero actualizar al cliente con excel ya preexistente

![image](https://github.com/marcosziadi/excels_afip/assets/82457357/b929ff17-c268-4a21-97ac-07fb3d4489a6)
   El programa me indica los clientes que va a actualizar (Basicamente identifica de que clientes descargue archivos de excel con información nueva para actualizar)

![image](https://github.com/marcosziadi/excels_afip/assets/82457357/5fb52dac-153b-4a18-9e81-5e65b1622292)
   El programa me confirma los clientes que actualizó, y eso es algo que podemos confirmar manualmente entrando al excel de dicha persona:
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/17ef06b3-bd1e-43ee-8c82-5b9e8955042f)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/ca382ae7-5384-4ca9-bfa0-3b84a7b75694)

   Ahora, creemos al nuevo cliente:
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/205073eb-882b-48c3-9c95-46bbb77aac5d)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/ca46c349-8e76-4631-938f-395754c4383e)
   El excel fue creado, asi que ahora vamos a ver si realmente se creo y como se creo. (En este caso, el cliente solo emite Facturas Recibidas, asi que la hoja de "VENTAS NUEVO" esta completamente vacia):
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/665768c4-2c7b-4457-b82f-d1c3ad57a546)
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/ae6e596e-79da-49e0-b150-5694fc54cb1d)

- Lo bueno de haber logrado esa visualización en forma de escalera de la información con comandos de excel es que si mi papa tiene que hacer cambios manualmente a los clientes o agregar/sacar algo, dicha información se va a actualizar inmediatamente en la "zona verde" de visualización, por lo que es bastante práctico.

  A continuación, dejo un ejemplo de como son los excels descargados directamente desde la afip:
![image](https://github.com/marcosziadi/excels_afip/assets/82457357/26113a96-0f28-41a8-b29d-d04e52dded16)


  

