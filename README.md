# Optimizador_horario
Este repositorio contiene un programa el cual utiliza Python y Programacion lineal para lograr un horario mas comodo. Es importatne considerar que este codigo fue creado con las bases de python, pero me parecion pertinente subirlo para el que quiera utilizarlo/modificarlo

# Utilizacion
Ahora bien, para utilizar este programa se debe ejecutar el archivo **Optimizador de Horario V8.py**. Hay dos consideracion importantes:
+ Para que el programa pueda hallar un OFG que encaje de forma optima se debe tener los textos **ramos de ...** descargados y en la misma carpeta. En caso de no tenerlos se puede utilizar para crear horarios con los ramos entragos, pero sin buscar un ofg.
+ Si no se tiene los modulos pertinentes (los cuales seran mencionados en el siguiente inciso) el programa puede resultar en error al ejecutarse por primera vez. Sin embargo, dirante esa ejecucion se insatalara los modulos necesarios, por solo se debe ejecutar nuevamente el programa.

# Resultado de la ejecucion
Este Programa devuelve dos cosas principalmente:
+ Crea una tabla en la terminal, la cual contiene la informacion de los ramos seleccionados, tal como el NRC, nombre, fecha, etc. Tambien informa si el resultado fue _Optimal_ (Que se creo correctamente el horario) o _Infeseable_ (No existe un horario posible).
+ Crea un archivo excel con el horario creado. Este tiene un pequeño inconveniente con los topes de horario. Este problema es que, en un modulo con tope, solo pondra una de las clases.

# Modulos utilizados

Los modulos utilizados en este archivo:

1. ```requests```
2. ```bs4```:```BeatifulSoup```
3. ```pulp```:```*```(Ahora se que esto no se hace, pero prefiero  no modificar el codigo)
4. ```subprocess```
5. ```tabulate```
6. ```pandas```
7. ```openpyxl```: ```Workbook, PatternFill```