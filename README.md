# PRE-REQUISITOS
- Tener instalada la ultima version de python
- Instalar pip
- Descargar y guardar el archivo credentials.json en el folder del script
- El archivo original con los alumnos tiene que tener el nombre 'los-3-listados-siu.xls'. Puede modificarse facil en el script, es la variable 'input_file'. Ademas, el archivo tiene que seguir el formato que ejemplico en ejemplo.xls.

# INSTALACION E INICIALIZACION ENTORNO VIRTUAL (venv)
python -m venv venv

inicio entorno virtual

// Con lo siguiente incializamos el entorno virtual

### Linux
source venv/bin/activa

### Windows
venv\Scripts\activate


# INSTALACION DE DEPENDENCIAS
pip install -r requeriments.txt

# EJECUCION
NOTA: Esto siempre se ejecuta desde la consola donde se inicializo en entorno virtual (venv))

### Linux
python3 ./main.py

### Windows
py .\main.py

NOTA: En la primer ejecucion va a pedir que apretes un URL que te redirecciona a google para que inicies sesion y des permisos. Usar siempre la cuenta institucional, el script lo que hace es solicitar a la API esos datos asociados a tu cuenta. La info obtenida es solo de los mails y solo se usa para generar el listado. No se obtiene nada mas ni se va a filtrar la informacion utilizada (mentira).

# ACLARACIONES
- La planilla resultante se guarda en el mismo folder donde se encuentra el Scripts con el nombre 'listado_con_mails.xlsx'.
- Creo que normalice todo de manera correcta para obtener todos los alumnos. En caso de que se de un caso particular no contemplado, el script pone el legajo y nombre poniendo un '-' En la casilla mail. Eso deberia buscarse manualmente.
- Detecte que hay alumnos que se inscribieron con dos legajos distintos. Eso genera que para los diferentes legajos se tiene un mail unico. Pero el script asocia para un mismo legajo los dos mails. Por eso puede haber mas filas que alumnos inscriptos.
- El formato del resultado se puede cambiar de manera sencilla modificando la parte final del script. Ahora genera archivos .xlsx.
