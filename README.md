![logo](https://web.sistemasgyg.com.co/sistemasgyg.com.co/wp-content/uploads/Grupo-366-217x150.png)

# Ser_Excel 
Ser_Excel es una aplicacion creada en python que genera un excel a partir de un reporte txt y una plantilla xlsx, 




## Descarga

##### Clonar el repositorio 
```sh
  git clone url-del-ser-excel
```

##### Cambiar directorio 
```sh
  cd ser-excel
```

##### Si se va a realizar una modificacion se debe crear una rama con el numero del requerimiento o incidente
```sh
  git checkout -b R00000n
```

##### Si la rama ya esta creada solo debes cambiar de rama 
```sh
  git checkout R00000n
```





    
## Estructura de Carpetas

##### A continuacion se muestra la estructura de carpetas que conforma el ambiente de prueba y los fuentes del programa

```bash
ser_excel/
│
├── SIIF_IDEA/            # Ambiente de prueba
│     ├── AYUDAS/         # Carpeta ingresar reportes
│     ├── DOCUMENTOS/     # logs y reportes generados
│     ├── PLANTILLAS/     # plantillas xlsx
│     └── SIIFNET/        # ingresar .exe generado si se quiere probar
│
├── ser_excel/            # carpeta con fuentes del programa
│     ├── dist/           # carpeta con .exe generados
│     ├── venv/           # entorno virtual
│     └── img_barcode/    # imagenes de codigo de barras
│
└── README.md             # Documentación principal
```
## Uso de la aplicación local

1. **Ingresar a la carpeta con las fuentes del programa:**

    ```sh
    cd ser_excel/ser_excel/
    ```

2. **Crear el entorno virtual:**  
   Asegúrate de tener instalado Python en tu equipo.

    ```sh
    py -m venv venv
    ```

3. **Activar el entorno virtual de Python:**

    ```sh
    .\venv\Scripts\activate
    ```

4. **Instalar las dependencias que se encuentran en el archivo `requirements.txt`:**

    ```sh
    pip install -r requirements.txt
    ```

5. **Abrir el proyecto en un editor de código:**

    ```sh
    code .
    ```

6. **Archivo principal (`main.py`):**  
   El archivo principal donde se ejecuta el programa es `main.py`.

   Para ejecutar el programa localmente, realiza las siguientes modificaciones en el código:

   En el archivo `main.py`, comenta desde la línea 17 hasta la línea 42. A continuación se encuentran las líneas que debes comentar para pruebas locales:

    ```python
    # Obtener los parámetros de la línea de comandos
    # param = sys.argv
    # print(param)
    # params = param[1].split(separator)
    # if params.count == 1:
    #     params = param.split(" ")
    # path = os.path.abspath(__file__)  
    # print("real path: ", path)

    # name_report_txt = params[0][1:]  
    # utilita = params[0][0]  
    # number_session = params[1]
    # print("number_session: ", number_session)
    # print("name_report_txt: ", name_report_txt)
    # print("utilita: ", utilita)

    # Obtener la ruta completa del archivo ejecutable
    # ruta_exe = sys.executable

    # rout_aplication =  os.path.dirname(ruta_exe) #ruta SIIFNET
    # print("ruta SIIFNET: ", rout_aplication)

    # partes_ruta = os.path.normpath(rout_aplication).split(os.sep)

    # rout_environment = '\\'.join(partes_ruta[0:-2]) #ruta del ambiente "IDEA"
    # print("ruta ambiente: ", rout_environment)
    ```

   Luego, inserta las siguientes líneas después del código comentado:

    ```python
    number_session = "0000018"
    name_report_txt = '000007462'
    utilita = 'P'
    rout_environment = "C:\\Users\\javier.puentes\\ser_excel"
    rout_aplication =  "C:\\Users\\javier.puentes\\ser_excel\\ser_excel"
    ```

### Descripción de Variables

A continuación, se detalla cada una de las variables mencionadas anteriormente:

- **`number_session`**: Número que se utiliza para nombrar el reporte en formato `.xlsx` o `.pdf`.
- **`name_report_txt`**: Nombre del reporte. Este archivo debe estar ubicado en el ambiente de pruebas en la siguiente ruta: `SII_IDEA/AYUDAS/PAGINAS/reporte/000007462.txt`.
- **`utilita`**: Indica el formato del reporte:
  - `'P'`: Genera el reporte en formato PDF.
  - `'Y'`: Genera el reporte en formato XLSX.
- **`rout_environment`**: Ruta del entorno donde se encuentra el proyecto clonado. Corresponde a la raíz del proyecto.
- **`rout_aplication`**: Ruta donde se encuentra el archivo `main.py` de la aplicación.

> **Nota:** Para las variables relacionadas con las rutas (`rout_environment` y `rout_aplication`), solo es necesario reemplazar la parte que corresponde a la rutas antes de la carpeta clonada del ser_excel.

5. **Ejecutar pruebas locales:**

para realizar las pruebas se debe ejecutar el archivo main.py

```sh
python main.py
```
Dicho comando ejecutara el programa si se tiene algun problema puede revisar el log que se encuentra en el ambiente de prueba:
`ser_excel/SII_IDEA/DOCUMENTOS/LOGS/`

Si el programa se ejecuta correctamente encontraras los archivos generados en la siguiente ruta del ambiente de prueba:
`ser_excel/SII_IDEA/DOCUMENTOS/ARCHIVOS/`



## Generar .exe e Instalacion

Para crear el archivo ejecutable .exe ejecuta el siguiente comando

```bash
  pyinstaller --onefile main.py
```

en la carpeta `ser_excel\ser_excel\dist`, encontraras unarchivo creado main.exe, se debe renombrar con el siguiente nombre Ser_Excel.exe para poder ser instalado

para instalar el programa se debe pasar el archivo Ser_Excel.exe ateriormente creado a la carpeta SIIFNET del cliente, con esto ya estara instalado
