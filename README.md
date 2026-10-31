# Organizar fotos y DPI de visitas domiciliarias (RPA)

Robot que automatiza el trabajo con las fotografías de WhatsApp de las visitas socioeconómicas del
Programa de Aporte Económico del Adulto Mayor (RGA): lee cada foto, identifica a la persona, renombra
los archivos con el formato que exige la macro y arma el informe de Word.
**nota:** se le envió un correo de invitación para que pueda visualizar nuestra documentacion en azure

> **Resultado:** el proceso manual de descargar, renombrar y dar formato a las fotos tomaba cerca de
> **10 horas**; con el robot toma alrededor de **10 minutos** (reducción de 98 %).
> **Todo se procesa en el equipo del usuario: ninguna imagen se envía a internet.**

## Cómo funciona

```
Carpeta de fotos ─► UiPath ─► Python ─► Word
                    (elige y    (lee, agrupa,   (macro que arma
                     valida)     renombra)       el informe)
```

| Paso | Qué hace | Quién |
|---|---|---|
| 1. Elegir la carpeta | La persona elige la carpeta con las fotos y los DPI | UiPath |
| 2. Validar la entrada | Comprueba carpeta, imágenes, programa instalado y que no haya una `_salida` vieja | UiPath |
| 3. Leer cada foto | OCR de la etiqueta de la cámara (6 lecturas con votación) | Python |
| 4. Agrupar y elegir el retrato | Junta las fotos de cada persona; la cara más grande es el retrato | Python |
| 5. Emparejar el DPI | Une frente y reverso por el CUI y su dígito de control | Python |
| 6. Renombrar y registrar | Copia con el nombre final; lo dudoso va a `Revision_manual`; genera Excel | Python |
| 7. Armar el informe | Inserta vivienda, retrato y DPI de cada persona en Word | Word (macro VBA) |

Nombres que genera el programa:

| Archivo | Contenido |
|---|---|
| `nombre_1.jpg` | Foto de medio cuerpo (retrato) |
| `nombre_2.jpg` … `nombre_4.jpg` | Fotos de la vivienda (de la quinta en adelante: `_5`, `_6`…) |
| `nombre_DPI1.jpg` / `nombre_DPI2.jpg` | DPI, frente y reverso |

Las personas se ordenan en el informe por la **hora de la visita** que trae la etiqueta de cada foto.
Los originales **nunca se modifican**: el programa trabaja con copias.

### ¿Usa inteligencia artificial?

Usa dos herramientas ya entrenadas que corren de forma local: **Tesseract** (OCR) y **cascadas de Haar de
OpenCV** (detección de rostros, solo para medir el tamaño de la cara). El agrupamiento, la asignación de
roles, el emparejamiento del DPI y la decisión de qué va a revisión manual son **reglas escritas en código**.
No usa IA generativa, ni servicios en la nube, y **no reconoce a las personas por su rostro**.

## Contenido del repositorio

```
README.md                                   Este archivo
requirements.txt                            Librerías de Python
Instalar_Python_Tesseract_y_librerias.bat   Instala Python, Tesseract y las librerías
Instalar_dependencias.bat                   Instala librerías y copia los archivos a C:\RpaFotos
Programa/
  organizar_fotos.py                        Programa de Python (código comentado)
  Programa_Documentado.Rmd                  El mismo programa explicado sección por sección
  Plantilla_fotos.docm                      Plantilla de Word con la macro
UiPath/
  OrganizarFotosyDPI/                       Proyecto de UiPath Studio (Main.xaml, project.json)
Documentacion/                              Diseño técnico, diagramas y código de las macros
```

## Requisitos (Windows 10 u 11)

- Python 3.12
- Tesseract OCR 5.x (`C:\Program Files\Tesseract-OCR`)
- Microsoft Word, con macros habilitadas
- UiPath Studio (edición Community) o UiPath Assistant para ejecutar el flujo
- Librerías de Python: `opencv-python` (**versión 4**, la 5 no trae los modelos de rostro que usa el programa),
  `numpy`, `pytesseract`, `openpyxl`

## Instalación

1. Ejecutar **`Instalar_Python_Tesseract_y_librerias.bat`** (doble clic). Instala Python 3.12, Tesseract y las
   librerías, y al final verifica que todo funcione. Puede pedir permiso de administrador. Necesita
   internet solo durante esta instalación.
2. Copiar `Plantilla_fotos.docm` dentro de `Programa/`.
3. Ejecutar **`Instalar_dependencias.bat`**: copia `organizar_fotos.py` y `Plantilla_fotos.docm` a `C:\RpaFotos`,
   que es donde los busca el robot.
4. Abrir `UiPath/OrganizarFotosyDPI` en UiPath Studio.

> Si el equipo no tiene `winget`, el instalador lo avisa; en ese caso instalar a mano
> [Python 3.12](https://www.python.org/downloads/) (marcando "Add python.exe to PATH") y
> [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki).

## Uso

**Con UiPath (recomendado):** ejecutar el proceso (desde UiPath Assistant, o `Main.xaml` en Studio), elegir
la carpeta con las fotos y esperar. Una ventana muestra el avance ("Imagen 12 de 35") y al terminar se abre la
carpeta `_salida` y el informe de Word. Si la ventana del selector de carpeta no aparece, buscarla con `Alt + Tab`.

**Solo el programa de Python:**

```
py -3.12 C:\RpaFotos\organizar_fotos.py "C:\ruta\carpeta_con_fotos"
```

Opciones: `--salida <carpeta>`, `--solo-analizar` (genera el Excel sin copiar archivos) y
`--sin-ventana` (sin ventana de avance).

Resultados en la carpeta de fotos:

```
_salida\                 Fotos y DPI con el nombre final (OK)
_salida\Revision_manual\ Lo que el programa no pudo asegurar
_salida\registro_fotos.xlsx   Una fila por imagen: nombre leído, confianza, rol, estado y motivo
_salida\resumen.json          Resumen que puede leer UiPath
Informe_fotos_HHMMSS.docm     Informe de Word generado
```

Si `_salida` ya contiene fotos de una corrida anterior, el programa se detiene a propósito: borrarla o
moverla antes de volver a ejecutar.

## Plantilla de Word

La macro vive en el `.docm`. Para recrearla: `Alt + F11`, insertar un módulo con
`Documentacion/Word_Modulo_ProcesarCarpeta.txt` y reemplazar `BotonAjustar_Click` en *ThisDocument* por el
bloque comentado al final de ese archivo. La macro ordena a las personas por la fecha de modificación del
archivo `_1.jpg`, que el programa fija con la hora de la visita.

## Cuándo algo va a revisión manual

El programa no adivina. Va a `Revision_manual` y queda anotado el motivo en el Excel cuando:
no se pudo abrir la imagen; hay etiqueta pero no se leyó el nombre; la imagen no tiene etiqueta y no es un
DPI; la persona no tiene un retrato claro; el nombre se leyó con baja confianza; o un DPI no se pudo asignar
con seguridad a una persona.

## Privacidad

Las fotos y los DPI son datos personales de adultos mayores. **No subir imágenes reales al repositorio**
(el `.gitignore` bloquea jpg y png). Usar datos ficticios o con los datos cubiertos en demostraciones y videos.

## Problemas conocidos

- Si Windows bloquea numpy o cv2 con "Control de aplicaciones", revisar *Seguridad de Windows > Control de
  aplicaciones y navegador > Control inteligente de aplicaciones* (desactivarlo es irreversible).
- Las fotos sin etiqueta de la cámara y los reversos de DPI ilegibles no se pueden asignar solos.
- Si la lectura de la nota deja una línea en blanco entre el número y el nombre, ese intento pierde el nombre;
  la votación entre seis lecturas lo compensa casi siempre.
- Si el selector de carpeta parece no abrir, puede estar detrás de otra ventana: `Alt + Tab`.

## Equipo

*(Completar: nombres de los integrantes, curso y docente.)*
