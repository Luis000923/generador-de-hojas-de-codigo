# Generador de Hojas de Código

Una herramienta desarrollada en Python para generar hojas de código automáticas a partir de datos de encuestas en Excel.

## � Capturas de Pantalla

### Interfaz Principal de la Aplicación
![Interfaz Principal](ejemplos/image.png)

*La aplicación cuenta con una interfaz gráfica intuitiva que permite cargar archivos Excel, previsualizar datos y generar hojas de código.*

## �📋 Descripción

Este proyecto es un generador automático de hojas de código que procesa datos de encuestas y genera documentación estructurada tanto en formato Excel como Word. La aplicación cuenta con una interfaz gráfica intuitiva que permite a los usuarios cargar archivos de Excel y generar automáticamente hojas de código con preguntas, alternativas y estadísticas de respuesta.

## ⚠️ Importante - Datos de Prueba

**NINGÚN CORREO ELECTRÓNICO HA SIDO FILTRADO NI COMPROMETIDO**

Los datos incluidos en este proyecto (archivos Excel) son únicamente **listas generadas aleatoriamente** con fines de prueba y demostración. No contienen información real de usuarios ni correos electrónicos reales.

## 🚀 Características

- **Carga automática de datos**: Importa archivos Excel con datos de encuestas
- **Generación automática de códigos**: Crea hojas de código estructuradas
- **Exportación a Excel**: Genera archivos .xlsx con formato profesional
- **Interfaz gráfica intuitiva**: Aplicación de escritorio fácil de usar con Tkinter
- **Detección automática de preguntas**: Identifica columnas relevantes automáticamente
- **Estadísticas integradas**: Calcula frecuencias y porcentajes automáticamente
- **Vista previa en tiempo real**: Muestra el contenido antes de generar
- **Campos personalizables**: Permite agregar información de institución, asignatura y docente
- **Control de muestra**: Opción para limitar el número de encuestados exportados

## 🛠️ Tecnologías Utilizadas

- **Python 3.x**
- **pandas 2.1.3**: Manipulación y análisis de datos
- **openpyxl 3.1.2**: Lectura y escritura de archivos Excel (.xlsx)
- **tkinter**: Interfaz gráfica de usuario (incluida con Python)
- **numpy 1.25.2**: Operaciones numéricas y estadísticas

### Dependencias del Sistema
- **Python 3.7+** requerido
- **tkinter** (incluido con Python estándar)
- **Excel compatible** para abrir archivos generados

## 📦 Instalación

### Método 1: Instalación automática con requirements.txt

1. Clona este repositorio:
```bash
git clone https://github.com/Luis000923/generador-de-hojas-de-codigo.git
cd generador-de-hojas-de-codigo
```

2. Instala todas las dependencias desde requirements.txt:
```bash
pip install -r requirements.txt
```

### Método 2: Instalación manual

Si prefieres instalar las dependencias manualmente:
```bash
pip install pandas==2.1.3 numpy==1.25.2 openpyxl==3.1.2
```

**Nota**: Tkinter viene incluido con la instalación estándar de Python, por lo que no necesita instalación adicional.

## 🎯 Uso

1. Ejecuta la aplicación:
```bash
python hoja_de_codigos.py
```

2. Usa la interfaz gráfica para:
   - **Cargar** tu archivo Excel con los datos de la encuesta
   - **Previsualizar** las preguntas detectadas automáticamente
   - **Configurar** opciones de generación (información institucional, límite de encuestados)
   - **Generar** la hoja de código en formato Excel

## 📖 Guía Visual Paso a Paso

### Paso 1: Cargar Archivo Excel
![Interfaz Principal](ejemplos/image.png)
- Haz clic en **"Seleccionar Excel"** para cargar tu archivo de encuesta
- La aplicación detectará automáticamente las preguntas y alternativas

### Paso 2: Verificar Vista Previa
- Revisa en la sección **"Vista Previa de Datos"** que las preguntas se hayan detectado correctamente
- Verifica que las alternativas de respuesta estén completas

### Paso 3: Configurar Opciones (Opcional)
- **Información de Institución**: Agrega nombre de institución, asignatura y docente
- **Límite de Encuestados**: Define cuántos encuestados exportar (0 = todos)

### Paso 4: Generar Archivo
- Haz clic en **"Generar Excel"** para crear la hoja de código
- Selecciona la ubicación donde guardar el archivo
- La aplicación creará un archivo Excel con formato profesional

### 📊 Funciones Adicionales
- **"Ver Estadísticas"**: Muestra un resumen detallado de frecuencias y porcentajes
- **"Limpiar"**: Reinicia la aplicación para procesar un nuevo archivo

## 📁 Estructura del Proyecto

```
├── hoja_de_codigos.py                                    # 🐍 Aplicación principal
├── requirements.txt                                      # 📋 Dependencias del proyecto
├── README.md                                            # 📖 Documentación
├── ejemplo.xlsx                                         # 📊 Archivo de ejemplo (datos de prueba)
├── Impacto de la alimentación encuesta (respuestas).xlsx # 📊 Ejemplo de encuesta real
└── ejemplos/                                            # 📁 Carpeta con recursos
    └── image.png                                        # 🖼️ Captura de pantalla de la app
```

### 📝 Descripción de Archivos
- **`hoja_de_codigos.py`**: Script principal con toda la lógica de la aplicación
- **`requirements.txt`**: Lista de todas las dependencias necesarias con versiones específicas
- **`ejemplo.xlsx`**: Archivo de muestra para probar la funcionalidad
- **`ejemplos/`**: Contiene recursos adicionales como imágenes y documentación visual

## 🔧 Funcionalidades Principales

### Clase `SurveyProcessor`
- `load_excel()`: Carga y procesa archivos Excel automáticamente
- `generate_coding_sheet()`: Genera la estructura de hoja de código
- `calculate_statistics()`: Calcula frecuencias y porcentajes
- `generate_excel_document()`: Crea documentos Excel con formato profesional

### Aplicación GUI (`SurveyApp`)
- **Interfaz gráfica intuitiva** con elementos organizados
- **Sistema de previsualización** para verificar datos antes de generar
- **Opciones de exportación personalizables**
- **Barra de estado informativa** con feedback en tiempo real
- **Campos opcionales** para información institucional

### 📊 Características del Procesamiento
- **Filtrado automático**: Excluye columnas administrativas (emails, timestamps, etc.)
- **Detección inteligente**: Identifica preguntas y alternativas automáticamente
- **Cálculos estadísticos**: Genera frecuencias absolutas y porcentajes
- **Formateo profesional**: Aplica estilos y bordes a las tablas Excel

## 📊 Formato de Entrada

El programa acepta archivos Excel (.xlsx) con las siguientes características:
- Primera fila como encabezados
- Columnas administrativas (correos, timestamps) son filtradas automáticamente
- Datos de respuestas en las columnas restantes

## 📄 Formato de Salida

### Archivo Excel (.xlsx)
- **Hoja de código estructurada** con preguntas y alternativas
- **Tablas profesionales** con bordes, colores y formato
- **Cálculos automáticos** de frecuencias y porcentajes
- **Múltiples hojas** organizadas por tipo de contenido:
  - Hoja de Código Principal
  - Estadísticas por pregunta
  - Resumen de datos

### 🎨 Características del Formato
- **Encabezados destacados** con formato de fuente y color
- **Bordes y líneas** para separar secciones
- **Colores alternados** en filas para mejor lectura
- **Ajuste automático** del ancho de columnas
- **Totales calculados** automáticamente

## ❓ Preguntas Frecuentes (FAQ)

### ¿Qué tipos de archivos Excel acepta?
- Archivos `.xlsx` (Excel 2007 o superior)
- Primera fila debe contener los encabezados de las preguntas
- Datos de respuestas en las filas siguientes

### ¿Cómo maneja la aplicación las columnas administrativas?
La aplicación **filtra automáticamente** columnas que contienen:
- Correos electrónicos (`email`, `correo`, `mail`, `@`)
- Marcas temporales (`timestamp`, `fecha`, `marca temporal`)
- Información de grado o nivel (`grado`)

### ¿Puedo procesar encuestas con respuestas múltiples?
Sí, la aplicación detecta automáticamente todas las alternativas únicas para cada pregunta y genera los códigos correspondientes.

### ¿Qué pasa si mi archivo Excel tiene errores?
La aplicación mostrará un mensaje de error específico. Verifica que:
- El archivo no esté corrupto
- Tenga al menos una columna con datos válidos
- Las respuestas no estén completamente vacías

## 🛠️ Solución de Problemas

### Error: "No se puede cargar el archivo"
- Verifica que el archivo esté en formato `.xlsx`
- Asegúrate de que el archivo no esté abierto en Excel
- Comprueba que tengas permisos de lectura en el archivo

### Error: "No se detectaron preguntas"
- Revisa que tu archivo tenga encabezados en la primera fila
- Verifica que las columnas no sean solo administrativas
- Asegúrate de que haya datos en las filas

### La aplicación se cierra inesperadamente
- Ejecuta desde la línea de comandos para ver errores específicos
- Verifica que todas las dependencias estén instaladas correctamente
- Comprueba la versión de Python (requiere 3.7+)

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📝 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.

## 👨‍💻 Autor

**Luis** - [Luis000923](https://github.com/Luis000923)

## 📧 Contacto

Si tienes preguntas o sugerencias, no dudes en crear un issue en este repositorio.

---

**Nota**: Este proyecto fue desarrollado con fines educativos y de demostración. Los datos incluidos son completamente ficticios y generados aleatoriamente.