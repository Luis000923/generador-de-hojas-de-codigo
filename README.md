# Generador de Hojas de Código

Una herramienta desarrollada en Python para generar hojas de código automáticas a partir de datos de encuestas en Excel.

## 📋 Descripción

Este proyecto es un generador automático de hojas de código que procesa datos de encuestas y genera documentación estructurada tanto en formato Excel como Word. La aplicación cuenta con una interfaz gráfica intuitiva que permite a los usuarios cargar archivos de Excel y generar automáticamente hojas de código con preguntas, alternativas y estadísticas de respuesta.

## ⚠️ Importante - Datos de Prueba

**NINGÚN CORREO ELECTRÓNICO HA SIDO FILTRADO NI COMPROMETIDO**

Los datos incluidos en este proyecto (archivos Excel) son únicamente **listas generadas aleatoriamente** con fines de prueba y demostración. No contienen información real de usuarios ni correos electrónicos reales.

## 🚀 Características

- **Carga automática de datos**: Importa archivos Excel con datos de encuestas
- **Generación automática de códigos**: Crea hojas de código estructuradas
- **Múltiples formatos de salida**: Genera archivos en Excel (.xlsx) y Word (.docx)
- **Interfaz gráfica amigable**: Aplicación de escritorio con Tkinter
- **Detección automática de preguntas**: Identifica columnas relevantes automáticamente
- **Estadísticas de respuesta**: Calcula frecuencias y porcentajes
- **Vista previa en tiempo real**: Muestra el contenido antes de generar

## 🛠️ Tecnologías Utilizadas

- **Python 3.x**
- **pandas**: Manipulación de datos
- **openpyxl**: Procesamiento de archivos Excel
- **python-docx**: Generación de documentos Word
- **tkinter**: Interfaz gráfica de usuario
- **numpy**: Operaciones numéricas

## 📦 Instalación

1. Clona este repositorio:
```bash
git clone https://github.com/Luis000923/generador-de-hojas-de-codigo.git
cd generador-de-hojas-de-codigo
```

2. Instala las dependencias requeridas:
```bash
pip install pandas openpyxl python-docx numpy
```

## 🎯 Uso

1. Ejecuta la aplicación:
```bash
python hoja_de_codigos.py
```

2. Usa la interfaz gráfica para:
   - Cargar tu archivo Excel con los datos de la encuesta
   - Previsualizar las preguntas detectadas
   - Configurar opciones de generación (opcional)
   - Generar la hoja de código en el formato deseado

## 📁 Estructura del Proyecto

```
├── hoja_de_codigos.py          # Aplicación principal
├── ejemplo.xlsx                # Archivo de ejemplo (datos de prueba)
├── Impacto de la alimentación encuesta (respuestas).xlsx  # Ejemplo de encuesta
├── ejemplos/                   # Carpeta con ejemplos
│   └── image.png              # Imagen de ejemplo
└── README.md                  # Este archivo
```

## 🔧 Funcionalidades Principales

### Clase `SurveyProcessor`
- `load_excel()`: Carga y procesa archivos Excel
- `generate_coding_sheet()`: Genera la hoja de código
- `generate_word_document()`: Crea documentos Word
- `generate_excel_document()`: Crea documentos Excel

### Aplicación GUI (`SurveyApp`)
- Interfaz gráfica intuitiva
- Sistema de previsualización
- Múltiples opciones de exportación
- Barra de estado informativa

## 📊 Formato de Entrada

El programa acepta archivos Excel (.xlsx) con las siguientes características:
- Primera fila como encabezados
- Columnas administrativas (correos, timestamps) son filtradas automáticamente
- Datos de respuestas en las columnas restantes

## 📄 Formato de Salida

### Documento Word
- Encabezado con información del proyecto
- Tabla estructurada con preguntas y alternativas
- Estadísticas de frecuencia y porcentajes

### Archivo Excel
- Hoja de código formateada
- Tablas con bordes y estilos
- Cálculos automáticos de estadísticas

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