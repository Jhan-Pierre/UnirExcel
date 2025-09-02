# 📊 Unir Excels - Consolidador de Archivos Excel

Una aplicación web desarrollada con Streamlit que permite consolidar múltiples archivos Excel en una plantilla base, manteniendo el formato y las macros originales.

## 🚀 Características

- ✅ **Interfaz Web Intuitiva**: Aplicación fácil de usar con Streamlit
- ✅ **Soporte Múltiple**: Compatible con archivos `.xlsx` y `.xlsm`
- ✅ **Preservación de Macros**: Mantiene las macros VBA en archivos `.xlsm`
- ✅ **Formato Automático**: Aplica formato de fuente Calibri 10pt
- ✅ **Formato de Fechas**: Convierte automáticamente fechas al formato `dd/mm/yyyy`
- ✅ **Barra de Progreso**: Seguimiento visual del proceso de consolidación
- ✅ **Filtrado Inteligente**: Solo procesa filas con más de 10 columnas de datos
- ✅ **Trazabilidad**: Añade referencias del archivo origen en el resultado

## 🔧 Requisitos

- Python 3.7+
- Streamlit
- openpyxl

## 📦 Instalación

1. **Clona o descarga el proyecto**:
   ```bash
   git clone [url-del-repositorio]
   cd Requisitoria-Detenidos
   ```

2. **Instala las dependencias**:
   ```bash
   pip install streamlit openpyxl
   ```

   O usando un archivo `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Uso

1. **Ejecuta la aplicación**:
   ```bash
   streamlit run app.py
   ```

2. **Sube tus archivos**:
   - Selecciona uno o múltiples archivos Excel (`.xlsx` o `.xlsm`)
   - Sube tu plantilla base donde se consolidarán los datos

3. **Procesa los datos**:
   - Haz clic en "🚀 Procesar"
   - Observa la barra de progreso mientras se procesan los archivos

4. **Descarga el resultado**:
   - Usa el botón "📥 Descargar Resultado" para obtener el archivo consolidado

## 📋 Funcionalidades Detalladas

### Procesamiento de Datos
- **Filtrado**: Solo procesa filas que contengan más de 10 columnas con datos
- **Limpieza**: Ignora filas vacías y celdas sin contenido
- **Cabeceras**: Omite la primera fila (cabecera) de cada archivo

### Formato y Estilo
- **Fuente**: Aplica Calibri 10pt a todos los datos
- **Fechas**: Formato automático `dd/mm/yyyy` para campos de fecha
- **Referencias**: Añade marcadores indicando el archivo origen de cada grupo de datos

### Compatibilidad
- **Archivos Excel**: `.xlsx` y `.xlsm`
- **Macros VBA**: Preserva macros en archivos `.xlsm`

## 🏗️ Estructura del Proyecto

```
Requisitoria-Detenidos/
│
├── app.py              # Aplicación principal de Streamlit
├── README.md           # Este archivo
└── requirements.txt    # Dependencias del proyecto (opcional)
```

## 🔍 Ejemplo de Uso

1. **Archivos de entrada**: Múltiples Excel con datos de requisitorias/detenidos
2. **Plantilla**: Archivo Excel base con formato y posibles macros
3. **Resultado**: Archivo consolidado con todos los datos organizados

## 🛠️ Tecnologías Utilizadas

- **[Streamlit](https://streamlit.io/)**: Framework para aplicaciones web de datos
- **[OpenPyXL](https://openpyxl.readthedocs.io/)**: Librería para manipulación de archivos Excel
- **Python**: Lenguaje de programación principal

## 📝 Notas Técnicas

- La aplicación utiliza carpetas temporales para el procesamiento de archivos
- Los archivos se procesan secuencialmente para evitar problemas de memoria
- Se mantiene la integridad de las macros VBA en archivos `.xlsm`
- El formato de fecha sigue el estándar español `dd/mm/yyyy`


**Desarrollado para DIRNIC** - Sistema de consolidación de datos de requisitorias y detenidos