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

- **Python 3.7+**
- **Streamlit**
- **openpyxl**

> ⚠️ **Importante**: Se recomienda encarecidamente usar un **entorno virtual** para evitar conflictos entre dependencias de diferentes proyectos.

## 📦 Instalación

### Opción 1: Instalación con Entorno Virtual (Recomendado)

1. **Clona el repositorio**:
   ```bash
   git clone https://github.com/Jhan-Pierre/UnirExcel.git
   cd UnirExcel
   ```

2. **Crea un entorno virtual**:
   ```bash
   # Windows
   python -m venv venv
   
   # Linux/macOS
   python3 -m venv venv
   ```

3. **Activa el entorno virtual**:
   ```bash
   # Windows
   venv\Scripts\activate
   
   # Linux/macOS
   source venv/bin/activate
   ```

4. **Instala las dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

### Opción 2: Instalación Global (No recomendado)

1. **Clona el repositorio**:
   ```bash
   git clone https://github.com/Jhan-Pierre/UnirExcel.git
   cd UnirExcel
   ```

2. **Instala las dependencias directamente**:
   ```bash
   pip install streamlit openpyxl
   ```

## 🚀 Uso

1. **Asegúrate de que el entorno virtual esté activado** (si usaste la Opción 1):
   ```bash
   # Windows
   venv\Scripts\activate
   
   # Linux/macOS
   source venv/bin/activate
   ```

2. **Ejecuta la aplicación**:
   ```bash
   streamlit run app.py
   ```

3. **Abre tu navegador**:
   - La aplicación se abrirá automáticamente en `http://localhost:8501`
   - Si no se abre automáticamente, copia la URL desde la terminal

4. **Sube tus archivos**:
   - Selecciona uno o múltiples archivos Excel (`.xlsx` o `.xlsm`)
   - Sube tu plantilla base donde se consolidarán los datos

5. **Procesa los datos**:
   - Haz clic en "🚀 Procesar"
   - Observa la barra de progreso mientras se procesan los archivos

6. **Descarga el resultado**:
   - Usa el botón "📥 Descargar Resultado" para obtener el archivo consolidado

### 🛑 Para detener la aplicación:
- Presiona `Ctrl + C` en la terminal donde se está ejecutando

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
├── requirements.txt    # Dependencias del proyecto
├── venv/              # Entorno virtual (creado después de la instalación)
│   ├── Scripts/       # Ejecutables del entorno virtual (Windows)
│   ├── bin/           # Ejecutables del entorno virtual (Linux/macOS)
│   └── ...
└── .gitignore         # Archivos a ignorar por Git (recomendado)
```

> 📝 **Nota**: La carpeta `venv/` se crea automáticamente al seguir las instrucciones de instalación y no debe subirse al repositorio.

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
- **Entorno Virtual**: Se recomienda usar un entorno virtual para evitar conflictos de dependencias

## 🔧 Requisitos del Sistema

- **Python**: 3.7 o superior
- **Sistema Operativo**: Windows, macOS, Linux
- **Memoria RAM**: Mínimo 2GB (recomendado 4GB para archivos grandes)
- **Espacio en disco**: 100MB para la instalación + espacio para archivos temporales

## 🚨 Solución de Problemas

### Error: "streamlit: command not found"
```bash
# Activa el entorno virtual primero
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/macOS

# Luego ejecuta streamlit
streamlit run app.py
```

### Error de permisos en Windows
```bash
# Ejecuta PowerShell como administrador y ejecuta:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### La aplicación no abre en el navegador
- Copia manualmente la URL desde la terminal (usualmente `http://localhost:8501`)
- Verifica que el puerto 8501 no esté en uso por otra aplicación

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Si encuentras algún error o tienes sugerencias de mejora:

1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.


**Desarrollado para DIRNIC** - Sistema de consolidación de datos de requisitorias y detenidos
