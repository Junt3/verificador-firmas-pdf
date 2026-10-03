# 📑 Verificador de Firmas PDF a Excel

[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Tkinter](https://img.shields.io/badge/Tkinter-GUI-blue?style=for-the-badge)](https://docs.python.org/3/library/tkinter.html)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel-green?style=for-the-badge)](https://openpyxl.readthedocs.io/en/stable/)
[![PyPDF](https://img.shields.io/badge/PyPDF-PDF-red?style=for-the-badge)](https://pypdf.readthedocs.io/en/stable/)

Una herramienta de escritorio desarrollada en Python que automatiza la auditoría de firmas electrónicas en lotes de archivos PDF. A través de una interfaz gráfica intuitiva, el usuario selecciona una carpeta con documentos PDF y la aplicación analiza los metadatos de cada uno para determinar si están firmados digitalmente, extrayendo el nombre del firmante y la fecha, para finalmente consolidar la información en un reporte estructurado de Excel.

## 🚀 Características Principales

*   **Interfaz Gráfica de Usuario (GUI):** Implementada con `tkinter` para facilitar la selección de carpetas y directorios de guardado sin necesidad de usar la línea de comandos.
*   **Procesamiento por Lotes:** Analiza automáticamente todos los archivos PDF contenidos en un directorio seleccionado.
*   **Extracción de Metadatos:** Utiliza `pypdf` para leer los campos del formulario del PDF e identificar firmas digitales (`/Sig`), extrayendo el nombre (`/Name`) y la fecha (`/M`).
*   **Generación de Reportes en Excel:** Emplea `openpyxl` para crear un archivo `.xlsx` estilizado (con encabezados en negrita, colores de fondo y bordes) y autoajuste del ancho de las columnas.
*   **Manejo de Errores:** Incluye bloques `try-except` para gestionar PDFs protegidos o sin firmas válidas sin interrumpir el proceso.

## 🛠️ Tecnologías Utilizadas

*   **Lenguaje:** Python 3.x
*   **Librerías Estándar:** `os`, `tkinter`
*   **Librerías de Terceros:** `pypdf`, `openpyxl`

## ⚙️ Instalación y Uso

1.  **Clonar el repositorio:**
    ```bash
    git clone [https://github.com/Junt3/verificador-firmas-pdf.git](https://github.com/Junt3/verificador-firmas-pdf.git)
    cd verificador-firmas-pdf
    ```

2.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Ejecutar el script:**
    ```bash
    python verificador_firmas.py
    ```

4.  **Uso de la aplicación:**
    *   Al ejecutar, se abrirá una ventana para seleccionar un archivo PDF de muestra dentro de la carpeta que deseas analizar.
    *   Luego, elige dónde guardar el reporte Excel.
    *   El programa procesará los archivos y mostrará un mensaje de éxito.

---
*Este proyecto es parte de mi portafolio como Ingeniero en Sistemas, demostrando habilidades en automatización de tareas y desarrollo con Python.*
