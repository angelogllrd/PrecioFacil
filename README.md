<p align="center">
  <img width="245" height="245" alt="precio_facil_logo" src="https://github.com/user-attachments/assets/a309bec7-4161-4392-aab1-18259777865e" />
</p>

# 🛒 PrecioFacil

**PrecioFacil** es una aplicación de escritorio diseñada para centralizar, actualizar y agilizar la búsqueda de listas de precios de proveedores específicos del sector agroindustrial y ferretero ([Tienda del Cardan](https://tiendadecardan.com.ar/), [Bulonera Camba](https://buloneracamba.com.ar/), [Rosario Agro Industrial](https://www.rosarioagroindustrial.com/)) y [VTM Transmisiones](https://vtmtransmisiones.com/).

<img width="1280" height="720" alt="2026-08-31 17-04-34 (2)" src="https://github.com/user-attachments/assets/d05ac143-3de9-45f3-aa01-2dc578bdf98c" />

Desarrollada en Python con una interfaz gráfica moderna en PyQt6, la herramienta descarga automáticamente las últimas listas de precios (Excel y PDF), las procesa y permite realizar búsquedas instantáneas, visualizar catálogos y trabajar sin conexión.

## ✨ Características Principales

* 🔄 **Actualización Automática de Listas:** Descarga los archivos Excel y PDF más recientes directamente desde las webs de los proveedores.
* ⚡ **Búsqueda Ultrarrápida:** Filtrado de productos en tiempo real por código, subcategoría o descripción sin demoras en la interfaz.
* 🖱️ **Acceso Directo a Productos:** Al hacer doble clic sobre cualquier artículo, este se abre automáticamente en la web del proveedor (incluye auto-login en segundo plano para acceder a precios con descuento en Tienda del Cardan y búsqueda automatizada en VTM).
* 📴 **Modo Offline (Respaldo Local):** Si no hay conexión a internet o la página del proveedor está caída, la aplicación utiliza inteligentemente la última lista descargada guardada de forma local.
* 📖 **Integración con Catálogos PDF:** Abre catálogos nativos en la página exacta del producto consultado, usando el navegador predeterminado o el visor de PDF del sistema.
* 🌬️ **Gestión de Ventiladores de aluminio:** Nueva pestaña con una tabla para registrar precios de ventiladores de aluminio, con autoguardado local.
* 🚀 **Actualizador Automático:** Comprueba automáticamente si hay nuevas versiones del software, descarga el instalador y lo ejecuta de forma silenciosa.
* 🌗 **Temas Visuales:** Soporte nativo para Modo Claro y Modo Oscuro.
* 📊 **Reporte de Estado:** Brinda un informe detallado en caso de problemas con la descarga y el procesamiento de las listas.

## 🛠️ Tecnologías Utilizadas

El proyecto está construido íntegramente en **Python** y aprovecha múltiples librerías para el manejo de web scraping, automatización web, bases de datos, hilos y UI:

* 🖥️ **[PyQt6](https://riverbankcomputing.com/software/pyqt/intro):** Framework principal para la Interfaz Gráfica de Usuario (GUI).
* 🎭 **[Playwright](https://playwright.dev/python/):** Utilizado para la automatización del navegador, manejo de inicio de sesión automático y navegación dinámica en los catálogos web.
* 🌐 **[Requests](https://requests.readthedocs.io/):** Para las peticiones HTTP y descargas en red.
* 🍲 **[BeautifulSoup4](https://beautiful-soup-4.readthedocs.io/):** Para el scraping y parseo del HTML de los proveedores.
* 📊 **[Openpyxl](https://openpyxl.readthedocs.io/):** Para la lectura y extracción rápida de datos de los archivos `.xlsx`.
* 🗄️ **SQLite:** Motor de base de datos integrado para el almacenamiento persistente, ligero y automático de celdas editables.
* 🧵 **Hilos Concurrentes (`QThread` y `ThreadPoolExecutor`):** Para garantizar un rendimiento fluido y evitar que la interfaz se congele durante descargas pesadas o automatizaciones del navegador. 

## 🚀 Instalación y Uso

1. Andá a la sección de [Releases](https://github.com/angelogllrd/PrecioFacil/releases/) en la derecha de este repositorio.
2. Descarga el instalador `PrecioFacil-x.x.x-Setup.exe`.
3. Instala y ejecuta. ¡La aplicación se mantendrá actualizada sola!
