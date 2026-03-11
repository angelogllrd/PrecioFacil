import openpyxl
import requests
import bs4
import tempfile
from pathlib import Path
import zipfile
from io import BytesIO
import os
from urllib.parse import urlparse, unquote
import re
from datetime import datetime
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import QMainWindow, QApplication, QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QMessageBox
from PyQt6.QtCore import Qt, QUrl, QObject, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont, QDesktopServices
import winreg
import subprocess
from PyQt6 import uic
import sys
from utils import SETTINGS, CAMBA_SHEETS, CAMBA_CATEGORIES, ROSARIO_URLS, MOST_USED_PRODUCTS_HH, MOST_USED_PRODUCTS_ETMA, MOST_USED_PRODUCTS_CAMBA 



class Worker(QObject):
	# Señales para actualizar el ProgressDialog
	progress_changed = pyqtSignal(int)
	message_changed = pyqtSignal(str)

	# Señal final que emite un diccionario con productos + reporte
	finished = pyqtSignal(dict)


	def update_progress(self, points_to_add, message=None):
		"""Suma puntos al progreso total y actualiza la UI."""

		self.current_progress += points_to_add
		
		# Evito pasarme de 100 por si hay algún redondeo raro
		if self.current_progress > 100:
			self.current_progress = 100
			
		if message:
			self.message_changed.emit(message)
		
		# Emito el entero a la barra de progreso
		self.progress_changed.emit(int(self.current_progress))

		# print(f'{self.current_progress:.2f}', message if message else '')


	# CÓDIGO PRINCIPAL
	# ------------------------------------------------------------------------------------------

	def run(self):
		"""
		Método principal ejecutado por el hilo secundario para gestionar la descarga
		y procesamiento de las listas.
		"""

		# Inicializo variables
		self.current_progress = 0
		self.report = {}
		self.all_data = {
			'hh': {'products': [], 'date': ''},
			'etma': {'products': [], 'date': ''},
			'camba': {'products': [], 'date': ''},
			'report': self.report
		}

		# Calculo puntajes de progreso
		camba_files = 1 + len(CAMBA_SHEETS) # 1 excel + N pdfs
		rosario_files = len(ROSARIO_URLS)
		self.points_per_brand = 25
		self.points_per_file_camba = self.points_per_brand / camba_files
		self.points_per_file_rosario = self.points_per_brand / rosario_files

		self.update_progress(0, 'Iniciando carga...')

		suppliers = {
			'tdc':   ('hh', 'etma'),
			'camba': ('camba',)
		}

		for supplier, brands in suppliers.items():

			# Recupero la URL del proveedor
			supplier_url = self.get_url_from_settings(supplier)
			if not supplier_url:
				self.handle_supplier_down(brands, 'no_url')
				continue

			# Obtengo el HTML de esa URL y lo parseo
			try:
				html = self.download_html(supplier_url)
				soup = bs4.BeautifulSoup(html, 'html.parser')
			except Exception:
				self.handle_supplier_down(brands, 'no_access')
				continue

			# Proceso los excel de cada marca
			for brand in brands:
				self.update_progress(0, f'Procesando {brand.upper()}...')
				self.process_brand(soup, brand)

			# Si el proveedor es CAMBA, busco los PDF de las hojas
			if supplier == 'camba':
				self.process_camba_pdfs(soup)

		# Proceso los links fijos de ROSARIO
		self.update_progress(0, 'Procesando ROSARIO AGRO...')
		self.process_rosario_pdfs()

		self.update_progress(100, '¡Carga completada!')

		# Devuelvo todos los datos recolectados al MainWindow
		self.finished.emit(self.all_data)


	# PROCESAMIENTO POR MARCA
	# ------------------------------------------------------------------------------------------

	def process_brand(self, soup, brand):
		"""Busca la URL de la lista excel en el soup, la descarga y la procesa."""

		if brand == 'camba':
			step_points = self.points_per_file_camba / 3
		else:
			step_points = self.points_per_brand / 3

		# Busco link de la lista
		list_url = self.get_list_url_from_soup(soup, brand)
		self.update_progress(step_points)
		if not list_url:
			self.check_local_excel_list(brand, 'no_link')
			self.update_progress(step_points * 2) # Como fui al fallback, sumo de golpe los 2 pasos restantes (descargar y procesar)
			return

		# Descargo la lista
		try:
			excel_file_path = self.download_excel_file(list_url, brand)

			# Tengo archivo nuevo: actualizo fecha de validez si es CAMBA
			if brand == 'camba':
				camba_last_date = self.resolve_camba_date(soup, excel_file_path)
				SETTINGS.setValue('camba_last_date', camba_last_date)

			self.update_progress(step_points)

		except Exception:
			self.check_local_excel_list(brand, 'no_download')
			self.update_progress(step_points) # Solo sumo el paso restante (procesar)
			return

		# Proceso excel descargado
		try:
			self.process_excel(excel_file_path, brand)
		except Exception as e:
			self.report.setdefault(brand, {})['excel'] = {
				'local_status': 'local_error'
			}

		self.update_progress(step_points)


	def process_camba_pdfs(self, soup):
		"""Procesa los PDF de las hojas de Bulonera Camba buscando en el soup."""

		# Construyo ruta de la carpeta destino
		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / 'camba'
		base_path.mkdir(parents=True, exist_ok=True)

		for sheet_num in CAMBA_SHEETS:
			# Busco el link de la hoja
			a_elem = soup.find(
				'a',
				href=True,
				string=lambda s: s and f'Hoja {sheet_num}' in s
			)
			if not a_elem:
				self.check_local_pdf_list('camba', sheet_num, 'no_link')
				self.update_progress(self.points_per_file_camba) # Sumo antes de saltar
				continue

			# Obtengo la ruta completa
			pdf_url = a_elem['href']
			pdf_original_name = pdf_url.split('=')[-1] + '.pdf'
			pdf_file_path = base_path / pdf_original_name

			# Si ya existe con este nombre exacto, lo salteamos
			if pdf_file_path.exists():
				self.update_progress(self.points_per_file_camba) # Sumo antes de saltar
				continue

			# Descargo el PDF
			try:
				response = requests.get(pdf_url, timeout=10)
				response.raise_for_status()

				# Borro versiones viejas del mismo número de hoja
				for old_pdf in base_path.glob(f'Hoja{sheet_num}*.pdf'):
					old_pdf.unlink()

				# Guardo el archivo descargado
				with open(pdf_file_path, 'wb') as f:
					f.write(response.content)
			except Exception:
				self.check_local_pdf_list('camba', sheet_num, 'no_download')

			# Sumo al final si todo el proceso normal terminó
			self.update_progress(self.points_per_file_camba)


	def process_rosario_pdfs(self):
		"""Descarga los PDF de ROSARIO AGRO desde links fijos, sobreescribiendo."""

		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / 'rosario'
		base_path.mkdir(parents=True, exist_ok=True)

		for pdf_url in ROSARIO_URLS:
			# Obtengo la ruta completa
			pdf_original_name = pdf_url.split('=')[-1]
			pdf_file_path = base_path / pdf_original_name

			# Descargo el PDF
			try:
				response = requests.get(pdf_url, timeout=10)
				response.raise_for_status()

				with open(pdf_file_path, 'wb') as f:
					f.write(response.content)
			except Exception:
				self.check_local_pdf_list('rosario', pdf_file_path.stem, 'no_download') # paso solo nombre del PDF
			
			self.update_progress(self.points_per_file_rosario)


	# FALLBACKS (hubo error y se deben buscar listas locales descargadas previamente)
	# ------------------------------------------------------------------------------------------

	def handle_supplier_down(self, brands, reason):
		"""
		Fallback general llamado cuando:
		  * No hay URL del proveedor.
		  * No se pudo obtener HTML de la URL.
		Llama al fallback de marca por cada marca del proveedor. Además, 
		si es CAMBA, chequea las hojas PDF locales.
		"""

		for brand in brands:
			self.update_progress(0, f'Procesando {brand.upper()}...')
			self.check_local_excel_list(brand, reason)

			# Si es CAMBA, compruebo PDFs locales
			if brand == 'camba':
				self.update_progress(self.points_per_file_camba) # excel recién procesado
				for sheet_num in CAMBA_SHEETS:
					self.check_local_pdf_list('camba', sheet_num, reason)
					self.update_progress(self.points_per_file_camba)
			else:
				self.update_progress(self.points_per_brand)


	def check_local_excel_list(self, brand, reason):
		"""
		Fallback por marca llamado cuando:
		  * No se encontró el link de la lista en el HTML.
		  * No se pudo descargar el excel.
		Comprueba si existe un excel local previamente descargado y lo procesa.
		"""

		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand
		excel_file_path = None

		if base_path.exists():
			excel_files = list(base_path.glob('*.xlsx'))
			if excel_files:
				excel_file_path = excel_files[0]

		if excel_file_path:
			try:
				self.process_excel(excel_file_path, brand)
				local_status = 'local_used'
			except Exception:
				local_status = 'local_error'
		else:
			local_status = 'local_missing'

		self.report.setdefault(brand, {})['excel'] = {
			'reason': reason,
			'local_status': local_status
		}


	def check_local_pdf_list(self, brand, identifier, reason):
		"""
		Verifica existencia local de un PDF cuando falla descarga/encontrar link.
		* Para CAMBA: identifier es el número de hoja (por ej: '02')
		* Para ROSARIO: identifier es el nombre del archivo (ej: 'Cuchillas_Jardin')
		"""

		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand
		has_local = False

		if base_path.exists():
			if brand == 'camba':
				# Busco PDFs con el número de hoja
				pdf_files = list(base_path.glob(f'Hoja{identifier}*.pdf'))
				if pdf_files:
					has_local = True
			elif brand == 'rosario':
				if (base_path / f'{identifier}.pdf').exists():
					has_local = True

		# Guardo resultado para agruparlo después
		pdfs = self.report.setdefault(brand, {}).setdefault('pdfs', {})
		entry = pdfs.setdefault(reason, {'missing': [], 'local': []})
		if has_local:
			entry['local'].append(identifier)
		else:
			entry['missing'].append(identifier)


	#  AUXILIARES
	# ------------------------------------------------------------------------------------------

	def get_url_from_settings(self, supplier):
		return SETTINGS.value(f'supplier_urls/{supplier}', '', type=str)


	def download_html(self, url):
		response = requests.get(url, timeout=10)
		response.raise_for_status()
		return response.text


	def get_list_url_from_soup(self, soup, brand):
		"""Obtiene del sitio del proveedor el link actual de la lista correspondiente."""

		if brand in ('hh', 'etma'):
			# Busco el título correcto
			h1_elem = soup.find(
				'h1', 
				string=lambda s: s and s.strip().upper() in (
					f'LISTA DE PRECIO {brand.upper()}',
					f'LISTA DE PRECIOS {brand.upper()}'
				)
			)
			if not h1_elem:
				return None

			# Subo al bloque contenedor
			container = h1_elem.find_parent('div', class_='widget-span')
			if not container:
				return None

			# Busco el link dentro del bloque
			a_elem = container.find_next('a', href=True)

		elif brand == 'camba':
			# Busco el título correcto
			h2_elem = soup.find(
				'h2', 
				string=lambda s: s and 'Lista de precios formato sabana' in s.strip()
			)
			if not h2_elem:
				return None

			# Busco el link
			a_elem = h2_elem.find_parent('a')

		return a_elem['href'] if a_elem else None


	def download_excel_file(self, url, brand):
		"""Descarga el excel en la carpeta correspondiente."""

		# Construyo ruta de la carpeta destino
		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand
		base_path.mkdir(parents=True, exist_ok=True)

		# Descargo el archivo
		response = requests.get(url, timeout=10)
		response.raise_for_status()

		# Borro excel previo
		for old_excel_file in base_path.glob('*.xlsx'):
			old_excel_file.unlink()

		if brand in ('hh', 'etma'): # la URL entrega un excel
			# Obtengo el nombre original del archivo desde la URL
			excel_original_name = os.path.basename(urlparse(url).path)
			excel_original_name = unquote(excel_original_name) # reemplaza %20 por espacios

			# Obtengo la ruta completa
			excel_file_path = base_path / excel_original_name			

			# Guardo el archivo descargado
			with open(excel_file_path, 'wb') as f:
				f.write(response.content)
		else: # la URL entrega un zip
			with zipfile.ZipFile(BytesIO(response.content)) as z:
				for name in z.namelist():
					if name.lower().endswith('.xlsx'):
						excel_file_path = base_path / name
						with z.open(name) as source, open(excel_file_path, 'wb') as target:
							target.write(source.read())
						break

		return excel_file_path


	def resolve_camba_date(self, soup, excel_file_path):
		"""
		Determina qué fecha asociar al excel descargado de CAMBA usando, por prioridad:
		1) HTML
		2) nombre del archivo
		3) fecha actual
		"""

		# Busco en el HTML
		date = self.extract_camba_date_from_soup(soup)
		if date:
			return date

		# Busco en el nombre de archivo
		date = self.extract_date_from_filename(excel_file_path)
		if date:
			return date

		# Como último recurso: fecha actual
		return datetime.now().strftime('%d/%m/%Y')


	def extract_camba_date_from_soup(self, soup):
		"""Extrae la fecha actual de las listas de CAMBA desde el HTML."""

		a_elem = soup.find(
			'a',
			href=True,
			string=lambda s: s and 'lista indice' in s.strip().lower()
		)

		if not a_elem:
			return None

		match = re.search(r'\d{2}/\d{2}/\d{4}', a_elem.get_text())
		return match.group() if match else None


	def extract_date_from_filename(self, path):
		"""Extrae la fecha desde el nombre del archivo excel de CAMBA."""
		
		match = re.search(r'\d{2}-\d{2}-\d{4}', path)
		if match:
			return match.group().replace('-', '/')
		return None


	def process_excel(self, excel_file_path, brand):
		"""
		Lee el excel, extrae la fecha de validez y los productos, y guarda todo 
		en all_data.
		"""

		# Creo workbook y extraigo la hoja de productos
		wb = openpyxl.load_workbook(excel_file_path)
		sheet = wb[wb.sheetnames[0]]

		# Busco letras de columnas de producto
		header_cols = self.search_header_cols(sheet, brand)

		# Busco número de fila de primer producto
		first_row = self.search_first_row(sheet, header_cols['price_col'])

		# Extraigo la fecha de validez de precios
		validity_date = self.extract_validity_date(brand, sheet)

		# Extraigo todos los productos en una lista de diccionarios
		products = self.obtain_products(sheet, first_row, header_cols, brand)

		# Guardo los datos recolectados en el diccionario general
		self.all_data[brand]['products'] = products
		self.all_data[brand]['date'] = validity_date


	def search_header_cols(self, sheet, brand):
		"""Retorna un diccionario con la posición (letra) de cada columna."""

		default_cols = {
			'hh': {
				'code_col': 'A',
				'subcategory_col': 'B',
				'description_col': 'C',
				'price_col': 'E'
			},
			'etma': {
				'code_col': 'A',
				'subcategory_col': 'B',
				'description_col': 'C',
				'price_col': 'E'
			},
			'camba': {
				'code_col': 'C',
				'subcategory_col': 'J',
				'description_col': 'B',
				'price_col': 'E'
			}
		}

		# Valores por defecto en caso de que no encuentre alguna
		code_col = default_cols[brand]['code_col']
		subcategory_col = default_cols[brand]['subcategory_col']
		description_col = default_cols[brand]['description_col']
		price_col = default_cols[brand]['price_col']

		for row in sheet['A1':'E20']:
			for cell in row:
				# Encuentro la fila de encabezados
				if str(cell.value).strip().lower() in ('codigo', 'código', 'cod', 'cód'):
					header_row = cell.row
					code_col = get_column_letter(cell.column)

					# Recorro fila de encabezados e identifico cada uno
					for header_cell in sheet[header_row]:
						value = str(header_cell.value).strip().lower()
						if value in ('subrubro', 'sub rubro', 'rubro'):
							subcategory_col = get_column_letter(header_cell.column)
						elif value in ('descripción', 'descripcion', 'desc', 'articulo', 'artículo'):
							description_col = get_column_letter(header_cell.column)
						elif value in ('precio + iva', 'precio'):
							price_col = get_column_letter(header_cell.column)
					break

		return {
			'code_col': code_col,
			'subcategory_col': subcategory_col,
			'description_col': description_col,
			'price_col': price_col
		}


	def search_first_row(self, sheet, price_col):
		"""Retorna la fila donde comienzan los productos."""

		for cell in sheet[price_col]:
				value = cell.value

				# Evito trabajo innecesario (evito convertir None)
				if value is None:
					continue

				# Compruebo si es un monto
				try:
					float(str(value).replace('.', '').replace(',', '.'))
					return cell.row
				except (ValueError, TypeError):
					continue


	def extract_validity_date(self, brand, sheet):
		"""Busca y retorna la fecha de validez del excel de la marca."""

		# Para CAMBA se busca en la configuración guardada
		if brand == 'camba':
			stored_date = SETTINGS.value('camba_last_date', '', type=str)
			if stored_date:
				return f'📆 Precios válidos para el: {stored_date}'
			else:
				return '📆 Fecha no disponible'

		# Para HH o ETMA se busca en las primeras celdas
		for row in sheet['A1':'E20']:
			for cell in row:
				# Evito trabajo innecesario (no analizo celdas vacías)
				if not cell.value:
					continue

				# Busco la fecha en la celda
				value = str(cell.value)
				if re.search(r'\d{1,2}/\d{1,2}/\d{2,4}', value) and ('valid' in value or 'válid' in value):
					return '📆 ' + value.replace('validos', 'válidos')

		return '📆 Fecha no encontrada'


	def obtain_products(self, sheet, first_row, header_cols, brand):
		"""Crea lista de diccionarios de productos para filtrar."""

		products = []
		for row in range(first_row, sheet.max_row + 1):
			if self.is_valid_row(sheet, row, header_cols):
				# Formateo precios de tipo float (necesario en CAMBA)
				price = sheet[header_cols['price_col'] + str(row)].value
				if isinstance(price, float):
					price = f'{price:,}'.replace('.', '_').replace(',', '.').replace('_', ',')

				# Mapeo categoría si es CAMBA
				subcategory = sheet[header_cols['subcategory_col'] + str(row)].value
				if brand == 'camba':
					subcategory = CAMBA_CATEGORIES.get(subcategory, 'CATEGORIA NO DEFINIDA')

				# Creo el diccionario y lo agrego a la lista
				product = {
					'code': sheet[header_cols['code_col'] + str(row)].value,
					'subcategory': subcategory,
					'description': sheet[header_cols['description_col'] + str(row)].value,
					'price': f'$ {price}'
				}
				products.append(product)
		return products


	def is_valid_row(self, sheet, row, header_cols):
		"""Retorna si una fila corresponde o no a un producto."""

		for col in header_cols.values():
			if sheet[col + str(row)].value is None:
				return False
		return True



class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo la UI
		uic.loadUi('ui/app.ui', self)

		# Señales de pushbuttons inferiores
		self.pushButton_theme.clicked.connect(self.change_theme)
		self.pushButton_config.clicked.connect(self.open_config)
		self.pushButton_about.clicked.connect(self.open_about)

		# Señales de pushbuttons de BULONERA CAMBA
		self.pushButton_alemite.clicked.connect(lambda: self.open_pdf('camba', '22', 2))
		self.pushButton_seeger.clicked.connect(lambda: self.open_pdf('camba', '35', 2))
		self.pushButton_arandela_grower.clicked.connect(lambda: self.open_pdf('camba', '10', 4))
		self.pushButton_arandela_plana.clicked.connect(lambda: self.open_pdf('camba', '16', 1))
		self.pushButton_bulon_unc.clicked.connect(lambda: self.open_pdf('camba', '02', 1))
		self.pushButton_bulon_unf.clicked.connect(lambda: self.open_pdf('camba', '07', 1))
		self.pushButton_chaveta_partida.clicked.connect(lambda: self.open_pdf('camba', '19', 1))
		self.pushButton_espina_elastica.clicked.connect(lambda: self.open_pdf('camba', '35', 1))
		self.pushButton_prisionero_cilindrica.clicked.connect(lambda: self.open_pdf('camba', '14', 2))
		self.pushButton_prisionero_sin.clicked.connect(lambda: self.open_pdf('camba', '14', 3))
		self.pushButton_prisionero_cuadrada.clicked.connect(lambda: self.open_pdf('camba', '13', 1))
		self.pushButton_tuerca_exagonal.clicked.connect(lambda: self.open_pdf('camba', '04', 1))
		self.pushButton_tuerca_castillo.clicked.connect(
			lambda: [
				self.open_pdf('camba', '04', 5),
				self.open_pdf('camba', '23', 1)
			]
		)
		self.pushButton_tuerca_torneada.clicked.connect(lambda: self.open_pdf('camba', '23', 1))
		self.pushButton_varilla_camba.clicked.connect(
			lambda: [
				self.open_pdf('camba', '11', 2),
				self.open_pdf('camba', '17', 1)
			]
		)
		self.pushButton_tornillo_metrico.clicked.connect(lambda: self.open_pdf('camba', '13', 2))
		self.pushButton_tornillo_inox.clicked.connect(lambda: self.open_pdf('camba', '36', 8))

		# Señales de pushbuttons de ROSARIO AGRO
		self.pushButton_gummi.clicked.connect(lambda: self.open_pdf('rosario', 'GUMMI'))
		self.pushButton_tupac.clicked.connect(lambda: self.open_pdf('rosario', 'Tupac'))
		self.pushButton_cadena.clicked.connect(lambda: self.open_pdf('rosario', 'Cadenas_LinkBelt'))
		self.pushButton_cruceta.clicked.connect(lambda: self.open_pdf('rosario', 'Crucetas_ETMA'))
		self.pushButton_cuchilla.clicked.connect(lambda: self.open_pdf('rosario', 'Cuchillas_Agro'))
		self.pushButton_forro.clicked.connect(lambda: self.open_pdf('rosario', 'FORRO_DE_EMBRAGUE'))
		self.pushButton_polea.clicked.connect(lambda: self.open_pdf('rosario', 'PoleasHF'))
		self.pushButton_cardan.clicked.connect(lambda: self.open_pdf('rosario', 'Repuestos_cardanicos'))
		self.pushButton_rotula.clicked.connect(lambda: self.open_pdf('rosario', 'Rotulas'))
		self.pushButton_varilla_rosario.clicked.connect(lambda: self.open_pdf('rosario', 'ROSCAS_ACME'))
		self.pushButton_soporte.clicked.connect(lambda: self.open_pdf('rosario', 'Soportes_FKD'))
		self.pushButton_termo.clicked.connect(lambda: self.open_pdf('rosario', 'Termoplasticos'))

		# Señales de comboboxes
		self.comboBox_most_used_hh.activated.connect(self.load_category)
		self.comboBox_most_used_etma.activated.connect(self.load_category)
		self.comboBox_most_used_camba.activated.connect(self.load_category)

		# Señales de lineedits
		self.lineEdit_search_hh.textEdited.connect(self.filter_products)
		self.lineEdit_search_etma.textEdited.connect(self.filter_products)
		self.lineEdit_search_camba.textEdited.connect(self.filter_products)

		# Configuraciones visuales varias
		self.format_headers() # Configuro headers de tablas
		self.apply_theme('light') # Tema claro por defecto
		self.showMaximized() # Ventana maximizada

		self.initialize()


	def initialize(self):
		"""Inicia el proceso en segundo plano con una barra de progreso."""

		# Vacio todo por si es una recarga
		self.empty_everything()

		# Creo el dialog de progreso
		self.progress_dialog = ProgressDialog(self)
		
		# Configuro el hilo y el worker
		self.thread = QThread()
		self.worker = Worker()
		self.worker.moveToThread(self.thread)

		# Conecto las señales del Worker al dialog
		self.worker.message_changed.connect(self.progress_dialog.label.setText)
		self.worker.progress_changed.connect(self.progress_dialog.progressBar.setValue)
		
		# Conecto señales de ciclo de vida e inicio
		self.thread.started.connect(self.worker.run)
		self.worker.finished.connect(self.on_worker_finished) # acá recibo los datos
		
		# Limpieza de memoria al terminar
		self.worker.finished.connect(self.thread.quit)
		self.worker.finished.connect(self.worker.deleteLater)
		self.thread.finished.connect(self.thread.deleteLater)

		# Inicio el hilo y muestro el dialog de forma modal
		self.thread.start()
		self.progress_dialog.exec()


	def on_worker_finished(self, final_data):
		"""Recibe los datos cuando el Worker termina."""

		# Cierro el dialog personalizado
		self.progress_dialog.accept()

		# Asigno los datos a la ventana principal
		self.all_products_hh = final_data['hh']['products']
		self.all_products_etma = final_data['etma']['products']
		self.all_products_camba = final_data['camba']['products']
		self.report = final_data['report']

		# Mapeo de marcas a sus correspondientes elementos
		bmap = {
			'hh': {
				'products': self.all_products_hh,
				'label': self.label_validity_date_hh,
				'table': self.tableWidget_search_hh,
				'combo': self.comboBox_most_used_hh,
				'most': MOST_USED_PRODUCTS_HH
			},
			'etma': {
				'products': self.all_products_etma,
				'label': self.label_validity_date_etma,
				'table': self.tableWidget_search_etma,
				'combo': self.comboBox_most_used_etma,
				'most': MOST_USED_PRODUCTS_ETMA
			},
			'camba': {
				'products': self.all_products_camba,
				'label': self.label_validity_date_camba,
				'table': self.tableWidget_search_camba,
				'combo': self.comboBox_most_used_camba,
				'most': MOST_USED_PRODUCTS_CAMBA
			}
		}

		for brand, elems in bmap.items():
			# Muestro fecha de validez de precios
			elems['label'].setText(final_data[brand]['date'])

			# Listo todos los productos
			self.list_products(elems['products'], elems['table'])

			# Listo los más usados
			self.load_more_used(elems['combo'], elems['products'], elems['most'])

		# Muestro el reporte si existe
		if self.report:
			QMessageBox.information(
				self,
				'Información de la carga',
				self.prepare_report()
			)


	def prepare_report(self):
		"""
		Lee el diccionario de reportes de PDF y genera los strings finales
		para sumarlos al reporte general.

		El diccionario de reporte puede terminar teniendo una estructura
		similar a esta, donde solamente se agrega algo cuando hubo un 
		problema:

		{
			'hh': {
				'excel': {
					'reason': 'no_link',
					'local_status': 'local_used'
				}
			}
			'camba': {
				'excel': {
					'reason': 'no_url',
					'local_status': 'local_used'
				},
				'pdfs': {
					'no_link': {
						'missing': ['05'],
						'local': ['01','04']
					},
					'no_download': {
						'missing': [],
						'local': ['10','11']
					}
				}
			}
		}
		"""
		
		maps = {
			'no_url': 'Sin URL configurada para',
			'no_access': 'Imposible acceder a',
			'no_link': 'No se encontró link',
			'no_download': 'No se pudo descargar',
			'local_used': 'Usando lista local',
			'local_missing': 'Lista local no encontrada',
			'local_error': 'Error al procesar lista'
		}
		
		brand_to_supplier = {
			'hh': ' Tienda del Cardan',
			'etma': ' Tienda del Cardan',
			'camba': ' Bulonera Camba',
			'rosario': ' Rosario Agro'
		}
		
		msg = ''

		for brand, data in self.report.items():
			msg += '<br><br>' if msg else ''
			msg += f'<b><u>{brand.upper()}</u></b>'
			
			if 'excel' in data:
				msg += '<br><b>Lista Excel</b>:'
				
				local_status = data['excel']['local_status']
				reason = data['excel']['reason']
				symbol = '⚠️' if local_status == 'local_used' else '❌'
				local_status_str = maps[local_status]

				if local_status == 'local_error':
					if reason is None:
						local_status_str += ' descargada'
					else:
						local_status_str += ' local'
					
				reason_str = '' if not reason else maps[reason]

				supplier_str = f'<i>{brand_to_supplier[brand]}</i>' if reason in ('no_url', 'no_access') else ''

				msg += (
					f' {symbol} {reason_str}{supplier_str}. '
					f'{local_status_str}.'
				)

			if 'pdfs' in data:
				msg += '<br><b>Listas PDF</b>:'

				for reason, info in data['pdfs'].items():

					sheets = info['missing'] + info['local']
					sheets.sort() # ordeno las hojas por número

					supplier_str = f'<i>{brand_to_supplier[brand]}</i>' if reason in ('no_url', 'no_access') else ''

					msg += f'<br>- Hojas {", ".join(sheets)}: {maps[reason]}{supplier_str}.'

					if info['local']:
						if set(info['local']) == set(sheets) and len(sheets) > 1:
							sheets_str = 'todas'
						else:
							sheets_str = ", ".join(info["local"])
						msg += f' ⚠️ Usando lista local para {sheets_str}.'

					if info['missing']:
						if set(info['missing']) == set(sheets) and len(sheets) > 1:
							sheets_str = 'ninguna'
						else:
							sheets_str = ", ".join(info["missing"])
						msg += f' ❌ Lista local no encontrada para {sheets_str}.'
						

		return msg



	############################################################################################
	# MÉTODOS QUE TOCAN LOS WIDGETS O SON DISPARADOS POR ACCIONES DEL USUARIO
	############################################################################################


	def apply_theme(self, theme):
		"""Aplica el tema y cambia los iconos en función del tema."""

		app = QApplication.instance()

		# Cambio esquema de color de la app
		if theme == 'dark':
			app.styleHints().setColorScheme(Qt.ColorScheme.Dark)
		else:
			app.styleHints().setColorScheme(Qt.ColorScheme.Light)

		# Actualizo ícono de botones
		self.pushButton_theme.setIcon(QIcon(f'resources/icons/icon_mode_{theme}.svg'))
		self.pushButton_config.setIcon(QIcon(f'resources/icons/icon_config_{theme}.svg'))
		self.pushButton_about.setIcon(QIcon(f'resources/icons/icon_about_{theme}.svg'))


	def change_theme(self):
		"""Invierte el tema actual."""

		current_theme = QApplication.instance().styleHints().colorScheme()
		new_theme ='light' if current_theme == Qt.ColorScheme.Dark else 'dark'
		self.apply_theme(new_theme)


	def format_headers(self):
		"""Distribuye el ancho de las columnas de todas las tablas."""

		tables = (
			self.tableWidget_search_hh, 
			self.tableWidget_defaults_hh,
			self.tableWidget_search_etma,
			self.tableWidget_defaults_etma,
			self.tableWidget_search_camba,
			self.tableWidget_defaults_camba
		)

		for table in tables:
			table.setColumnWidth(0, 110) # fijo
			table.setColumnWidth(1, 400) # fijo
			table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch) # ocupa el resto
			table.setColumnWidth(3, 160) # fijo


	def load_more_used(self, combo_box, all_products, most_used_products):
		"""Carga los productos más usados."""

		# Muestro texto por defecto en el combo box
		combo_box.setPlaceholderText('Seleccione una categoría...')

		# Cargo categorías y sus productos por detrás
		for category, products_in_category in most_used_products.items():
			products = []
			for product_code, product_description in products_in_category.items():
				# Busco el producto dentro de todos los productos
				for product in all_products:
					# if product_code.startswith('CR1024') and product_code != product['code']: # Evita duplicados en este caso particular
					# 	continue
					if product_code == product['code']:
						products.append({
								'code': product['code'],
								'subcategory': product['subcategory'],
								'description': product_description,
								'price': product['price']
							}
						)
			combo_box.addItem(category, products)

		# Establezco que no haya uno seleccionado
		combo_box.setCurrentIndex(-1)


	def load_category(self):
		"""Lista los productos mas usados de la categoría seleccionada."""

		# Determino si se seleccionó categoría en HH, ETMA, o CAMBA, y asigno variables
		sender = self.sender()
		if sender is self.comboBox_most_used_hh:
			table_widget = self.tableWidget_defaults_hh
			combo_box = self.comboBox_most_used_hh
		elif sender is self.comboBox_most_used_etma:
			table_widget = self.tableWidget_defaults_etma
			combo_box = self.comboBox_most_used_etma
		else:
			table_widget = self.tableWidget_defaults_camba
			combo_box = self.comboBox_most_used_camba

		# Vacio la tabla y listo los productos
		table_widget.setRowCount(0)
		self.list_products(combo_box.currentData(), table_widget)


	def filter_products(self, query):
		"""Filtra la lista de productos al escribir en el buscador."""

		sender = self.sender()

		# Determino si se buscó en HH o en ETMA, y asigno variables
		if sender is self.lineEdit_search_hh:
			table_widget = self.tableWidget_search_hh
			all_products = self.all_products_hh
		elif sender is self.lineEdit_search_etma:
			table_widget = self.tableWidget_search_etma
			all_products = self.all_products_etma
		else:
			table_widget = self.tableWidget_search_camba
			all_products = self.all_products_camba

		# Evito lógica innecesaria si no se cargaron productos en la marca
		if not all_products:
			return

		# Normalizo el texto del filtro
		query = ' '.join(query.split()).lower()

		# Busco productos coincidentes
		if query:
			filtered_products = []
			for product in all_products:
				if query in product['code'].lower() or query in product['subcategory'].lower() or query in product['description'].lower():
					filtered_products.append(product)
			self.list_products(filtered_products, table_widget)
		else: # Si no hay nada escrito, muestro todos los productos
			self.list_products(all_products, table_widget)


	def list_products(self, products, table_widget):
		"""Lista los productos en la tabla correspondiente."""

		# Vacio la tabla y cargo los productos
		table_widget.setRowCount(0)
		for product in products:
			row = table_widget.rowCount()
			table_widget.insertRow(row)

			code_item = QTableWidgetItem(product['code'])
			table_widget.setItem(row, 0, code_item)

			subcat_item = QTableWidgetItem(product['subcategory'])
			if table_widget is self.tableWidget_search_camba:
				font = subcat_item.font()
				font.setPointSize(9) # tamaño deseado
				subcat_item.setFont(font)
			table_widget.setItem(row, 1, subcat_item)

			descr_item = QTableWidgetItem(product['description'])
			table_widget.setItem(row, 2, descr_item)

			price_item = QTableWidgetItem(product['price'])
			price_item.setFont(QFont('Consolas', 12))
			table_widget.setItem(row, 3, price_item)

		# Muestro el número de productos listado
		search_tables = {
			self.tableWidget_search_hh: self.label_search_hh,
			self.tableWidget_search_etma: self.label_search_etma,
			self.tableWidget_search_camba: self.label_search_camba,
			self.tableWidget_defaults_hh: self.label_most_used_hh,
			self.tableWidget_defaults_etma: self.label_most_used_etma,
			self.tableWidget_defaults_camba: self.label_most_used_camba,
		}
		quantity = len(products)
		s = '' if quantity == 1 else 's'
		search_tables[table_widget].setText(f'{quantity} producto{s} encontrado{s}')


	def empty_everything(self):
		"""Vacia la interfaz para la recarga de listas."""

		# Junto widgets que usan clear()
		widgets = {
			self.lineEdit_search_hh,
			self.lineEdit_search_etma,
			self.lineEdit_search_camba,
			self.label_search_hh,
			self.label_search_etma,
			self.label_search_camba,
			self.label_most_used_hh,
			self.label_most_used_etma,
			self.label_most_used_camba,
			self.label_validity_date_hh,
			self.label_validity_date_etma,
			self.label_validity_date_camba,
			self.comboBox_most_used_hh,
			self.comboBox_most_used_etma,
			self.comboBox_most_used_camba
		}

		tables = (
			self.tableWidget_search_hh,
			self.tableWidget_defaults_hh,
			self.tableWidget_search_etma,
			self.tableWidget_defaults_etma,
			self.tableWidget_search_camba,
			self.tableWidget_defaults_camba
		)

		for widget in widgets:
			widget.clear() 

		for table in tables:
			table.setRowCount(0)


	def open_pdf(self, brand, identifier, page_number=1):
		"""
		Busca el PDF correspondiente y lo abre, en orden de disponibilidad, con:
		* Navegador predeterminado, en la página indicada.
		* Visor PDF predeterminado del sistema, sin poder indicar la página.
		
		Parámetros:
		- brand: 'camba' o 'rosario'
		- identifier: número de hoja ('02') para Camba, o nombre ('Cuchillas_Jardin') para Rosario.
		- page_number: La página donde se quiere arrancar (por defecto 1).
		"""

		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand
		pdf_file_path = None

		if not base_path.exists():
			supplier = 'Bulonera Camba' if brand == 'camba' else 'Rosario Agro'
			QMessageBox.warning(
				self, 
				'Carpeta no encontrada', 
				f'No existe la carpeta de listas para {supplier}.'
			)
			return

		# Busco la ruta del PDF respetando mayúsculas/minúsculas
		if brand == 'camba':
			pdf_files = list(base_path.glob(f'Hoja{identifier}*.pdf'))
			if pdf_files:
				pdf_file_path = pdf_files[0]
		elif brand == 'rosario':
			exact_path = base_path / f'{identifier}.pdf'
			if exact_path.exists():
				pdf_file_path = exact_path

		# # Busco la ruta del PDF ignorando mayúsculas/minúsculas
		# if brand == 'camba':
		# 	target_prefix = f'hoja{identifier}'.lower()
		# 	for pdf_file in base_path.glob('*.pdf'):
		# 		if pdf_file.name.lower().startswith(target_prefix):
		# 			pdf_file_path = pdf_file
		# 			break  # Encontré el archivo, salgo del ciclo
					
		# elif brand == 'rosario':
		# 	target_name = f'{identifier}.pdf'.lower()
		# 	for pdf_file in base_path.glob('*.pdf'):
		# 		if pdf_file.name.lower() == target_name:
		# 			pdf_file_path = pdf_file
		# 			break  # Encontré el archivo, salgo del ciclo

		# Si encontré el archivo, lo abro
		if pdf_file_path:
			try:
				# Intento con el navegador predeterminado
				default_browser_exe = self.get_default_browser_exe()
				if default_browser_exe:
					# Formateo la ruta de Windows a un formato URI que el navegador entienda
					pdf_uri = f'file:///{str(pdf_file_path).replace(os.sep, "/")}#page={page_number}'
					subprocess.Popen([default_browser_exe, pdf_uri])
				else:
					# Como último recurso, intento con el lector de PDF predeterminado
					url = QUrl.fromLocalFile(str(pdf_file_path))
					QDesktopServices.openUrl(url)
			except Exception as e:
				QMessageBox.critical(
					self, 
					'Error', 
					f'No se pudo abrir el PDF:\n{str(e)}'
				)
		else:
			filename = f'Hoja {identifier}' if brand == 'camba' else identifier
			QMessageBox.warning(
				self, 
				'Archivo no encontrado', 
				f'No se pudo encontrar el PDF local para: <b>{filename}</b>.'
			)


	def get_default_browser_exe(self):
		"""
		Consulta el Registro de Windows para obtener el ejecutable del navegador 
		web predeterminado.
		"""

		try:
			# Busco qué programa maneja los links de internet (HTTP)
			reg_url = r'Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice'
			with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_url) as key:
				prog_id = winreg.QueryValueEx(key, 'ProgId')[0]

			# Busco la ruta del ejecutable para ese programa
			reg_cmd = rf'{prog_id}\shell\open\command'
			with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reg_cmd) as key:
				command = winreg.QueryValueEx(key, '')[0]

			# Limpio la ruta para tener la ruta "pura" del ejecutable, sin parámetros
			if command.startswith('"'):
				return command.split('"')[1]
			else:
				return command.split(' ')[0]
		except Exception:
			return None


	def open_config(self):
		"""Abre un dialog para editar la configuración."""

		dialog = ConfigurationDialog(self)
		dialog.exec()

		# Verifico si recargar
		if dialog.new_supplier_urls:
			self.initialize()


	def open_about(self):
		"""Abre un dialog de Acerca de."""

		dialog = AboutDialog(self)
		dialog.exec()



class ConfigurationDialog(QDialog):
	def __init__(self, parent=None):
		super().__init__(parent)

		# Cargo la UI
		uic.loadUi('ui/config.ui', self)

		self.load_config()

		# Defino variables
		self.new_supplier_urls = False # Flag para recargar al cerrar dialog

		# Conecto señales
		self.pushButton_ok.clicked.connect(self.save_config)
		self.pushButton_cancel.clicked.connect(self.close)


	def load_config(self):
		self.lineEdit_url_tdc.setText(SETTINGS.value('supplier_urls/tdc', '', type=str))
		self.lineEdit_url_camba.setText(SETTINGS.value('supplier_urls/camba', '', type=str))


	def save_config(self):
		SETTINGS.setValue('supplier_urls/tdc', self.lineEdit_url_tdc.text())
		SETTINGS.setValue('supplier_urls/camba', self.lineEdit_url_camba.text())
		self.new_supplier_urls = True # Para recargar al cerrar configuración
		self.close()



class AboutDialog(QDialog):
	def __init__(self, parent=None):
		super().__init__(parent)

		# Cargo la UI
		uic.loadUi('ui/about.ui', self)



class ProgressDialog(QDialog):
	def __init__(self, parent=None):
		super().__init__(parent)

		# Cargo la UI
		uic.loadUi('ui/progress.ui', self)

		# Quito el botón de cerrar de la ventana para que el usuario no lo interrumpa
		self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, False)

		# Centro el dialog en la pantalla
		self.center()


	def center(self):
		"""Centra el dialog en la pantalla"""

		# Rectángulo que define la geometría del dialog
		qr = self.frameGeometry()

		# Punto central de la pantalla
		cp = self.screen().availableGeometry().center()

		# Muevo el centro del rectángulo del dialog al centro de la pantalla
		qr.moveCenter(cp)

		# Muevo el dialog a la posición de la esquina superior izquierda del rectángulo ya centrado
		self.move(qr.topLeft())


# Initialize the app
if __name__ == "__main__":
	app = QApplication(sys.argv)
	app.setStyle('Fusion')
	window = MainWindow()
	window.show()
	sys.exit(app.exec())