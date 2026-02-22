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
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import QMainWindow, QApplication, QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont
from PyQt6 import uic
import sys
from utils import MOST_USED_PRODUCTS_HH, MOST_USED_PRODUCTS_ETMA, MOST_USED_PRODUCTS_CAMBA, RUBROS_CAMBA, URLS_ROSARIO_AGRO, SETTINGS



class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo la UI
		uic.loadUi('ui/app.ui', self)

		# Inicializo variables
		self.all_products_hh = None
		self.all_products_etma = None
		self.all_products_camba = None
		self.camba_last_date = None

		# Conecto señales
		self.pushButton_theme.clicked.connect(self.change_theme)
		self.pushButton_config.clicked.connect(self.open_config)
		self.pushButton_about.clicked.connect(self.open_about)
		self.comboBox_most_used_hh.activated.connect(self.load_category)
		self.comboBox_most_used_etma.activated.connect(self.load_category)
		self.lineEdit_search_hh.textEdited.connect(self.filter_products)
		self.lineEdit_search_etma.textEdited.connect(self.filter_products)

		# Configuro headers de tablas
		self.format_headers()

		# Aplico tema claro por defecto
		self.apply_theme('light')

		# self.showMaximized() # Abro la ventana maximizada

		self.initialize()



	############################################################################################
	# PROCESAMIENTO INICIAL DE LISTAS
	############################################################################################

	# CÓDIGO PRINCIPAL PARA LISTAS EXCEL
	# ------------------------------------------------------------------------------------------

	def initialize(self):
		"""Método principal para gestionar la descarga y procesamiento de las listas."""

		suppliers = {
			'tdc': {
				'name': 'Tienda del Cardan',
				'brands': ('hh', 'etma')
			},
			'camba': {
				'name': 'Bulonera Camba',
				'brands': ('camba',)
			}
		}
		
		for supplier, config in suppliers.items():
			
			# Recupero la URL del proveedor
			price_lists_url = self.get_url_from_settings(supplier)
			if not price_lists_url:
				QMessageBox.warning(
					self,
					'Configuración faltante',
					f'No hay URL configurada para {config["name"]}'
				)
				self.try_local_lists(config['brands'])
				continue

			# Obtengo el HTML de esa URL
			try:
				html = self.download_html(price_lists_url)
			except Exception as e:
				QMessageBox.warning(
					self,
					'Error de conexión',
					f'No se pudo acceder a la página de {config["name"]}:\n{e}'
				)
				self.try_local_lists(config['brands'])
				continue

			# Tengo HTML válido: proceso cada marca
			for brand in config['brands']:
				self.process_brand(html, brand)


	# PROCESAMIENTO POR MARCA
	# ------------------------------------------------------------------------------------------

	def process_brand(self, html, brand):
		"""Busca la URL de la lista en el html, la descarga y la procesa."""

		# Busco link de la lista
		list_url = self.get_list_url_from_html(html, brand)
		if not list_url:
			QMessageBox.warning(
				self,
				'Advertencia',
				f'No se encontró link para la lista de {brand.upper()}')
			self.try_local_list(brand)
			return

		# Descargo la lista
		try:
			excel_file_path = self.download_excel_file(list_url, brand)
		except Exception as e:
			QMessageBox.warning(
				self,
				'Advertencia',
				f'No se pudo descargar la lista de {brand.upper()}:\n{e}')
			self.try_local_list(brand)
			return

		# Proceso excel descargado
		try:
			self.process_excel(excel_file_path, brand)
		except Exception as e:
			QMessageBox.critical(
				self,
				'Error',
				f'Error procesando lista descargada de {brand.upper()}:\n{e}'
			)


	# FALLBACKS (hubo error y se debe buscar lista local descargada previamente)
	# ------------------------------------------------------------------------------------------

	def try_local_lists(self, brands):
		"""
		Fallback general llamado cuando:
		  * No hay URL del proveedor.
		  * No se pudo obtener HTML de la URL.
		Llama al fallback de marca por cada marca del proveedor.
		"""
		for brand in brands:
			self.try_local_list(brand)


	def try_local_list(self, brand):
		"""
		Fallback por marca llamado cuando hay:
		  * No se encontró el link de la lista en el HTML.
		  * No se pudo descargar el excel.
		Comprueba si existe una excel local previamente descargado y lo procesa.
		"""

		excel_file_path = self.search_existing_excel(brand)
		if not excel_file_path:
			QMessageBox.information(
				self,
				'Información',
				f'No hay lista previamente descargada para {brand.upper()}'
			)
			return

		# Proceso excel previo
		try:
			self.process_excel(excel_file_path, brand)
		except Exception as e:
			QMessageBox.critical(
				self,
				'Error',
				f'Error procesando lista previa de {brand.upper()}:\n{e}'
			)


	#  AUXILIARES
	# ------------------------------------------------------------------------------------------

	def get_url_from_settings(self, supplier):
		return SETTINGS.value(f'price_lists_urls/{supplier}', '', type=str)


	def download_html(self, url):
		response = requests.get(url, timeout=10)
		response.raise_for_status()
		return response.text


	def get_list_url_from_html(self, html, brand):
		"""Obtiene del sitio del proveedor el link actual de la lista correspondiente."""

		soup = bs4.BeautifulSoup(html, 'html.parser')

		if brand in ('hh', 'etma'):
			# Busco el título correcto
			h1 = soup.find(
				'h1', 
				string=lambda s: s and s.strip().upper() in (f'LISTA DE PRECIO {brand.upper()}', f'LISTA DE PRECIOS {brand.upper()}')
			)
			if not h1:
				return None

			# Subo al bloque contenedor
			bloque = h1.find_parent('div', class_='widget-span')
			if not bloque:
				return None

			# Busco el link dentro del bloque
			url = bloque.find_next('a', href=True)
		else: # camba
			# Busco el título correcto
			h2 = soup.find(
				'h2', 
				string=lambda s: s and 'Lista de precios formato sabana' in s.strip()
			)
			if not h2:
				return None

			# Busco el link
			url = h2.find_parent('a')

			# Aprovecho el html y extraigo la fecha de actualización
			a = soup.find(
				'a',
				href=True,
				string=lambda s: s and 'lista indice' in s.strip().lower()
			)

			if a:
				match = re.search(r'\d{2}/\d{2}/\d{4}', a.get_text())
				if match:
					self.camba_last_date = match.group()

		return url['href'] if url else None


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


	def search_existing_excel(self, brand):
		"""Busca un excel previo en la carpeta de la marca, y si existe, retorna su ruta."""
		
		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand

		if not base_path.exists():
			return None

		excel_files = list(base_path.glob('*.xlsx'))

		return excel_files[0] if excel_files else None


	def process_excel(self, excel_file_path, brand):
		"""Lee el excel y carga los productos en la interfaz."""

		# Mapeo de marcas a su correspondiente widget en la UI
		bmap = {
			'hh': {
				'label': self.label_validity_date_hh,
				'prods': self.all_products_hh,
				'table': self.tableWidget_search_hh,
				'combo': self.comboBox_most_used_hh,
				'most': MOST_USED_PRODUCTS_HH
			},
			'etma': {
				'label': self.label_validity_date_etma,
				'prods': self.all_products_etma,
				'table': self.tableWidget_search_etma,
				'combo': self.comboBox_most_used_etma,
				'most': MOST_USED_PRODUCTS_ETMA
			},
			'camba': {
				'label': self.label_validity_date_camba,
				'prods': self.all_products_camba,
				'table': self.tableWidget_search_camba,
				'combo': self.comboBox_most_used_camba,
				'most': MOST_USED_PRODUCTS_CAMBA
			}
		}

		# Creo workbook y extraigo la hoja de productos
		wb = openpyxl.load_workbook(excel_file_path)
		sheet = wb[wb.sheetnames[0]]

		# Busco letras de columnas de producto
		header_cols = self.search_header_cols(sheet)

		# Busco número de fila de primer producto
		first_row = self.search_first_row(sheet, header_cols['price_col'])

		# Muestro fecha de validez de precios de cada hoja
		self.show_validity_date(sheet, bmap[brand]['label'])

		# Paso los productos a un diccionario
		bmap[brand]['prods'] = self.obtain_products(sheet, first_row, header_cols)

		# Listo todos los productos
		self.list_products(bmap[brand]['prods'], bmap[brand]['table'])

		# Listo los más usados
		# self.load_more_used(bmap[brand]['combo'], bmap[brand]['prods'], bmap[brand]['most'])


	def search_header_cols(self, sheet):
		"""Retorna un diccionario con la posición (letra) de cada columna."""

		for row in sheet['A1':'E20']:
			for cell in row:
				if str(cell.value).strip().lower() in ('codigo', 'código', 'cod', 'cód'):
					header_row = cell.row
					code_col = get_column_letter(cell.column)

					for header_cell in sheet[header_row]:
						print(get_column_letter(header_cell.column))
						value = str(header_cell.value).strip().lower()
						if value in ('subrubro', 'sub rubro', 'rubro'):
							subcategory_col = get_column_letter(header_cell.column)
						elif value in ('none', 'descripción', 'descripcion', 'desc', 'articulo', 'artículo'):
							description_col = get_column_letter(header_cell.column)
						elif value in ('precio + iva', 'precio'):
							price_col = get_column_letter(header_cell.column)

					return {
						'code_col': code_col,
						'subcategory_col': subcategory_col if subcategory_col else None, # camba no tiene
						'description_col': description_col,
						'price_col': price_col
					}


	def search_first_row(self, sheet, price_col):
		"""Retorna la fila donde comienzan los productos."""

		for cell in sheet[price_col]:
			if cell.value is not None:
				try:
					float(str(cell.value).replace('.', '').replace(',', '.'))
					return cell.row
				except (ValueError, TypeError):
					continue


	def obtain_products(self, sheet, first_row, header_cols):
		"""Crea lista de diccionarios de productos para filtrar."""

		products = []
		for row in range(first_row, sheet.max_row + 1):
			if self.is_valid_row(sheet, row, header_cols):
				price = sheet[header_cols['price_col'] + str(row)].value
				if isinstance(price, float):
					# Convierto float a formato de moneda argentina
					price = f'{price:,}'.replace('.', '_').replace(',', '.').replace('_', ',')
				product = {
					'code': sheet[header_cols['code_col'] + str(row)].value,
					'subcategory': sheet[header_cols['subcategory_col'] + str(row)].value if header_cols['subcategory_col'] else '',
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



	############################################################################################
	# MÉTODOS QUE MODIFICAN LA INTERFAZ O SON DISPARADOS POR USUARIO
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

		# Fijo ancho de "CÓDIGO", "SUBCATEGORÍA" y "PRECIO + IVA" y hago que "DESCRIPCIÓN" ocupe el resto
		for table in tables:
			table.setColumnWidth(0, 150)
			table.setColumnWidth(1, 250)
			table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
			table.setColumnWidth(3, 200)


	def show_validity_date(self, sheet, label):
		"""Muestra la fecha de validez de precios presente en la hoja."""

		if label is self.label_validity_date_camba:
			date = '¿?' if self.camba_last_date is None else self.camba_last_date
			label.setText(f'📆 Precios válidos para el: {date}')
		else:
			for row in sheet['A1':'E20']:
				for cell in row:
					value = str(cell.value)
					if re.findall(r'\d{1,2}/\d{1,2}/\d{2,4}', value) and ('valid' in value or 'válid' in value):
						label.setText('📆 ' + value.replace('validos', 'válidos'))
						return


	def load_more_used(self, combo_box, all_products, most_used_products):
		"""Carga los productos más usados."""

		# Muestro texto por defecto en el combo box
		combo_box.setPlaceholderText('Seleccione una categoría...')

		# Cargo categorías y sus productos por detrás
		for category, products_in_category in most_used_products.items():
			products = []
			for product_code, product_description in products_in_category.items():
				for product in all_products:
					if product_code.startswith('CR1024') and product_code != product['code']: # Evita duplicados en este caso particular
						continue
					if product_code in product['code']:
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

		# Determino si se seleccionó categoría en HH o en ETMA, y asigno variables
		if self.sender() is self.comboBox_most_used_hh:
			table_widget = self.tableWidget_defaults_hh
			combo_box = self.comboBox_most_used_hh
		else:
			table_widget = self.tableWidget_defaults_etma
			combo_box = self.comboBox_most_used_etma

		# Vacio la tabla y listo los productos
		table_widget.setRowCount(0)
		self.list_products(combo_box.currentData(), table_widget)


	def filter_products(self, query):
		"""Filtra la lista de productos al escribir en el buscador."""

		# Determino si se buscó en HH o en ETMA, y asigno variables
		if self.sender() is self.lineEdit_search_hh:
			table_widget = self.tableWidget_search_hh
			all_products = self.all_products_hh
		else:
			table_widget = self.tableWidget_search_etma
			all_products = self.all_products_etma

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
			self.tableWidget_search_camba: self.label_search_camba
		}
		if table_widget in search_tables:
			quantity = len(products)
			s = '' if quantity == 1 else 's'
			search_tables[table_widget].setText(f'{quantity} producto{s} encontrado{s}')


	def open_config(self):
		"""Abre un dialogo para editar la configuración."""

		dialog = ConfigurationDialog(self)
		dialog.exec()

		# Verifico si recargar
		if dialog.new_price_lists_url:
			pass
			# self.cargarClientes()


	def open_about(self):
		"""Abre un diálogo de Acerca de."""

		dialog = AboutDialog(self)
		dialog.exec()



class ConfigurationDialog(QDialog):
	def __init__(self, parent=None):
		super().__init__(parent)

		# Cargo la UI
		uic.loadUi('ui/config.ui', self)

		self.load_config()

		# Defino variables
		self.new_price_lists_url = False # Flag para recargar al cerrar dialog

		# Conecto señales
		self.pushButton_ok.clicked.connect(self.save_config)
		self.pushButton_cancel.clicked.connect(self.close)


	def load_config(self):
		self.lineEdit_url_tdc.setText(SETTINGS.value('price_lists_urls/tdc', '', type=str))
		self.lineEdit_url_camba.setText(SETTINGS.value('price_lists_urls/camba', '', type=str))


	def save_config(self):
		SETTINGS.setValue('price_lists_urls/tdc', self.lineEdit_url_tdc.text())
		SETTINGS.setValue('price_lists_urls/camba', self.lineEdit_url_camba.text())
		self.new_price_lists_url = True # Para recargar al cerrar configuración
		self.close()



class AboutDialog(QDialog):
	def __init__(self, parent=None):
		super().__init__(parent)

		# Cargo la UI
		uic.loadUi('ui/about.ui', self)



# Initialize the app
if __name__ == "__main__":
	app = QApplication(sys.argv)
	app.setStyle('Fusion')
	window = MainWindow()
	window.show()
	sys.exit(app.exec())