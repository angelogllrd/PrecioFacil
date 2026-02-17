import openpyxl
import requests
import bs4
import tempfile
from pathlib import Path
import os
from urllib.parse import urlparse, unquote
import re
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import QMainWindow, QApplication, QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QMessageBox
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QIcon, QFont
from PyQt6 import uic
import sys


SETTINGS = QSettings('COMET', 'PrecioFacil')

MOST_USED_PRODUCTS_HH = {
	'CRUCETAS K5-18': {
		'5412': 'HORQUILLA CIEGA CON BASE Ø 58 mm',
		'5415': 'HORQUILLA CON AGUJERO REDONDO Ø 30 mm',
		'5419': 'HORQUILLA CON AGUJERO CUADRADO 32 mm',
		'5418': 'HORQUILLA CON 6 ESTRIAS Ø 35 mm',
		'5420': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOTON',
		'5417': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOLITAS',
		'5414': 'HORQUILLA CON 6 ESTRIAS Ø 45',
		'2278': 'CRUCETA K5-18 CON RODILLO',
		'5409': 'CRUCETA K5-18 CON BUJE'
	},
	'CRUCETA K5-21': {
		'5441': 'HORQUILLA CIEGA CON BASE Ø 58 mm',
		'5447': 'HORQUILLA CON AGUJERO REDONDO Ø 30 mm',
		'5431': 'HORQUILLA CON AGUJERO CUADRADO 25,4 mm',
		'5446': 'HORQUILLA CON AGUJERO CUADRADO 32 mm',
		'5443': 'HORQUILLA CON 6 ESTRIAS Ø 35 mm',
		'5456': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOTON',
		'5442': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOLITAS',
		'6654': 'CRUCETA K5-21 CON RODILLO',
		'5449': 'CRUCETA K5-21 CON BUJE'
	},
	'CRUCETA K5-26': {
		'5491': 'HORQUILLA CON AGUJERO Ø 31,8 mm Y CHAVETERO',
		'5460': 'HORQUILLA CON AGUJERO CUADRADO 22 mm',
		'5478': 'HORQUILLA CON AGUJERO CUADRADO 25,4 mm',
		'5465': 'HORQUILLA CON 6 ESTRIAS Ø 35 mm',
		'5466': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOLITAS',
		'5490': 'HORQUILLA CON 6 ESTRIAS CON SEGURO A BOTON',
		'3775': 'CRUCETA K5-L1 CON RODILLOS',
		'5475': 'CRUCETA K5-L1 A BUJE'
	},
	'CRUCETA RYCSA': {
		'5376': 'HORQUILLA CIEGA BASE Ø 41 mm',
		'5377': 'HORQUILLA CON AGUJERO Ø 19 mm',
		'5378': 'HORQUILLA CON AGUJERO Ø 22 mm',
		'5349': 'HORQUILLA CON AGUJERO Ø 22 mm Y CHAVETERO',
		'5379': 'HORQUILLA CON AGUJERO Ø 19 mm Y ORIFICIO P/AJUSTE',
		'5358': 'HORQUILLA CON AGUJERO Ø 25,4 mm Y AGUJERO PASANTE Ø 8 mm',
		'5381': 'CRUCETA RYCSA A BUJE'
	},
	'ACCESORIOS Y COMPLEMENTOS PARA ARMADO DE CARDANES': {
		'5450': 'MANCHON CON 6 ESTRIAS Ø 35 mm. LARGO 100 mm',
		'5452': 'MANCHON CON AGUJERO CUADRADO 25,4 mm. LARGO 100 mm',
		'5454': 'MANCHON CON AGUJERO CUADRADO 32 mm. LARGO 100 mm',
		'5455': 'MANCHON CON AGUJERO CUADRADO 32 mm. LARGO 150 mm',
		'5444': 'MANCHON REDUCTOR 6 ESTRIAS Ø 45 mm A EJE 6 ESTRIAS Ø 35 mm',
		'5476': 'TACITA-SEGURO-BOLITAS-REPARACION DE HORQUILLAS A BOLITAS',
		'5406': 'PERNO Y RESORTE-REPARACION DE HORQUILLAS A BOTON'
	}
}

MOST_USED_PRODUCTS_ETMA = {
	# Según https://zanini.com.ar/categoria-producto/transmisiones-cardanicas/crucetas/
	'CRUCETA K-530': {
		'CR1047': 'CRUCETA PARA HORQUILLA K-530 A PALITOS',
	},
	'CRUCETA K-526': {
		'CR1004AC': 'CRUCETA PARA HORQUILLA K-526 A PALILLO',
		'CR1004B': 'CRUCETA PARA HORQUILLA K-526 A BUJE',
	},
	'CRUCETA K-521': {
		'CR1001ACXL': 'CRUCETA PARA HORQUILLA K-521 A PALILLO',
		'CR1001BXL': 'CRUCETA PARA HORQUILLA K-521 A BUJE',
	},
	'CRUCETA K-518': {
		'CR1003AC': 'CRUCETA PARA HORQUILLA K-518 A PALILLO',
		'CR1003B': 'CRUCETA PARA HORQUILLA K-518 A BUJE',
	},
	'CRUCETA K-514': {
		'CR1024': 'CRUCETA PARA HORQUILLA K-514 A PALILLO',
		'CR1024B': 'CRUCETA PARA HORQUILLA K-514 A BUJE'
	}
}



class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo la UI
		uic.loadUi('ui/app.ui', self)

		# Inicializo variables
		self.all_products_hh = None
		self.all_products_etma = None

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





	def initialize(self):
		"""."""

		brands = ('etma', 'hh')
		
		# Recupero la URL de las listas de precios de TDC
		price_lists_url_tdc = self.get_url_from_settings('tdc')
		if not price_lists_url_tdc:
			QMessageBox.warning(
				self,
				'Error',
				'No hay URL configurada para Tienda del Cardan'
			)
			self.try_local_lists(brands)
			return

		# Obtengo el HTML de esa URL
		try:
			html = self.download_html(price_lists_url_tdc)
		except Exception as e:
			QMessageBox.warning(
				self,
				'Error',
				f'No se pudo acceder a la página de Tienda Del Cardan:\n{e}')
			self.try_local_lists(brands)
			return

		# Tengo HTML válido: proceso cada marca
		for brand in brands:
			self.process_brand(html, brand)


	def process_brand(self, html, brand):
		"""."""

		# Busco link del excel
		excel_url = self.get_excel_url_tdc(html, brand)
		if not excel_url:
			QMessageBox.warning(
				self,
				'Advertencia',
				f'No se encontró link para la lista de {brand.upper()}')
			self.try_local_list(brand)
			return

		# Descargo el excel
		try:
			excel_file_path = self.download_excel_file_tdc(excel_url, brand)
		except Exception as e:
			QMessageBox.warning(
				self,
				'Advertencia',
				f'No se pudo descargar la lista de {brand.upper()}:\n{e}')
			self.try_local_list(brand)
			return

		# Proceso excel descargado
		try:
			self.process_excel_tdc(excel_file_path, brand)
		except Exception as e:
			QMessageBox.critical(
				self,
				'Error',
				f'Error procesando lista descargada de {brand.upper()}:\n{e}'
			)


	def get_excel_url_tdc(self, html, brand):
		"""Obtiene en TDC el link actual del excel correspondiente."""

		soup = bs4.BeautifulSoup(html, 'html.parser')

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

		return url['href'] if url else None


	def download_excel_file_tdc(self, url, brand):
		"""Descarga el excel en la carpeta correspondiente."""

		# Obtengo la ruta de la carpeta de destino
		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand
		base_path.mkdir(parents=True, exist_ok=True)

		# Obtengo el nombre original del archivo descargado
		excel_original_name = os.path.basename(urlparse(url).path)
		excel_original_name = unquote(excel_original_name) # Quito los %20 (espacios)

		# Obtengo la ruta completa
		excel_file_path = base_path / excel_original_name

		response = requests.get(url, timeout=10)
		response.raise_for_status()

		# Si hay un excel previo, lo borro (no quiero que se acumulen, siempre solo 1)
		for old_excel_file in base_path.glob('*.xlsx'):
			old_excel_file.unlink()

		# Guardo el nuevo excel
		with open(excel_file_path, 'wb') as f:
			f.write(response.content)
		
		return excel_file_path


	def process_excel_tdc(self, excel_file_path, brand):
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
		self.load_more_used(bmap[brand]['combo'], bmap[brand]['prods'], bmap[brand]['most'])


	def try_local_lists(self, brands):
		for brand in brands:
			self.try_local_list(brand)


	def try_local_list(self, brand):
		"""."""

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
			self.process_excel_tdc(excel_file_path, brand)
		except Exception as e:
			QMessageBox.critical(
				self,
				'Error',
				f'Error procesando lista previa de {brand.upper()}:\n{e}'
			)


	def search_existing_excel(self, brand):
		base_path = Path(os.getenv('APPDATA')) / 'PrecioFacil' / 'listas' / brand

		if not base_path.exists():
			return None

		excel_files = list(base_path.glob('*.xlsx'))

		return excel_files[0] if excel_files else None


	def get_url_from_settings(self, brand):
		return SETTINGS.value(f'price_lists/{brand}', '', type=str)


	def download_html(self, url):
		response = requests.get(url, timeout=10)
		response.raise_for_status()
		return response.text


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
			self.tableWidget_defaults_etma
		)

		# Fijo ancho de "CÓDIGO", "SUBCATEGORÍA" y "PRECIO + IVA" y hago que "DESCRIPCIÓN" ocupe el resto
		for table in tables:
			table.setColumnWidth(0, 150)
			table.setColumnWidth(1, 250)
			table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
			table.setColumnWidth(3, 200)


	def search_header_cols(self, sheet):
		"""Retorna un diccionario con la posición (letra) de cada columna."""

		for row in sheet['A1':'E20']:
			for cell in row:
				if str(cell.value).strip().lower() in ('codigo', 'código', 'cod', 'cód'):
					header_row = cell.row
					code_col = get_column_letter(cell.column)

					for header_cell in sheet[header_row]:
						value = str(header_cell.value).strip().lower()
						if value in ('subrubro', 'sub rubro', 'rubro'):
							subcategory_col = get_column_letter(header_cell.column)
						elif value in ('none', 'descripción', 'descripcion', 'desc'):
							description_col = get_column_letter(header_cell.column)
						elif value in ('precio + iva', 'precio'):
							price_col = get_column_letter(header_cell.column)

					return {
						'code_col': code_col,
						'subcategory_col': subcategory_col,
						'description_col': description_col,
						'price_col': price_col
					}


	def search_first_row(self, sheet, price_col):
		"""Retorna la fila donde comienzan los productos."""

		for cell in sheet[price_col]:
			if cell.value is not None:
				try:
					float(cell.value.replace('.', '').replace(',', '.'))
					return cell.row
				except (ValueError, TypeError):
					continue


	def show_validity_date(self, sheet, label):
		"""Muestra la fecha de validez de precios presente en la hoja."""

		for row in sheet['A1':'E20']:
			for cell in row:
				value = str(cell.value)
				if re.findall(r'\d{1,2}/\d{1,2}/\d{2,4}', value) and ('valid' in value or 'válid' in value):
					label.setText('📆 ' + value)
					return


	def obtain_products(self, sheet, first_row, header_cols):
		"""Creo lista de diccionarios de productos para filtrar."""

		products = []
		for row in range(first_row, sheet.max_row + 1):
			if self.is_valid_row(sheet, row, header_cols):
				product = {
					'code': sheet[header_cols['code_col'] + str(row)].value,
					'subcategory': sheet[header_cols['subcategory_col'] + str(row)].value,
					'description': sheet[header_cols['description_col'] + str(row)].value,
					'price': '$ ' + sheet[header_cols['price_col'] + str(row)].value
				}
				products.append(product)
		return products


	def is_valid_row(self, sheet, row, header_cols):
		"""Retorna si una fila corresponde o no a un producto."""

		for col in header_cols.values():
			if sheet[col + str(row)].value is None:
				return False
		return True


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
			self.tableWidget_search_etma: self.label_search_etma
		}
		if table_widget in search_tables:
			quantity = len(products)
			s = '' if quantity == 1 else 's'
			search_tables[table_widget].setText(f'{quantity} producto{s} encontrado{s}')


	def open_config(self):
		"""Abre un dialog para editar la configuración."""

		dialog = ConfigurationDialog(self)
		dialog.exec()

		# Verifico si recargar
		if dialog.new_price_lists_url:
			pass
			# self.cargarClientes()


	def open_about(self):
		"""Abre un único diálogo de Acerca de."""

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
		self.lineEdit_url_tdc.setText(SETTINGS.value('price_lists/tdc', '', type=str))
		self.lineEdit_url_camba.setText(SETTINGS.value('price_lists/camba', '', type=str))


	def save_config(self):
		SETTINGS.setValue('price_lists/tdc', self.lineEdit_url_tdc.text())
		SETTINGS.setValue('price_lists/camba', self.lineEdit_url_camba.text())
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