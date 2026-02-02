import openpyxl
import requests
import bs4
import tempfile
from pathlib import Path
import re
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import QMainWindow, QApplication, QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QMessageBox
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QIcon, QFont
from PyQt6 import uic
import sys

# URL_EXCEL = 'https://comunicaciones.tiendadecardan.com.ar/hubfs/LISTA%20DE%20PRECIO%20HH%20ENERO%206-1-2026.xlsx?hsLang=es-ar'

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


def get_excel_link(page_url):
	html = requests.get(page_url).text
	soup = bs4.BeautifulSoup(html, 'html.parser')

	h1 = soup.select()





# def download_excel_file(url):
# 	response = requests.get(url, timeout=30)
# 	response.raise_for_status()

# 	temp_dir = Path(tempfile.gettempdir())
# 	file_path = temp_dir / "lista_precios.xlsx"

# 	with open(file_path, "wb") as f:
# 		f.write(response.content)

# 	return file_path


# ruta_excel = download_excel_file(URL_EXCEL)
# print("Archivo descargado en:", ruta_excel)


class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo la UI
		uic.loadUi('ui/app.ui', self)

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


	# def get_excels(self):



	def initialize(self):
		"""."""

		# Obtengo URL de TDC donde se encuentran las listas de precios
		price_lists_url = SETTINGS.value('price_lists_url', '', type=str)
		if not price_lists_url:
			QMessageBox.warning(
				self,
				'Sin URL de listas',
				'Debe cargar una URL válida para buscar las listas.'
			)
			return

		# Obtengo el HTML para buscar los links de los excel
		try:
			response = requests.get(price_lists_url, timeout=10)
			response.raise_for_status()
			html = response.text
		except Exception as e:
			QMessageBox.critical(
				self,
				'Error',
				f'No se pudo acceder a la página de las listas de precios\nde Tienda Del Cardan'
			)
			return


		for brand in ('etma', 'hh'):
			link = self.get_excel_link_tdc(html, brand)





		# Creo los workbook y extraigo las hojas de productos
		wb_hh = openpyxl.load_workbook('LISTA DE PRECIO HH ENERO 6-1-2026.xlsx')
		wb_etma = openpyxl.load_workbook('LISTA DE PRECIO ETMA ENERO 6-1-2026.xlsx')
		sheet_hh = wb_hh[wb_hh.sheetnames[0]]
		sheet_etma = wb_etma[wb_etma.sheetnames[0]]

		# Busco letras de columnas de producto en cada hoja
		header_cols_hh = self.search_header_cols(sheet_hh)
		header_cols_etma = self.search_header_cols(sheet_etma)

		# Busco número de fila de primer producto en cada hoja
		first_row_hh = self.search_first_row(sheet_hh, header_cols_hh['price_col'])
		first_row_etma = self.search_first_row(sheet_etma, header_cols_etma['price_col'])

		# Muestro fecha de validez de precios de cada hoja
		self.show_validity_date(sheet_hh, self.label_validity_date_hh)
		self.show_validity_date(sheet_etma, self.label_validity_date_etma)

		# Paso a diccionarios los productos de cada hoja
		self.all_products_hh = self.obtain_products(sheet_hh, first_row_hh, header_cols_hh)
		self.all_products_etma = self.obtain_products(sheet_etma, first_row_etma, header_cols_etma)

		self.list_products(self.all_products_hh, self.tableWidget_search_hh)
		self.list_products(self.all_products_etma, self.tableWidget_search_etma)

		self.load_more_used(self.comboBox_most_used_hh, self.all_products_hh, MOST_USED_PRODUCTS_HH)
		self.load_more_used(self.comboBox_most_used_etma, self.all_products_etma, MOST_USED_PRODUCTS_ETMA)


	def get_excel_link_tdc(html, brand):
    	"""
    	Obtiene en TDC el link actual del excel de la lista de precios de 
    	la marca que se le pasa.
    	"""

	    soup = bs4.BeautifulSoup(html, 'html.parser')

	    
        # Busco el título correcto
        h1 = soup.find(
            'h1', 
            string=lambda s: s and s.strip().upper() in (f'LISTA DHE PRECIO {brand}', f'LISTA DE PRECIOS {brand}')
        )
        if not h1:
            raise ValueError(f'No se encontró la lista para {brand.upper()}')
            continue

        # Subo al bloque contenedor
        bloque = h1.find_parent('div', class_='widget-span')

        # Busco el link dentro del bloque
        link = bloque.find_next('a', href=True)
        if not link:
            # raise ValueError(f'No se encontró el link de descarga para {brand.upper()}')
            continue

        return link['href']


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
					float(cell.value.replace('.', ''). replace(',', '.'))
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
		self.lineEdit_url.setText(SETTINGS.value('price_lists_url', '', type=str))


	def save_config(self):
		SETTINGS.setValue('price_lists_url', self.lineEdit_url.text())
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