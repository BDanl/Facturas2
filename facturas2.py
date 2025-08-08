import sys
import json
import csv
import locale
import logging
import os
from datetime import datetime, date
from pathlib import Path
from collections import defaultdict
import webbrowser
import configparser
from database import Database
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.worksheet.table import Table, TableStyleInfo

# Importaciones de PyQt6
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QDateEdit, 
                             QTableWidget, QTableWidgetItem, QTabWidget, QMessageBox, 
                             QFileDialog, QHeaderView, QTextEdit, QCheckBox, QSplitter,
                             QStyleFactory, QStyle, QTableWidgetSelectionRange, QStatusBar,
                             QGroupBox, QFormLayout, QSpacerItem, QSizePolicy, QTreeWidget, 
                             QTreeWidgetItem, QMenu, QDialog, QListWidget, QDialogButtonBox, 
                             QListWidgetItem, QProgressDialog)
from PyQt6.QtGui import QAction, QFont, QColor, QIcon, QDoubleValidator, QTextCursor
from PyQt6.QtCore import Qt, QSize, QDate, QTimer


# Importaciones para Excel
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as e:
    logging.error(f"Error al importar dependencias de Excel: {str(e)}")
    QMessageBox.critical(None, "Error", "Error al importar dependencias de Excel. Asegúrate de tener openpyxl instalado.")
    sys.exit(1)

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('facturas_qt.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuración de la aplicación
CONFIG_FILE = 'config.ini'

def get_config():
    """Obtener la configuración de la aplicación"""
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
    
    # Configuración por defecto
    if 'APP' not in config:
        config['APP'] = {}
    if 'last_export_dir' not in config['APP']:
        config['APP']['last_export_dir'] = str(Path.home() / 'Documents')
    
    return config

def save_config(config):
    """Guardar la configuración de la aplicación"""
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

# Configurar formato de moneda colombiana
try:
    locale.setlocale(locale.LC_ALL, 'es_CO.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Spanish_Colombia.1252')
    except:
        locale.setlocale(locale.LC_ALL, '')

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestor de Facturas")
        # Remove fixed geometry and use showMaximized() after UI is initialized
        
        # Determinar si estamos ejecutando desde un ejecutable o desde el código fuente
        if getattr(sys, 'frozen', False):
            # Si es un ejecutable, usamos el directorio del ejecutable
            self.app_dir = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(sys.executable).parent
            self.data_dir = Path.home() / "FacturasApp"
        else:
            # Si es código fuente, usamos el directorio del proyecto
            self.app_dir = Path(__file__).parent
            self.data_dir = self.app_dir
        
        # Asegurarse de que el directorio de datos exista
        self.data_dir.mkdir(exist_ok=True, parents=True)
        
        # Inicializar la base de datos SQLite
        self.db = Database(str(self.data_dir / "facturas.db"))
        
        # Verificar si hay que migrar datos desde el archivo JSON antiguo
        self._migrar_datos_desde_json()
        
        # Configuración del tema
        self.tema_oscuro = False
        
        # Inicializar atributos
        self.facturas = []
        self.tipos_gasto = []
        
        # Cargar datos sin actualizar la UI aún
        self.cargar_datos(actualizar_ui=False)
        
        # Inicializar la interfaz de usuario
        self.init_ui()
        
        # Actualizar la UI con los datos cargados
        self.actualizar_lista_facturas()
        self.actualizar_filtros()
        self.actualizar_resumen()
        
        # Maximizar la ventana después de inicializar la UI
        self.showMaximized()
    
    def _migrar_datos_desde_json(self):
        """Migra los datos desde el archivo JSON antiguo a la base de datos SQLite si es necesario."""
        json_path = Path("facturas_qt.json")
        if json_path.exists():
            try:
                # Verificar si ya hay datos en la base de datos
                facturas = self.db.obtener_facturas()
                if not facturas:
                    # Migrar solo si no hay datos en la base de datos
                    self.db.migrar_desde_json(str(json_path))
                    # Opcional: respaldar el archivo JSON después de la migración
                    backup_path = json_path.with_suffix(f".{datetime.now().strftime('%Y%m%d%H%M%S')}.json")
                    json_path.rename(backup_path)
                    logging.info(f"Datos migrados exitosamente. Archivo original respaldado como {backup_path}")
            except Exception as e:
                logging.error(f"Error al migrar datos desde JSON: {str(e)}")
    
    def init_ui(self):
        """Inicializar la interfaz de usuario"""
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(10)
        
        # Barra de herramientas
        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 10)
        
        # Título
        title = QLabel("Sistema de Gestión de Facturas")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Botón para alternar modo oscuro/claro
        self.btn_tema = QPushButton("Modo Oscuro")
        self.btn_tema.setToolTip("Haz clic para cambiar entre modo claro y oscuro")
        self.btn_tema.setFixedSize(120, 32)
        self.btn_tema.setStyleSheet("""
            QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
            QPushButton:pressed {
                background-color: #4e555b;
            }
        """)
        self.btn_tema.clicked.connect(self.cambiar_tema)
        
        # Agregar widgets a la barra de herramientas
        toolbar.addStretch()
        toolbar.addWidget(title)
        toolbar.addStretch()
        toolbar.addWidget(self.btn_tema)
        
        # Agregar barra de herramientas al layout principal
        main_layout.addLayout(toolbar)
        
        # Inicializar tema (por defecto: modo claro)
        self.cargar_preferencia_tema()
        
        # Configurar pestañas
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Pestaña de registro
        self.tab_registro = QWidget()
        self.tabs.addTab(self.tab_registro, "Registrar Factura")
        self.setup_registro_tab()
        
        # Pestaña de resúmenes (antes Filtros Avanzados)
        self.tab_filtros = QWidget()
        self.tabs.addTab(self.tab_filtros, "Resúmenes")
        self.setup_filtros_tab()
        
        # Pestaña de lista de facturas
        self.tab_lista = QWidget()
        self.tabs.addTab(self.tab_lista, "Lista de Facturas")
        self.setup_lista_tab()
        
        # Barra de estado
        self.statusBar().showMessage("Listo")
    
    def setup_registro_tab(self):
        """Configurar la pestaña de registro de facturas"""
        layout = QFormLayout(self.tab_registro)
        
        # Tipo de gasto
        self.cmb_tipo_gasto = QComboBox()
        self.cmb_tipo_gasto.addItems(["Mercado", "Transporte", "Compra tienda", 
                                    "Farmacia", "Varios", "Gastos urgentes"])
        
        # Descripción
        self.txt_descripcion = QLineEdit()
        
        # Valor
        self.txt_valor = QLineEdit()
        self.txt_valor.setValidator(QDoubleValidator(0, 999999999, 2, self))
        
        # Fecha
        self.date_fecha = QDateEdit()
        self.date_fecha.setCalendarPopup(True)
        self.date_fecha.setDate(QDate.currentDate())
        self.date_fecha.setDisplayFormat("dd/MM/yyyy")
        
        # Botón de guardar
        self.btn_guardar = QPushButton("Guardar Factura")
        self.btn_guardar.clicked.connect(self.guardar_factura)
        
        # Agregar widgets al layout
        layout.addRow("Tipo de Gasto:", self.cmb_tipo_gasto)
        layout.addRow("Descripción:", self.txt_descripcion)
        layout.addRow("Valor (COP):", self.txt_valor)
        layout.addRow("Fecha:", self.date_fecha)
        layout.addRow(self.btn_guardar)
    
    def setup_resumen_tab(self):
        """Configurar la pestaña de resumen"""
        # Crear pestañas para diferentes vistas de resumen
        self.tabs_resumen = QTabWidget()
        
        # Pestaña de resumen diario
        self.tab_diario = QWidget()
        self.tabs_resumen.addTab(self.tab_diario, "Resumen Diario")
        self.setup_resumen_diario_tab()
        
        # Pestaña de resumen mensual
        self.tab_mensual = QWidget()
        self.tabs_resumen.addTab(self.tab_mensual, "Resumen Mensual")
        self.setup_resumen_mensual_tab()
        
        # Pestaña de resumen anual
        self.tab_anual = QWidget()
        self.tabs_resumen.addTab(self.tab_anual, "Resumen Anual")
        self.setup_resumen_anual_tab()
        
        # Pestaña de filtros avanzados
        self.tab_filtros = QWidget()
        self.tabs_resumen.addTab(self.tab_filtros, "Filtros Avanzados")
        self.setup_filtros_tab()
        
        # Layout principal
        layout = QVBoxLayout(self.tab_resumen)
        layout.addWidget(self.tabs_resumen)
        
        # Actualizar los resúmenes
        self.actualizar_resumen()
    
    def setup_resumen_diario_tab(self):
        """Configurar la pestaña de resumen diario"""
        layout = QVBoxLayout(self.tab_diario)
        
        # Fecha para el resumen diario
        self.date_resumen_diario = QDateEdit()
        self.date_resumen_diario.setCalendarPopup(True)
        self.date_resumen_diario.setDate(QDate.currentDate())
        self.date_resumen_diario.setDisplayFormat("dd/MM/yyyy")
        self.date_resumen_diario.dateChanged.connect(self.actualizar_resumen_diario)
        
        # Botón para ir a hoy
        btn_hoy = QPushButton("Hoy")
        btn_hoy.clicked.connect(lambda: self.date_resumen_diario.setDate(QDate.currentDate()))
        
        # Layout para controles de fecha
        fecha_layout = QHBoxLayout()
        fecha_layout.addWidget(QLabel("Seleccione una fecha:"))
        fecha_layout.addWidget(self.date_resumen_diario)
        fecha_layout.addWidget(btn_hoy)
        fecha_layout.addStretch()
        
        # Área de texto para el resumen
        self.texto_resumen_diario = QTextEdit()
        self.texto_resumen_diario.setReadOnly(True)
        
        # Agregar al layout principal
        layout.addLayout(fecha_layout)
        layout.addWidget(self.texto_resumen_diario)
    
    def setup_resumen_mensual_tab(self):
        """Configurar la pestaña de resumen mensual"""
        layout = QVBoxLayout(self.tab_mensual)
        
        # Controles para seleccionar mes y año
        control_layout = QHBoxLayout()
        
        # Combo para seleccionar mes
        self.combo_mes_resumen = QComboBox()
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.combo_mes_resumen.addItems(meses)
        self.combo_mes_resumen.setCurrentIndex(QDate.currentDate().month() - 1)
        self.combo_mes_resumen.currentIndexChanged.connect(self.actualizar_resumen_mensual)
        
        # Combo para seleccionar año
        self.combo_anio_resumen = QComboBox()
        anio_actual = QDate.currentDate().year()
        self.combo_anio_resumen.addItems([str(anio) for anio in range(anio_actual - 5, anio_actual + 1)])
        self.combo_anio_resumen.setCurrentText(str(anio_actual))
        self.combo_anio_resumen.currentTextChanged.connect(self.actualizar_resumen_mensual)
        
        # Agregar controles al layout
        control_layout.addWidget(QLabel("Mes:"))
        control_layout.addWidget(self.combo_mes_resumen)
        control_layout.addWidget(QLabel("Año:"))
        control_layout.addWidget(self.combo_anio_resumen)
        control_layout.addStretch()
        
        # Área de texto para el resumen
        self.texto_resumen_mensual = QTextEdit()
        self.texto_resumen_mensual.setReadOnly(True)
        
        # Agregar al layout principal
        layout.addLayout(control_layout)
        layout.addWidget(self.texto_resumen_mensual)
        
        # Cargar datos iniciales
        self.actualizar_resumen_mensual()
    
    def setup_resumen_anual_tab(self):
        """Configurar la pestaña de resumen anual"""
        layout = QVBoxLayout(self.tab_anual)
        
        # Controles para seleccionar año
        control_layout = QHBoxLayout()
        
        # Combo para seleccionar año
        self.combo_anio_anual = QComboBox()
        anio_actual = QDate.currentDate().year()
        self.combo_anio_anual.addItems([str(anio) for anio in range(anio_actual - 10, anio_actual + 1)])
        self.combo_anio_anual.setCurrentText(str(anio_actual))
        self.combo_anio_anual.currentTextChanged.connect(self.actualizar_resumen_anual)
        
        # Agregar controles al layout
        control_layout.addWidget(QLabel("Año:"))
        control_layout.addWidget(self.combo_anio_anual)
        control_layout.addStretch()
        
        # Área de texto para el resumen
        self.texto_resumen_anual = QTextEdit()
        self.texto_resumen_anual.setReadOnly(True)
        
        # Agregar al layout principal
        layout.addLayout(control_layout)
        layout.addWidget(self.texto_resumen_anual)
        
        # Cargar datos iniciales
        self.actualizar_resumen_anual()
    
    def setup_filtros_tab(self):
        """Configurar la pestaña de filtros avanzados"""
        layout = QVBoxLayout(self.tab_filtros)
        
        # Grupo para filtros
        grupo_filtros = QGroupBox("Filtrar por")
        form_filtros = QFormLayout()
        
        # Filtro por año
        self.combo_filtro_anio = QComboBox()
        self.combo_filtro_anio.addItem("Todos los años", None)
        self.combo_filtro_anio.currentIndexChanged.connect(self.aplicar_filtros)
        
        # Filtro por mes
        self.combo_filtro_mes = QComboBox()
        self.combo_filtro_mes.addItem("Todos los meses", None)
        self.combo_filtro_mes.addItems([f"{i:02d} - {mes}" for i, mes in enumerate([
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ], 1)])
        self.combo_filtro_mes.currentIndexChanged.connect(self.aplicar_filtros)
        
        # Filtro por día
        self.combo_filtro_dia = QComboBox()
        self.combo_filtro_dia.addItem("Todos los días", None)
        self.combo_filtro_dia.addItems([f"{i:02d}" for i in range(1, 32)])
        self.combo_filtro_dia.currentIndexChanged.connect(self.aplicar_filtros)
        
        # Filtro por tipo de gasto
        self.combo_filtro_tipo = QComboBox()
        self.combo_filtro_tipo.addItem("Todos los tipos", None)
        self.combo_filtro_tipo.addItems(["Mercado", "Transporte", "Compra tienda", 
                                       "Farmacia", "Varios", "Gastos urgentes"])
        self.combo_filtro_tipo.currentIndexChanged.connect(self.aplicar_filtros)
        
        # Botón para limpiar filtros
        btn_limpiar = QPushButton("Limpiar Filtros")
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        
        # Agregar controles al formulario
        form_filtros.addRow("Año:", self.combo_filtro_anio)
        form_filtros.addRow("Mes:", self.combo_filtro_mes)
        form_filtros.addRow("Día:", self.combo_filtro_dia)
        form_filtros.addRow("Tipo de gasto:", self.combo_filtro_tipo)
        form_filtros.addRow(btn_limpiar)
        
        grupo_filtros.setLayout(form_filtros)
        
        # Tabla para mostrar resultados filtrados
        self.tabla_filtrada = QTableWidget()
        self.tabla_filtrada.setColumnCount(4)
        self.tabla_filtrada.setHorizontalHeaderLabels(["Fecha", "Tipo", "Descripción", "Valor"])
        
        # Configurar cabeceras
        header = self.tabla_filtrada.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # Fecha
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # Tipo
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)  # Descripción
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)  # Valor
        
        # Agregar widgets al layout principal
        layout.addWidget(grupo_filtros)
        layout.addWidget(self.tabla_filtrada)
        
        # Inicializar filtros y aplicar automáticamente
        self.inicializar_filtros()
        self.aplicar_filtros()
    
    def setup_lista_tab(self):
        """Configurar la pestaña de lista de facturas"""
        
        layout = QVBoxLayout(self.tab_lista)
        
        # Layout para los botones de acción
        btn_layout = QHBoxLayout()
        
        # Menú desplegable para importar desde diferentes formatos
        self.menu_importar = QPushButton("Importar Facturas ▼")
        self.menu_importar.setToolTip("Importar facturas desde diferentes formatos")
        self.menu_importar.setFixedSize(150, 32)
        
        # Crear menú desplegable
        self.import_menu = QMenu(self)
        self.import_menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #dee2e6;
                padding: 5px;
            }
            QMenu::item {
                padding: 5px 15px;
            }
            QMenu::item:selected {
                background-color: #0d6efd;
                color: white;
            }
        """)
        
        # Acciones del menú
        accion_json = QAction("Importar desde JSON", self)
        accion_csv = QAction("Importar desde CSV", self)
        accion_excel = QAction("Importar desde Excel", self)
        
        # Conectar acciones
        accion_json.triggered.connect(self.importar_desde_json)
        accion_csv.triggered.connect(self.importar_desde_csv)
        accion_excel.triggered.connect(self.importar_desde_excel)
        
        # Agregar acciones al menú
        self.import_menu.addAction(accion_json)
        self.import_menu.addAction(accion_csv)
        self.import_menu.addAction(accion_excel)
        
        # Configurar el botón para mostrar el menú
        self.menu_importar.setMenu(self.import_menu)
        btn_layout.addWidget(self.menu_importar)
        
        # Botón para limpiar todo
        self.btn_limpiar_todo = QPushButton("Limpiar Todo")
        self.btn_limpiar_todo.clicked.connect(self.confirmar_limpiar_todo)
        self.btn_limpiar_todo.setToolTip("Eliminar todas las facturas")
        self.btn_limpiar_todo.setStyleSheet("background-color: #f8d7da; color: #721c24;")
        btn_layout.addWidget(self.btn_limpiar_todo)
        
        # Botón para exportar a Excel
        self.btn_exportar = QPushButton("Exportar a Excel")
        self.btn_exportar.clicked.connect(self.exportar_a_excel)
        self.btn_exportar.setToolTip("Exportar facturas a un archivo Excel")
        btn_layout.addWidget(self.btn_exportar)
        
        # Layout para la tabla y botón de eliminar seleccionadas
        table_layout = QVBoxLayout()
        
        # Tabla de facturas
        self.tabla_facturas = QTableWidget()
        self.tabla_facturas.setColumnCount(5)  # Una columna extra para el checkbox
        self.tabla_facturas.setHorizontalHeaderLabels(["", "Fecha", "Tipo", "Descripción", "Valor"])
        self.tabla_facturas.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla_facturas.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        
        # Ocultar la columna de checkboxes (la usaremos para selección)
        self.tabla_facturas.setColumnHidden(0, True)
        
        # Ajustar el tamaño de las columnas
        header = self.tabla_facturas.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # Columna oculta para checkboxes
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # Fecha
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)  # Tipo
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)  # Descripción
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)  # Valor
        
        # Hacer que las filas sean seleccionables
        self.tabla_facturas.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        
        # Layout para el botón de eliminar seleccionadas
        bottom_btn_layout = QHBoxLayout()
        self.btn_eliminar = QPushButton("Eliminar Seleccionadas")
        self.btn_eliminar.clicked.connect(self.eliminar_facturas_seleccionadas)
        self.btn_eliminar.setToolTip("Eliminar las facturas seleccionadas")
        self.btn_eliminar.setEnabled(False)  # Deshabilitado inicialmente
        self.btn_eliminar.setStyleSheet("background-color: #f8d7da; color: #721c24;")
        
        # Conectar señal de selección para habilitar/deshabilitar el botón
        self.tabla_facturas.itemSelectionChanged.connect(self.actualizar_boton_eliminar)
        
        bottom_btn_layout.addWidget(self.btn_eliminar)
        bottom_btn_layout.addStretch()
        
        # Agregar widgets al layout principal
        layout.addLayout(btn_layout)
        layout.addWidget(self.tabla_facturas)
        layout.addLayout(bottom_btn_layout)
    
    def cargar_datos(self):
        """Cargar datos desde la base de datos"""
        try:
            self.facturas = self.db.obtener_facturas()
            logger.info(f"Datos cargados correctamente desde la base de datos")
        except Exception as e:
            logger.error(f"Error al cargar los datos: {str(e)}")
            QMessageBox.critical(self, "Error", f"No se pudieron cargar los datos: {str(e)}")
    
    def guardar_datos(self):
        """Guardar datos en la base de datos"""
        try:
            self.db.guardar_facturas(self.facturas)
            return True
        except Exception as e:
            logger.error(f"Error al guardar datos: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Error al guardar datos: {str(e)}")
            return False
    
    def guardar_factura(self):
        """Guardar una nueva factura"""
        # Validar campos
        if not self.validar_campos():
            return
        
        # Crear diccionario con los datos de la factura
        factura = {
            'fecha': self.date_fecha.date().toString("dd/MM/yyyy"),
            'tipo': self.cmb_tipo_gasto.currentText(),
            'descripcion': self.txt_descripcion.text().strip(),
            'valor': float(self.txt_valor.text().replace(',', '.'))
        }
        
        # Agregar a la lista
        self.facturas.append(factura)
        
        # Guardar datos
        if self.guardar_datos():
            # Actualizar interfaz
            self.limpiar_campos()
            self.actualizar_lista_facturas()
            self.actualizar_resumen()
            self.statusBar().showMessage("Factura guardada correctamente", 3000)
    
    def validar_campos(self):
        """Validar los campos del formulario"""
        if not self.txt_descripcion.text().strip():
            QMessageBox.warning(self, "Validación", "La descripción es obligatoria")
            self.txt_descripcion.setFocus()
            return False
            
        if not self.txt_valor.text().strip() or float(self.txt_valor.text().replace(',', '.')) <= 0:
            QMessageBox.warning(self, "Validación", "El valor debe ser mayor a cero")
            self.txt_valor.setFocus()
            return False
            
        return True
    
    def limpiar_campos(self):
        """Limpiar los campos del formulario"""
        self.txt_descripcion.clear()
        self.txt_valor.clear()
        self.date_fecha.setDate(QDate.currentDate())
        self.cmb_tipo_gasto.setCurrentIndex(0)
    
    def actualizar_lista_facturas(self):
        """Actualizar la tabla de facturas"""
        self.tabla_facturas.setRowCount(len(self.facturas))
        
        for i, factura in enumerate(self.facturas):
            # Columna 0 está oculta (para checkboxes), empezamos desde la columna 1
            self.tabla_facturas.setItem(i, 1, QTableWidgetItem(factura['fecha']))
            self.tabla_facturas.setItem(i, 2, QTableWidgetItem(factura['tipo']))
            self.tabla_facturas.setItem(i, 3, QTableWidgetItem(factura['descripcion']))
            self.tabla_facturas.setItem(i, 4, QTableWidgetItem(f"${factura['valor']:,.0f} COP".replace(',', '.')))
    
    def actualizar_resumen(self):
        """Actualizar todos los resúmenes"""
        self.actualizar_resumen_diario()
        self.actualizar_resumen_mensual()
        self.actualizar_resumen_anual()
        self.actualizar_filtros()
    
    def actualizar_resumen_diario(self):
        """Actualizar el resumen diario"""
        # Verificar si los widgets de la interfaz están inicializados
        if not hasattr(self, 'date_resumen_diario') or not hasattr(self, 'texto_resumen_diario'):
            return
            
        try:
            fecha_seleccionada = self.date_resumen_diario.date()
            fecha_str = fecha_seleccionada.toString("dd/MM/yyyy")
            
            resumen = defaultdict(float)
            total = 0.0
            
            for factura in self.facturas:
                if factura['fecha'] == fecha_str:
                    tipo = factura['tipo']
                    valor = factura['valor']
                    resumen[tipo] += valor
                    total += valor
            
            # Formatear el resumen
            texto = f"Resumen de gastos para {fecha_str}\n\n"
            for tipo, monto in sorted(resumen.items()):
                texto += f"{tipo}: ${monto:,.0f} COP\n"
            
            texto += f"\nTotal del día: ${total:,.0f} COP"
            
            # Mostrar en el área de texto
            self.texto_resumen_diario.setPlainText(texto)
        except Exception as e:
            logger.error(f"Error al actualizar resumen diario: {str(e)}")
    
    def actualizar_resumen_mensual(self):
        """Actualizar el resumen mensual"""
        # Verificar si los widgets de la interfaz están inicializados
        if not all(hasattr(self, attr) for attr in ['combo_mes_resumen', 'combo_anio_resumen', 'texto_resumen_mensual']):
            return
            
        try:
            mes = self.combo_mes_resumen.currentIndex() + 1
            anio = int(self.combo_anio_resumen.currentText())
            
            resumen = defaultdict(float)
            total = 0.0
            
            for factura in self.facturas:
                try:
                    fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    if fecha.month == mes and fecha.year == anio:
                        tipo = factura['tipo']
                        valor = factura['valor']
                        resumen[tipo] += valor
                        total += valor
                except Exception as e:
                    logger.warning(f"Error al procesar factura: {str(e)}")
                    continue
            
            # Formatear el resumen
            nombre_mes = self.combo_mes_resumen.currentText()
            texto = f"Resumen de gastos para {nombre_mes} de {anio}\n\n"
            
            if resumen:
                for tipo, monto in sorted(resumen.items()):
                    porcentaje = (monto / total) * 100 if total > 0 else 0
                    texto += f"{tipo}: ${monto:,.0f} COP ({porcentaje:.1f}%)\n"
                
                texto += f"\nTotal del mes: ${total:,.0f} COP"
            else:
                texto += "No hay datos para mostrar en este período."
            
            # Mostrar en el área de texto
            self.texto_resumen_mensual.setPlainText(texto)
        except Exception as e:
            logger.error(f"Error al actualizar resumen mensual: {str(e)}")
    
    def actualizar_resumen_anual(self):
        """Actualizar el resumen anual"""
        # Verificar si los widgets de la interfaz están inicializados
        if not all(hasattr(self, attr) for attr in ['combo_anio_anual', 'texto_resumen_anual']):
            return
            
        try:
            anio = int(self.combo_anio_anual.currentText())
            
            resumen_mensual = defaultdict(lambda: defaultdict(float))
            resumen_anual = defaultdict(float)
            total_anual = 0.0
            
            for factura in self.facturas:
                try:
                    fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    if fecha.year == anio:
                        mes = fecha.month
                        tipo = factura['tipo']
                        valor = factura['valor']
                        
                        resumen_mensual[mes][tipo] += valor
                        resumen_anual[tipo] += valor
                        total_anual += valor
                except Exception as e:
                    logger.warning(f"Error al procesar factura: {str(e)}")
                    continue
            
            # Formatear el resumen
            texto = f"Resumen de gastos para el año {anio}\n\n"
            
            if resumen_anual:
                # Resumen por tipo de gasto
                texto += "=== Resumen por categoría ===\n"
                for tipo, monto in sorted(resumen_anual.items()):
                    porcentaje = (monto / total_anual) * 100 if total_anual > 0 else 0
                    texto += f"{tipo}: ${monto:,.0f} COP ({porcentaje:.1f}%)\n"
                
                # Resumen mensual
                texto += "\n=== Resumen mensual ===\n"
                meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                
                for mes in range(1, 13):
                    total_mes = sum(resumen_mensual[mes].values())
                    if total_mes > 0:
                        porcentaje = (total_mes / total_anual) * 100 if total_anual > 0 else 0
                        texto += f"{meses[mes-1]}: ${total_mes:,.0f} COP ({porcentaje:.1f}%)\n"
                
                texto += f"\nTotal anual: ${total_anual:,.0f} COP"
            else:
                texto += "No hay datos para mostrar en este año."
            
            # Mostrar en el área de texto
            self.texto_resumen_anual.setPlainText(texto)
        except Exception as e:
            logger.error(f"Error al actualizar resumen anual: {str(e)}")
    
    def inicializar_filtros(self):
        """Inicializar los valores de los filtros"""
        # Obtener años únicos
        anios = set()
        for factura in self.facturas:
            try:
                fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                anios.add(fecha.year)
            except:
                continue
        
        # Ordenar años de mayor a menor
        anios_ordenados = sorted(anios, reverse=True)
        
        # Actualizar combo de años
        self.combo_filtro_anio.clear()
        self.combo_filtro_anio.addItem("Todos los años", None)
        for anio in anios_ordenados:
            self.combo_filtro_anio.addItem(str(anio), anio)
    
    def actualizar_filtros(self):
        """Actualizar los controles de filtro con los datos actuales"""
        self.inicializar_filtros()
        self.aplicar_filtros()
    
    def aplicar_filtros(self):
        """Aplicar los filtros seleccionados"""
        anio = self.combo_filtro_anio.currentData()
        mes = self.combo_filtro_mes.currentIndex()  # 0 = Todos, 1-12 = meses
        dia = self.combo_filtro_dia.currentIndex()  # 0 = Todos, 1-31 = días
        tipo = self.combo_filtro_tipo.currentText() if self.combo_filtro_tipo.currentIndex() > 0 else None
        
        # Filtrar facturas
        facturas_filtradas = []
        for factura in self.facturas:
            try:
                fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                
                # Verificar año
                if anio is not None and fecha.year != anio:
                    continue
                
                # Verificar mes
                if mes > 0 and fecha.month != mes:  # Si no es "Todos los meses"
                    continue
                
                # Verificar día
                if dia > 0 and fecha.day != dia:  # Si no es "Todos los días"
                    continue
                
                # Verificar tipo
                if tipo is not None and factura['tipo'] != tipo:
                    continue
                
                facturas_filtradas.append(factura)
            except Exception as e:
                logger.error(f"Error al procesar factura {factura}: {str(e)}")
                continue
        
        # Mostrar resultados
        self.mostrar_resultados_filtrados(facturas_filtradas)
    
    def limpiar_filtros(self):
        """Limpiar todos los filtros"""
        self.combo_filtro_anio.setCurrentIndex(0)
        self.combo_filtro_mes.setCurrentIndex(0)
        self.combo_filtro_dia.setCurrentIndex(0)
        self.combo_filtro_tipo.setCurrentIndex(0)
        self.aplicar_filtros()
    
    def mostrar_resultados_filtrados(self, facturas):
        """Mostrar las facturas filtradas en la tabla"""
        self.tabla_filtrada.setRowCount(len(facturas))
        
        for i, factura in enumerate(facturas):
            self.tabla_filtrada.setItem(i, 0, QTableWidgetItem(factura['fecha']))
            self.tabla_filtrada.setItem(i, 1, QTableWidgetItem(factura['tipo']))
            self.tabla_filtrada.setItem(i, 2, QTableWidgetItem(factura['descripcion']))
            self.tabla_filtrada.setItem(i, 3, QTableWidgetItem(f"${factura['valor']:,.0f} COP".replace(',', '.')))
        
        # Calcular total
        total = sum(factura['valor'] for factura in facturas)
        
        # Agregar fila de total
        self.tabla_filtrada.setRowCount(len(facturas) + 1)
        self.tabla_filtrada.setItem(len(facturas), 0, QTableWidgetItem(""))
        self.tabla_filtrada.setItem(len(facturas), 1, QTableWidgetItem(""))
        self.tabla_filtrada.setItem(len(facturas), 2, QTableWidgetItem("TOTAL:"))
        self.tabla_filtrada.setItem(len(facturas), 3, QTableWidgetItem(f"${total:,.0f} COP".replace(',', '.')))
        
        # Resaltar la fila de total
        for col in range(self.tabla_filtrada.columnCount()):
            item = self.tabla_filtrada.item(len(facturas), col)
            if item:
                item.setBackground(QColor(230, 230, 230))
                if col == 2 or col == 3:  # Solo las celdas de texto
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)

    def _mostrar_vista_previa(self, facturas, tipo_archivo):
        """Mostrar una vista previa de las facturas a importar"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Vista Previa - {tipo_archivo}")
        dialog.resize(800, 500)
        
        layout = QVBoxLayout(dialog)
        
        # Crear tabla para la vista previa
        tabla = QTableWidget(dialog)
        tabla.setColumnCount(5)
        tabla.setHorizontalHeaderLabels(["Fecha", "Tipo", "Descripción", "Valor (COP)", "Estado"])
        tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # Mostrar hasta 10 filas como vista previa
        num_filas = min(10, len(facturas))
        tabla.setRowCount(num_filas)
        
        # Llenar la tabla con datos de vista previa
        for i in range(num_filas):
            factura = facturas[i]
            tabla.setItem(i, 0, QTableWidgetItem(factura.get('fecha', '')))
            tabla.setItem(i, 1, QTableWidgetItem(factura.get('tipo', '')))
            tabla.setItem(i, 2, QTableWidgetItem(factura.get('descripcion', '')))
            tabla.setItem(i, 3, QTableWidgetItem(f"${factura.get('valor', 0):,.0f} COP".replace(',', '.')))
            tabla.setItem(i, 4, QTableWidgetItem("✅ Válido"))
        
        # Mostrar resumen
        total_facturas = len(facturas)
        resumen = QLabel(f"Mostrando {num_filas} de {total_facturas} facturas encontradas.")
        
        # Botones
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        
        layout.addWidget(QLabel("Vista previa de las primeras facturas a importar:"))
        layout.addWidget(tabla)
        layout.addWidget(resumen)
        layout.addWidget(btn_box)
        
        return dialog.exec() == QDialog.DialogCode.Accepted
    
    def _procesar_importacion(self, facturas_importadas, tipo_archivo):
        """Procesar la importación con diálogo de progreso"""
        # Crear diálogo de progreso
        progress = QProgressDialog(
            f"Importando {len(facturas_importadas)} facturas desde {tipo_archivo}...",
            "Cancelar", 0, len(facturas_importadas), self)
        progress.setWindowTitle("Importando facturas")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(1000)  # Mostrar después de 1 segundo
        
        # Mostrar diálogo de confirmación con vista previa
        if not self._mostrar_vista_previa(facturas_importadas, tipo_archivo):
            return False
        
        # Preguntar al usuario si desea sobrescribir o agregar
        reply = QMessageBox.question(
            self,
            f'Importar {tipo_archivo}',
            f'¿Desea sobrescribir las facturas existentes con las {len(facturas_importadas)} facturas del archivo?\n' \
            '"No" agregará las facturas a las existentes.',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        
        if reply == QMessageBox.StandardButton.Cancel:
            return False
        
        # Guardar copia de seguridad de los datos actuales
        facturas_originales = self.facturas.copy()
        
        try:
            # Actualizar la barra de estado
            self.statusBar().showMessage(f"Importando {len(facturas_importadas)} facturas...")
            
            # Procesar la importación con barra de progreso
            for i, factura in enumerate(facturas_importadas, 1):
                if progress.wasCanceled():
                    break
                    
                progress.setValue(i)
                QApplication.processEvents()  # Mantener la interfaz responsiva
                
            # Aplicar los cambios según la opción seleccionada
            if reply == QMessageBox.StandardButton.Yes:
                self.facturas = facturas_importadas
            else:
                self.facturas.extend(facturas_importadas)
            
            # Guardar los datos
            if not self.guardar_datos():
                raise Exception("No se pudieron guardar los datos")
            
            # Actualizar la interfaz
            self.actualizar_tabla_facturas()
            self.actualizar_resumenes()
            
            # Mostrar notificación de éxito
            self.statusBar().showMessage(
                f"Se importaron {len(facturas_importadas)} facturas correctamente.", 
                5000  # Mostrar por 5 segundos
            )
            
            # Mostrar mensaje de éxito con más detalles
            QMessageBox.information(
                self, 
                "Importación exitosa", 
                f"Se importaron {len(facturas_importadas)} facturas desde el archivo {tipo_archivo}.\n\n"
                f"Total de facturas actuales: {len(self.facturas)}"
            )
            
            return True
            
        except Exception as e:
            # En caso de error, restaurar los datos originales
            self.facturas = facturas_originales
            error_msg = (
                f"Error al importar las facturas:\n\n"
                f"Error: {str(e)}\n\n"
                f"Se han restaurado los datos originales."
            )
            logger.error(f"Error en importación: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error en la importación", error_msg)
            return False
    
    def importar_desde_csv(self):
        """Importar facturas desde un archivo CSV"""
        # Abrir diálogo para seleccionar archivo
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo CSV",
            "",
            "Archivos CSV (*.csv);;Todos los archivos (*)"
        )
        
        if not file_path:
            return  # Usuario canceló el diálogo
        
        # Mostrar indicador de carga
        loading_box = QMessageBox(self)
        loading_box.setWindowTitle("Cargando archivo")
        loading_box.setText("Procesando archivo CSV, por favor espere...")
        loading_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
        loading_box.show()
        QApplication.processEvents()
        
        try:
            facturas_importadas = []
            errores = []
            total_lineas = 0
            
            # Primera pasada: contar líneas para la barra de progreso
            with open(file_path, 'r', encoding='utf-8') as f:
                total_lineas = sum(1 for _ in f) - 1  # Restar 1 por el encabezado
            
            # Crear diálogo de progreso para la lectura
            progress = QProgressDialog("Leyendo archivo CSV...", "Cancelar", 0, total_lineas, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setMinimumDuration(0)
            progress.setValue(0)
            
            # Segunda pasada: procesar el archivo
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                lineas_procesadas = 0
                
                # Verificar que el CSV tenga las columnas necesarias
                required_columns = {'fecha', 'tipo', 'descripcion', 'valor'}
                if not required_columns.issubset(reader.fieldnames):
                    progress.close()
                    QMessageBox.critical(
                        self, 
                        "Error de formato", 
                        f"El archivo CSV debe contener las columnas: {', '.join(required_columns)}\n\n"
                        f"Columnas encontradas: {', '.join(reader.fieldnames) if reader.fieldnames else 'Ninguna'}"
                    )
                    return
                
                for row in reader:
                    if progress.wasCanceled():
                        return
                        
                    lineas_procesadas += 1
                    progress.setValue(lineas_procesadas)
                    QApplication.processEvents()
                    
                    try:
                        # Validar y convertir los datos
                        fecha = row['fecha'].strip()
                        tipo = row['tipo'].strip()
                        descripcion = row['descripcion'].strip()
                        
                        # Validar campos obligatorios
                        if not all([fecha, tipo, descripcion, row['valor']]):
                            errores.append(f"Fila {lineas_procesadas + 1}: Faltan campos obligatorios")
                            continue
                        
                        try:
                            # Manejar diferentes formatos de valor (con comas, puntos, etc.)
                            valor_str = str(row['valor']).replace('$', '').replace(',', '').strip()
                            valor = float(valor_str)
                            if valor <= 0:
                                raise ValueError("El valor debe ser mayor a cero")
                        except (ValueError, TypeError) as ve:
                            errores.append(f"Fila {lineas_procesadas + 1}: Valor inválido - {str(ve)}")
                            continue
                        
                        # Validar formato de fecha (DD/MM/YYYY)
                        try:
                            fecha_dt = datetime.strptime(fecha, '%d/%m/%Y')
                            fecha = fecha_dt.strftime('%d/%m/%Y')  # Estandarizar formato
                        except ValueError:
                            errores.append(f"Fila {lineas_procesadas + 1}: Formato de fecha inválido (debe ser DD/MM/YYYY)")
                            continue
                        
                        # Crear objeto de factura
                        factura = {
                            'fecha': fecha,
                            'tipo': tipo,
                            'descripcion': descripcion,
                            'valor': valor,
                            'timestamp': datetime.now().isoformat()
                        }
                        facturas_importadas.append(factura)
                        
                    except Exception as e:
                        error_msg = f"Fila {lineas_procesadas + 1}: {str(e)}"
                        logger.error(f"Error al procesar fila del CSV: {row}. Error: {error_msg}")
                        errores.append(error_msg)
                        continue
            
            progress.close()
            loading_box.close()
            
            # Mostrar advertencias si hay errores
            if errores:
                error_dialog = QDialog(self)
                error_dialog.setWindowTitle("Advertencias de importación")
                error_dialog.resize(600, 400)
                
                layout = QVBoxLayout()
                
                # Mostrar resumen
                resumen = QLabel(
                    f"Se encontraron {len(errores)} advertencias durante la importación.\n"
                    f"Se importaron {len(facturas_importadas)} facturas correctamente."
                )
                
                # Área de texto para mostrar los errores
                error_text = QTextEdit()
                error_text.setReadOnly(True)
                error_text.setPlainText("\n".join(errores))
                
                # Botón para continuar
                btn_continuar = QPushButton("Continuar con la importación")
                btn_continuar.clicked.connect(error_dialog.accept)
                
                # Botón para cancelar
                btn_cancelar = QPushButton("Cancelar importación")
                btn_cancelar.clicked.connect(error_dialog.reject)
                
                # Layout de botones
                btn_layout = QHBoxLayout()
                btn_layout.addStretch()
                btn_layout.addWidget(btn_continuar)
                btn_layout.addWidget(btn_cancelar)
                
                # Agregar widgets al layout
                layout.addWidget(resumen)
                layout.addWidget(QLabel("Detalles de las advertencias:"))
                layout.addWidget(error_text)
                layout.addLayout(btn_layout)
                
                error_dialog.setLayout(layout)
                
                if error_dialog.exec() != QDialog.DialogCode.Accepted:
                    return False
            
            if not facturas_importadas:
                QMessageBox.warning(
                    self, 
                    "Ninguna factura válida", 
                    "No se encontraron facturas válidas en el archivo."
                )
                return
            
            # Procesar la importación con la nueva función de ayuda
            return self._procesar_importacion(facturas_importadas, "CSV")
            
        except Exception as e:
            loading_box.close()
            error_msg = (
                f"Error al procesar el archivo CSV:\n\n"
                f"Error: {str(e)}\n\n"
                f"Asegúrese de que el archivo no esté abierto en otro programa y que tenga el formato correcto."
            )
            logger.error(f"Error al importar desde CSV: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error en la importación", error_msg)
            return False
    
    def importar_desde_excel(self):
        """Importar facturas desde un archivo Excel"""
        # Abrir diálogo para seleccionar archivo
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        
        if not file_path:
            return  # Usuario canceló el diálogo
        
        # Mostrar indicador de carga
        loading_box = QMessageBox(self)
        loading_box.setWindowTitle("Cargando archivo")
        loading_box.setText("Procesando archivo Excel, por favor espere...")
        loading_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
        loading_box.show()
        QApplication.processEvents()
        
        try:
            # Cargar el libro de trabajo de Excel
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            # Crear diálogo para seleccionar hoja
            sheet_dialog = QDialog(self)
            sheet_dialog.setWindowTitle("Seleccionar hoja")
            sheet_dialog.setMinimumWidth(300)
            
            layout = QVBoxLayout()
            sheet_dialog.setLayout(layout)
            
            layout.addWidget(QLabel("Seleccione la hoja que contiene los datos:"))
            
            # Lista de hojas disponibles
            sheet_list = QListWidget()
            sheet_list.addItems(wb.sheetnames)
            sheet_list.setCurrentRow(0)  # Seleccionar la primera hoja por defecto
            layout.addWidget(sheet_list)
            
            # Botones
            btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            btn_box.accepted.connect(sheet_dialog.accept)
            btn_box.rejected.connect(sheet_dialog.reject)
            layout.addWidget(btn_box)
            
            if sheet_dialog.exec() != QDialog.DialogCode.Accepted:
                loading_box.close()
                return
            
            # Obtener la hoja seleccionada
            selected_sheet = sheet_list.currentItem().text()
            sheet = wb[selected_sheet]
            
            # Actualizar mensaje de carga
            loading_box.setText(f"Procesando hoja: {selected_sheet}...")
            QApplication.processEvents()
            
            # Obtener los encabezados
            headers = [str(cell.value).strip().lower() if cell.value else '' for cell in sheet[1]]
            
            # Verificar que los encabezados requeridos estén presentes
            required_columns = ['fecha', 'tipo', 'descripcion', 'valor']
            
            # Mapear los índices de las columnas
            column_indices = {}
            missing_columns = []
            
            for col in required_columns:
                try:
                    column_indices[col] = headers.index(col)
                except ValueError:
                    missing_columns.append(col)
            
            if missing_columns:
                loading_box.close()
                QMessageBox.critical(
                    self, 
                    "Error de formato", 
                    f"El archivo Excel debe contener las columnas: {', '.join(required_columns)}.\n"
                    f"Columnas faltantes: {', '.join(missing_columns)}\n\n"
                    f"Columnas encontradas: {', '.join([h for h in headers if h]) or 'Ninguna'}"
                )
                return
            
            # Contar filas para la barra de progreso
            total_rows = sheet.max_row - 1  # Restar la fila de encabezados
            
            # Crear diálogo de progreso para la lectura
            progress = QProgressDialog("Leyendo archivo Excel...", "Cancelar", 0, total_rows, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setMinimumDuration(0)
            progress.setValue(0)
            
            # Leer las facturas
            facturas_importadas = []
            errores = []
            
            for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):  # Empezar desde la fila 2
                if progress.wasCanceled():
                    loading_box.close()
                    return
                
                progress.setValue(i-1)  # Ajustar para empezar desde 0
                QApplication.processEvents()
                
                try:
                    # Obtener los valores de las celdas
                    fecha = str(row[column_indices['fecha']]).strip() if row[column_indices['fecha']] else ""
                    tipo = str(row[column_indices['tipo']]).strip() if row[column_indices['tipo']] else ""
                    descripcion = str(row[column_indices['descripcion']]).strip() if row[column_indices['descripcion']] else ""
                    valor_celda = row[column_indices['valor']]
                    
                    # Validar campos obligatorios
                    if not all([fecha, tipo, descripcion, valor_celda is not None]):
                        errores.append(f"Fila {i}: Faltan campos obligatorios")
                        continue
                    
                    try:
                        # Manejar diferentes formatos de valor
                        valor_str = str(valor_celda).replace('$', '').replace(',', '').strip()
                        valor = float(valor_str)
                        if valor <= 0:
                            raise ValueError("El valor debe ser mayor a cero")
                    except (ValueError, TypeError) as ve:
                        errores.append(f"Fila {i}: Valor inválido - {str(ve)}")
                        continue
                    
                    # Validar formato de fecha (DD/MM/YYYY)
                    try:
                        fecha_dt = datetime.strptime(fecha, '%d/%m/%Y')
                        fecha = fecha_dt.strftime('%d/%m/%Y')  # Estandarizar formato
                    except ValueError:
                        errores.append(f"Fila {i}: Formato de fecha inválido (debe ser DD/MM/YYYY)")
                        continue
                    
                    # Crear objeto de factura
                    factura = {
                        'fecha': fecha,
                        'tipo': tipo,
                        'descripcion': descripcion,
                        'valor': valor,
                        'timestamp': datetime.now().isoformat()
                    }
                    facturas_importadas.append(factura)
                    
                except Exception as e:
                    error_msg = f"Fila {i}: {str(e)}"
                    logger.error(f"Error al procesar fila del Excel: {row}. Error: {error_msg}")
                    errores.append(error_msg)
                    continue
            
            progress.close()
            loading_box.close()
            
            # Mostrar advertencias si hay errores
            if errores:
                error_dialog = QDialog(self)
                error_dialog.setWindowTitle("Advertencias de importación")
                error_dialog.resize(600, 400)
                
                layout = QVBoxLayout()
                
                # Mostrar resumen
                resumen = QLabel(
                    f"Se encontraron {len(errores)} advertencias durante la importación.\n"
                    f"Se importaron {len(facturas_importadas)} facturas correctamente."
                )
                
                # Área de texto para mostrar los errores
                error_text = QTextEdit()
                error_text.setReadOnly(True)
                error_text.setPlainText("\n".join(errores))
                
                # Botón para continuar
                btn_continuar = QPushButton("Continuar con la importación")
                btn_continuar.clicked.connect(error_dialog.accept)
                
                # Botón para cancelar
                btn_cancelar = QPushButton("Cancelar importación")
                btn_cancelar.clicked.connect(error_dialog.reject)
                
                # Layout de botones
                btn_layout = QHBoxLayout()
                btn_layout.addStretch()
                btn_layout.addWidget(btn_continuar)
                btn_layout.addWidget(btn_cancelar)
                
                # Agregar widgets al layout
                layout.addWidget(resumen)
                layout.addWidget(QLabel("Detalles de las advertencias:"))
                layout.addWidget(error_text)
                layout.addLayout(btn_layout)
                
                error_dialog.setLayout(layout)
                
                if error_dialog.exec() != QDialog.DialogCode.Accepted:
                    return False
            
            if not facturas_importadas:
                QMessageBox.warning(
                    self, 
                    "Ninguna factura válida", 
                    "No se encontraron facturas válidas en el archivo."
                )
                return
            
            # Procesar la importación con la nueva función de ayuda
            return self._procesar_importacion(facturas_importadas, "Excel")
            
        except Exception as e:
            loading_box.close()
            error_msg = (
                f"Error al procesar el archivo Excel:\n\n"
                f"Error: {str(e)}\n\n"
                f"Asegúrese de que el archivo no esté abierto en otro programa y que tenga el formato correcto."
            )
            logger.error(f"Error al importar desde Excel: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error en la importación", error_msg)
            return False
    
    def importar_desde_json(self):
        """Importar facturas desde un archivo JSON con diálogo de progreso y validación"""  
        # Abrir diálogo para seleccionar archivo
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo JSON",
            "",
            "Archivos JSON (*.json);;Todos los archivos (*)"
        )
    
        if not file_path:
            return  # Usuario canceló el diálogo
        
        # Mostrar indicador de carga
        loading_box = QMessageBox(self)
        loading_box.setWindowTitle("Cargando archivo")
        loading_box.setText("Procesando archivo JSON, por favor espere...")
        loading_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
        loading_box.show()
        QApplication.processEvents()
        
        try:
            # Leer el archivo JSON
            with open(file_path, 'r', encoding='utf-8') as f:
                facturas_importadas = json.load(f)
            
            loading_box.close()
            
            if not isinstance(facturas_importadas, list):
                QMessageBox.critical(
                    self,
                    "Error",
                    "El archivo JSON debe contener una lista de facturas."
                )
                return
                
            # Validar facturas
            facturas_validas = []
            errores = []
            
            # Crear diálogo de progreso
            progress = QProgressDialog(
                f"Validando {len(facturas_importadas)} facturas...",
                "Cancelar", 0, len(facturas_importadas), self)
            progress.setWindowTitle("Validando facturas")
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setMinimumDuration(500)  # Mostrar después de 500ms
            
            # Validar cada factura
            for i, factura in enumerate(facturas_importadas):
                if progress.wasCanceled():
                    QMessageBox.information(self, "Importación cancelada", 
                                        "La importación ha sido cancelada por el usuario.")
                    return
                    
                progress.setValue(i)
                progress.setLabelText(f"Validando factura {i+1} de {len(facturas_importadas)}...")
                QApplication.processEvents()
                
                try:
                    # Validar campos obligatorios
                    if not all(key in factura for key in ['tipo', 'descripcion', 'valor', 'fecha']):
                        errores.append(f"Factura {i+1}: Faltan campos obligatorios")
                        continue
                    
                    # Validar y convertir tipos de datos
                    try:
                        factura['valor'] = float(factura['valor'])
                        if factura['valor'] <= 0:
                            raise ValueError("El valor debe ser mayor a cero")
                    except (ValueError, TypeError) as e:
                        errores.append(f"Factura {i+1}: Valor inválido - {str(e)}")
                        continue
                    
                    # Validar fecha
                    try:
                        datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    except (ValueError, TypeError):
                        errores.append(f"Factura {i+1}: Formato de fecha inválido. Use DD/MM/AAAA")
                        continue
                    
                    facturas_validas.append(factura)
                    
                except Exception as e:
                    errores.append(f"Factura {i+1}: Error inesperado - {str(e)}")
                    continue
                    
            progress.close()
            
            # Mostrar advertencias si hay errores
            if errores:
                error_dialog = QDialog(self)
                error_dialog.setWindowTitle("Advertencias de importación")
                error_dialog.resize(600, 400)
                
                layout = QVBoxLayout()
                
                label = QLabel(f"Se encontraron {len(errores)} advertencias al importar {len(facturas_importadas)} facturas. "
                            f"Se importarán {len(facturas_validas)} facturas correctas.")
                label.setWordWrap(True)
                layout.addWidget(label)
                
                text_area = QTextEdit()
                text_area.setReadOnly(True)
                text_area.setPlainText("\n".join(errores))
                layout.addWidget(text_area)
                
                button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
                button_box.accepted.connect(error_dialog.accept)
                layout.addWidget(button_box)
                
                error_dialog.setLayout(layout)
                error_dialog.exec()
                
                if not facturas_validas:
                    return  # No hay facturas válidas para importar
            
            # Usar el método _procesar_importacion para manejar la importación
            if facturas_validas:
                # Guardar copia de seguridad
                facturas_originales = self.facturas.copy()
                
                try:
                    # Agregar facturas importadas
                    self.facturas.extend(facturas_validas)
                    
                    # Actualizar interfaz
                    self.actualizar_lista_facturas()
                    
                    # Actualizar resumen solo si el widget existe
                    if hasattr(self, 'actualizar_resumen'):
                        try:
                            self.actualizar_resumen()
                        except AttributeError as e:
                            logger.warning(f"No se pudo actualizar el resumen: {str(e)}")
                    
                    # Guardar datos
                    self.guardar_datos()
                    
                    QMessageBox.information(
                        self,
                        "Importación exitosa",
                        f"Se importaron {len(facturas_validas)} facturas desde {file_path}."
                    )
                    logger.info(f"Se importaron {len(facturas_validas)} facturas desde {file_path}")
                    
                except Exception as e:
                    # Revertir cambios en caso de error
                    self.facturas = facturas_originales
                    self.actualizar_lista_facturas()
                    if hasattr(self, 'actualizar_resumen'):
                        try:
                            self.actualizar_resumen()
                        except:
                            pass
                    
                    error_msg = f"Error al procesar las facturas: {str(e)}"
                    logger.error(error_msg, exc_info=True)
                    QMessageBox.critical(self, "Error", error_msg)
        
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Error", "El archivo seleccionado no es un JSON válido.")
        except Exception as e:
            error_msg = f"Error al importar desde JSON: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def cargar_datos(self, actualizar_ui=True):
        """Cargar datos desde la base de datos
        
        Args:
            actualizar_ui (bool): Si es True, actualiza la interfaz de usuario
        """
        try:
            # Obtener todas las facturas de la base de datos
            self.facturas = self.db.obtener_facturas()
            
            # Actualizar la interfaz si está solicitado y los componentes existen
            if actualizar_ui:
                if hasattr(self, 'tabla_facturas'):
                    self.actualizar_lista_facturas()
                if hasattr(self, 'actualizar_resumen'):
                    self.actualizar_resumen()
            
            logger.info(f"Se cargaron {len(self.facturas)} facturas desde la base de datos")
            return True
            
        except Exception as e:
            error_msg = f"Error al cargar los datos de la base de datos: {str(e)}"
            logger.error(error_msg, exc_info=True)
            if hasattr(self, 'isVisible'):  # Solo mostrar mensaje si la ventana está visible
                QMessageBox.critical(self, "Error", error_msg)
            self.facturas = []
            return False
            
    def guardar_datos(self):
        """Guardar datos en la base de datos"""
        try:
            # Guardar todas las facturas en la base de datos
            if hasattr(self, 'facturas'):
                # Obtener las facturas actuales de la base de datos
                facturas_actuales = {f['id']: f for f in self.db.obtener_facturas()}
                
                # Actualizar o insertar facturas
                for factura in self.facturas:
                    if 'id' in factura and factura['id'] in facturas_actuales:
                        # Actualizar factura existente
                        self.db.actualizar_factura(
                            factura_id=factura['id'],
                            fecha=factura['fecha'],
                            tipo=factura['tipo'],
                            descripcion=factura['descripcion'],
                            valor=factura['valor']
                        )
                        # Eliminar de facturas_actuales para marcar como procesada
                        facturas_actuales.pop(factura['id'], None)
                    else:
                        # Insertar nueva factura
                        factura_id = self.db.agregar_factura(
                            fecha=factura['fecha'],
                            tipo=factura['tipo'],
                            descripcion=factura['descripcion'],
                            valor=factura['valor']
                        )
                        # Actualizar el ID en la factura local
                        factura['id'] = factura_id
                
                # Eliminar facturas que ya no están en self.facturas
                for factura_id in facturas_actuales:
                    self.db.eliminar_factura(factura_id)
                
                logger.info(f"Se guardaron {len(self.facturas)} facturas en la base de datos")
                return True
            return False
            
        except Exception as e:
            error_msg = f"Error al guardar los datos en la base de datos: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
            return False
            if reply == QMessageBox.StandardButton.Cancel:
                return
            
            # Guardar copia de seguridad de los datos actuales
            facturas_originales = self.facturas.copy()
            
            try:
                if reply == QMessageBox.StandardButton.Yes:
                    # Sobrescribir facturas existentes
                    self.facturas = facturas_validas
                else:
                    # Agregar a las facturas existentes
                    self.facturas.extend(facturas_validas)
                
                # Guardar los datos
                self.guardar_datos()
                
                # Actualizar la interfaz
                self.actualizar_tabla_facturas()
                self.actualizar_resumenes()
                
                QMessageBox.information(
                    self, 
                    "Importación exitosa", 
                    f"Se importaron {len(facturas_validas)} facturas desde el archivo JSON."
                )
                
            except Exception as e:
                # En caso de error, restaurar los datos originales
                self.facturas = facturas_originales
                error_msg = f"Error al importar las facturas: {str(e)}"
                logger.error(error_msg, exc_info=True)
                QMessageBox.critical(self, "Error", error_msg)
            
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Error", "El archivo seleccionado no es un JSON válido.")
        except Exception as e:
            error_msg = f"Error al importar desde JSON: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def exportar_a_excel(self):
        """Exportar los datos a un archivo Excel con formato de tabla"""
        if not self.facturas:
            QMessageBox.warning(self, "Exportar a Excel", "No hay datos para exportar.")
            return
        
        # Obtener la configuración
        config = get_config()
        last_dir = config['APP'].get('last_export_dir', str(Path.home() / 'Documents'))
        
        # Nombre de archivo predeterminado con la fecha actual
        default_filename = f"Facturas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        default_path = str(Path(last_dir) / default_filename)
        
        # Solicitar al usuario la ubicación para guardar el archivo
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar como",
            default_path,
            "Archivos Excel (*.xlsx);;Todos los archivos (*)"
        )
        
        if not file_path:
            return  # Usuario canceló el diálogo
        
        # Asegurarse de que el archivo tenga la extensión .xlsx
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'
            
        try:
            # Crear un nuevo libro de Excel
            wb = Workbook()
            
            # Eliminar la hoja por defecto si no tiene datos
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            
            # Estilos
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            money_format = '#,##0.00" COP"'
            
            # 1. Hoja de Detalle (todas las facturas)
            ws_detalle = wb.create_sheet("Todas las Facturas")
            
            # Encabezados
            headers = ["Fecha", "Tipo", "Descripción", "Valor (COP)"]
            for col_num, header in enumerate(headers, 1):
                cell = ws_detalle.cell(row=1, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
            
            # Datos
            for row_num, factura in enumerate(self.facturas, 2):
                ws_detalle.cell(row=row_num, column=1, value=factura['fecha'])
                ws_detalle.cell(row=row_num, column=2, value=factura['tipo'])
                ws_detalle.cell(row=row_num, column=3, value=factura['descripcion'])
                
                # Formato de moneda para la columna de valor
                cell_valor = ws_detalle.cell(row=row_num, column=4, value=float(factura['valor']))
                cell_valor.number_format = money_format
            
            # Crear tabla
            table = Table(displayName="TablaDetalle", ref=f"A1:D{len(self.facturas) + 1}")
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                 showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            table.tableStyleInfo = style
            ws_detalle.add_table(table)
            
            # Ajustar ancho de columnas
            for column in ws_detalle.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        value_length = len(str(cell.value).encode('utf-8'))
                        if value_length > max_length:
                            max_length = value_length
                    except:
                        pass
                adjusted_width = (max_length + 4) * 1.1
                ws_detalle.column_dimensions[column_letter].width = min(adjusted_width, 40)
            
            # 2. Agrupar facturas por mes y año
            facturas_por_mes = {}
            for factura in self.facturas:
                try:
                    # Intentar parsear la fecha en formato dd/mm/yyyy
                    fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    mes_anio = fecha.strftime('%Y-%m')  # Formato: YYYY-MM

                    if mes_anio not in facturas_por_mes:
                        facturas_por_mes[mes_anio] = []
                    facturas_por_mes[mes_anio].append(factura)
                except ValueError as e:
                    logger.warning(f"No se pudo parsear la fecha: {factura['fecha']}. Error: {e}")
                    continue
            
            # 3. Crear una hoja para cada mes
            for mes_anio, facturas_mes in facturas_por_mes.items():
                try:
                    # Crear nombre de hoja en formato 'YYYY-MM Mes' (ej. '2025-08 Agosto')
                    fecha = datetime.strptime(mes_anio, '%Y-%m')
                    nombre_mes = f"{fecha.strftime('%Y-%m')} {fecha.strftime('%B').capitalize()}"

                    # Limitar el nombre de la hoja a 31 caracteres (límite de Excel)
                    nombre_hoja = nombre_mes[:31]

                    # Crear hoja para el mes
                    ws_mes = wb.create_sheet(nombre_hoja)

                    # Título del mes
                    titulo_mes = f"Facturas de {fecha.strftime('%B').capitalize()} {fecha.year}"
                    ws_mes.append([titulo_mes])
                    ws_mes.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
                    titulo_cell = ws_mes.cell(row=1, column=1)
                    titulo_cell.font = Font(size=14, bold=True)
                    titulo_cell.alignment = Alignment(horizontal='center')

                    # Resumen del mes
                    total_mes = sum(f['valor'] for f in facturas_mes)
                    ws_mes.append([f"Total del mes: ${total_mes:,.0f} COP"])
                    ws_mes.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
                    total_cell = ws_mes.cell(row=2, column=1)
                    total_cell.font = Font(bold=True)

                    # Espacio antes de la tabla
                    ws_mes.append([])

                    # Encabezados de la tabla
                    for col_num, header in enumerate(headers, 1):
                        cell = ws_mes.cell(row=4, column=col_num, value=header)
                        cell.font = header_font
                        cell.fill = header_fill

                    # Datos de las facturas del mes
                    # Escribir los datos de las facturas
                    for row_num, factura in enumerate(facturas_mes, 5):
                        ws_mes.cell(row=row_num, column=1, value=factura['fecha'])
                        ws_mes.cell(row=row_num, column=2, value=factura['tipo'])
                        ws_mes.cell(row=row_num, column=3, value=factura['descripcion'])

                        # Formato de moneda para la columna de valor
                        cell_valor = ws_mes.cell(row=row_num, column=4, value=float(factura['valor']))
                        cell_valor.number_format = money_format
                    
                    # Crear tabla para el mes después de escribir todos los datos
                    if facturas_mes:  # Solo crear tabla si hay datos
                        try:
                            # Calcular la última fila de datos (5 + len - 1 ya que empezamos en 5)
                            last_row = 4 + len(facturas_mes)
                            table_ref = f"A4:D{last_row}"  # Rango desde A4 hasta D{última fila}
                            
                            # Crear un nombre de tabla único y válido para Excel
                            table_name = f"Tabla_{mes_anio.replace('-', '_')}"
                            
                            # Asegurarse de que no haya tablas con el mismo nombre
                            if table_name in ws_mes._tables:
                                del ws_mes._tables[table_name]
                            
                            # Crear la tabla
                            table = Table(displayName=table_name, ref=table_ref)
                            
                            # Configurar el estilo de la tabla
                            style = TableStyleInfo(
                                name="TableStyleMedium9",
                                showFirstColumn=False,
                                showLastColumn=False,
                                showRowStripes=True,
                                showColumnStripes=False
                            )
                            table.tableStyleInfo = style
                            
                            # Agregar la tabla a la hoja
                            ws_mes.add_table(table)
                            
                        except Exception as e:
                            logger.warning(f"No se pudo crear la tabla para el mes {mes_anio}: {str(e)}")
                            # Continuar con la ejecución a pesar del error en la tabla
                        
                        # Ajustar ancho de columnas
                        # Primero, obtener el rango de celdas que no están fusionadas
                        data_rows = []
                        for row in ws_mes.iter_rows(min_row=4):  # Empezar después de los títulos
                            if all(cell.value is not None for cell in row):
                                data_rows.append(row)
                        
                        # Si no hay filas de datos, saltar el ajuste
                        if not data_rows:
                            continue
                            
                        # Obtener el número de columnas
                        num_columns = len(data_rows[0])
                        
                        # Ajustar el ancho de cada columna
                        for col_idx in range(num_columns):
                            try:
                                # Obtener la letra de la columna
                                column_letter = get_column_letter(col_idx + 1)
                                max_length = 0
                                
                                # Ajustar para el encabezado (fila 4)
                                header_cell = ws_mes.cell(row=4, column=col_idx + 1)
                                if header_cell and header_cell.value:
                                    max_length = len(str(header_cell.value))
                                
                                # Especificar un ancho mínimo para la columna de valor (columna D)
                                if column_letter == 'D':
                                    max_length = max(max_length, 15)  # Mínimo para valores monetarios
                                
                                # Revisar las celdas de datos
                                for row in data_rows:
                                    cell = row[col_idx]
                                    if cell and cell.value is not None:
                                        # Para celdas con formato de moneda
                                        if hasattr(cell, 'number_format') and 'COP' in str(cell.number_format):
                                            formatted_value = f"{float(cell.value):,.2f} COP"
                                            value_length = len(formatted_value)
                                        else:
                                            value_length = len(str(cell.value))
                                        
                                        if value_length > max_length:
                                            max_length = value_length
                                
                                # Ajustar el ancho con espacio adicional
                                if max_length > 0:
                                    adjusted_width = (max_length + 2) * 1.2
                                    # Limitar el ancho máximo a 50 caracteres
                                    ws_mes.column_dimensions[column_letter].width = min(adjusted_width, 50)
                                    
                                    # Asegurar que la columna de valor tenga un ancho mínimo
                                    if column_letter == 'D':
                                        ws_mes.column_dimensions[column_letter].width = max(
                                            ws_mes.column_dimensions[column_letter].width, 15
                                        )
                                        
                            except Exception as e:
                                logger.warning(f"No se pudo ajustar el ancho de la columna {col_idx + 1}: {str(e)}")
                                continue
                
                except Exception as e:
                    logger.error(f"Error al crear la hoja para el mes {mes_anio}: {str(e)}")
                    continue

            
            # 4. Hoja de Resumen por Tipo de Gasto
            ws_resumen = wb.create_sheet("Resumen por Tipo")
            
            # Calcular totales por tipo
            total_por_tipo = defaultdict(float)
            for factura in self.facturas:
                total_por_tipo[factura['tipo']] += factura['valor']
            
            # Ordenar por valor descendente
            total_por_tipo = dict(sorted(total_por_tipo.items(), key=lambda x: x[1], reverse=True))
            
            # Escribir encabezados
            ws_resumen.append(["Tipo de Gasto", "Total (COP)", "Porcentaje"])
            
            # Calcular total general para porcentajes
            total_general = sum(total_por_tipo.values())
            
            # Escribir datos
            for tipo, total in total_por_tipo.items():
                if total_general > 0:
                    porcentaje = (total / total_general) * 100
                    ws_resumen.append([tipo, total, porcentaje/100])
                else:
                    ws_resumen.append([tipo, total, 0])
            
            # Aplicar formato a la tabla de resumen
            table_resumen = Table(displayName="TablaResumenTipo", ref=f"A1:C{len(total_por_tipo) + 1}")
            style_resumen = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                         showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            table_resumen.tableStyleInfo = style_resumen
            ws_resumen.add_table(table_resumen)
            
            # Formato de moneda para la columna de total
            for row in ws_resumen.iter_rows(min_row=2, min_col=2, max_col=2):
                for cell in row:
                    cell.number_format = money_format
            
            # Formato de porcentaje
            for row in ws_resumen.iter_rows(min_row=2, min_col=3, max_col=3):
                for cell in row:
                    cell.number_format = '0.00%'
            
            # Añadir fila de total
            total_row = len(total_por_tipo) + 2
            ws_resumen.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
            
            cell_total = ws_resumen.cell(row=total_row, column=2, value=float(total_general))
            cell_total.number_format = money_format
            cell_total.font = Font(bold=True)
            
            cell_porcentaje = ws_resumen.cell(row=total_row, column=3, value=1.0 if total_general > 0 else 0)
            cell_porcentaje.number_format = '0.00%'
            cell_porcentaje.font = Font(bold=True)
            
            # Ajustar ancho de columnas
            for column in ws_resumen.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        # Considerar el ancho del encabezado también
                        value_length = len(str(cell.value).encode('utf-8'))  # Contar bytes para caracteres especiales
                        if value_length > max_length:
                            max_length = value_length
                    except:
                        pass
                # Añadir un poco más de espacio para el borde y padding
                adjusted_width = (max_length + 4) * 1.1
                ws_resumen.column_dimensions[column_letter].width = min(adjusted_width, 50)  # Aumentado el ancho máximo a 50
            
            # 5. Hoja de Resumen Mensual
            ws_mensual = wb.create_sheet("Resumen Mensual")
            
            # Calcular totales por mes y año
            total_por_mes = defaultdict(float)
            for factura in self.facturas:
                try:
                    fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    mes_anio = fecha.strftime('%Y-%m')
                    total_por_mes[mes_anio] += factura['valor']
                except Exception as e:
                    logger.warning(f"Error al procesar fecha: {factura['fecha']} - {str(e)}")
                    continue
            
            # Ordenar por fecha
            total_por_mes = dict(sorted(total_por_mes.items()))
            
            # Escribir encabezados
            ws_mensual.append(["Mes", "Año", "Total (COP)"])
            
            # Escribir datos
            for mes_anio, total in total_por_mes.items():
                anio, mes = mes_anio.split('-')
                nombre_mes = [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ][int(mes) - 1]
                ws_mensual.append([nombre_mes, anio, total])
            
            # Aplicar formato a la tabla de resumen mensual
            if len(total_por_mes) > 0:
                table_mensual = Table(displayName="TablaResumenMensual", ref=f"A1:C{len(total_por_mes) + 1}")
                style_mensual = TableStyleInfo(name="TableStyleMedium3", showFirstColumn=False,
                                             showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                table_mensual.tableStyleInfo = style_mensual
                ws_mensual.add_table(table_mensual)
                
                # Formato de moneda para la columna de total
                for row in ws_mensual.iter_rows(min_row=2, min_col=3, max_col=3):
                    for cell in row:
                        cell.number_format = money_format
                
                # Ajustar ancho de columnas
                for column in ws_mensual.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            # Considerar el ancho del encabezado también
                            value_length = len(str(cell.value).encode('utf-8'))  # Contar bytes para caracteres especiales
                            if value_length > max_length:
                                max_length = value_length
                        except:
                            pass
                    # Añadir un poco más de espacio para el borde y padding
                    adjusted_width = (max_length + 4) * 1.1
                    ws_mensual.column_dimensions[column_letter].width = min(adjusted_width, 40)  # Ajustado el ancho máximo
            
            # Guardar el archivo con configuración para evitar advertencias
            temp_file = file_path + '.tmp'
            try:
                # Guardar primero en un archivo temporal
                wb.save(temp_file)
                
                # Si el archivo de destino existe, eliminarlo
                if os.path.exists(file_path):
                    os.remove(file_path)
                
                # Renombrar el archivo temporal al nombre final
                os.rename(temp_file, file_path)
                
                # Actualizar el directorio de exportación en la configuración
                config['APP']['last_export_dir'] = str(Path(file_path).parent)
                save_config(config)
                
                # Mostrar mensaje de éxito
                QMessageBox.information(
                    self,
                    "Exportación exitosa",
                    f"Los datos se han exportado correctamente a:\n{file_path}"
                )
                
                logger.info(f"Datos exportados exitosamente a {file_path}")
                
            except Exception as e:
                # Si hay un error, intentar limpiar el archivo temporal
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                raise e
            
        except PermissionError:
            error_msg = "No se pudo guardar el archivo. Asegúrese de que el archivo no esté abierto en otro programa."
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error de permisos", error_msg)
            
        except Exception as e:
            error_msg = f"Ocurrió un error al exportar a Excel: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(
                self,
                "Error al exportar",
                f"{error_msg}\n\nPor favor, revise los logs para más detalles."
            )
            # Encabezados
            headers = ["Tipo de Gasto", "Total (COP)", "Porcentaje"]
            for col_num, header in enumerate(headers, 1):
                cell = ws_resumen.cell(row=1, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
            
            # Datos
            total_general = sum(total_por_tipo.values())
            for row_num, (tipo, total) in enumerate(total_por_tipo.items(), 2):
                ws_resumen.cell(row=row_num, column=1, value=tipo)
                
                # Formato de moneda para el total
                cell_total = ws_resumen.cell(row=row_num, column=2, value=float(total))
                cell_total.number_format = money_format
                
                # Porcentaje
                if total_general > 0:
                    porcentaje = (total / total_general) * 100
                    cell_porcentaje = ws_resumen.cell(row=row_num, column=3, value=porcentaje/100)
                    cell_porcentaje.number_format = '0.00%'
            
            # Fila de total
            row_num = len(total_por_tipo) + 2
            ws_resumen.cell(row=row_num, column=1, value="TOTAL").font = Font(bold=True)
            
            cell_total = ws_resumen.cell(row=row_num, column=2, value=float(total_general))
            cell_total.number_format = money_format
            cell_total.font = Font(bold=True)
            
            # Ajustar ancho de columnas
            for col in ws_resumen.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.1
                ws_resumen.column_dimensions[column].width = min(adjusted_width, 30)
            
            # 3. Hoja de Resumen Mensual
            ws_mensual = wb.create_sheet("Resumen Mensual")
            
            # Calcular totales por mes y tipo
            total_por_mes_tipo = defaultdict(lambda: defaultdict(float))
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            
            for factura in self.facturas:
                try:
                    fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y')
                    mes = fecha.month - 1  # 0-11
                    total_por_mes_tipo[mes][factura['tipo']] += factura['valor']
                except:
                    continue
            
            # Obtener todos los tipos únicos
            tipos = sorted({tipo for factura in self.facturas for tipo in [factura['tipo']]})
            
            # Encabezados
            headers = ["Mes"] + tipos + ["Total"]
            for col_num, header in enumerate(headers, 1):
                cell = ws_mensual.cell(row=1, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
            
            # Datos por mes
            for mes in range(12):
                if mes in total_por_mes_tipo:
                    row_num = mes + 2  # +2 porque la fila 1 es el encabezado
                    ws_mensual.cell(row=row_num, column=1, value=meses[mes])
                    
                    # Valores por tipo
                    total_mes = 0
                    for col_num, tipo in enumerate(tipos, 2):  # Empieza en columna 2
                        valor = total_por_mes_tipo[mes].get(tipo, 0)
                        if valor > 0:
                            cell = ws_mensual.cell(row=row_num, column=col_num, value=float(valor))
                            cell.number_format = money_format
                            total_mes += valor
                    
                    # Total del mes
                    cell_total = ws_mensual.cell(row=row_num, column=len(tipos)+2, value=float(total_mes))
                    cell_total.number_format = money_format
            
            # Fila de totales por tipo
            row_num = 14  # Después de los 12 meses
            ws_mensual.cell(row=row_num, column=1, value="TOTAL").font = Font(bold=True)
            
            # Calcular totales por tipo
            for col_num, tipo in enumerate(tipos, 2):
                total_tipo = sum(total_por_mes_tipo[mes].get(tipo, 0) for mes in range(12))
                if total_tipo > 0:
                    cell = ws_mensual.cell(row=row_num, column=col_num, value=float(total_tipo))
                    cell.number_format = money_format
                    cell.font = Font(bold=True)
            
            # Total general
            cell_total = ws_mensual.cell(row=row_num, column=len(tipos)+2, 
                                       value=float(sum(sum(m.values()) for m in total_por_mes_tipo.values())))
            cell_total.number_format = money_format
            cell_total.font = Font(bold=True)
            
            # Ajustar ancho de columnas
            for col in ws_mensual.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.1
                ws_mensual.column_dimensions[column].width = min(adjusted_width, 20)
            
            # Guardar el archivo
            wb.save(file_path)
            
            # Mostrar mensaje de éxito
            QMessageBox.information(
                self, 
                "Exportación exitosa", 
                f"Los datos se han exportado correctamente a:\n{file_path}"
            )
            
            logger.info(f"Datos exportados exitosamente a {file_path}")
            
        except Exception as e:
            error_msg = f"Error al exportar a Excel: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def actualizar_boton_eliminar(self):
        """Actualizar el estado del botón de eliminar basado en la selección"""
        seleccionados = self.tabla_facturas.selectedItems()
        self.btn_eliminar.setEnabled(len(seleccionados) > 0)
    
    def actualizar_lista_facturas(self):
        """Actualizar la tabla de facturas"""
        self.tabla_facturas.setRowCount(len(self.facturas))
        
        for i, factura in enumerate(self.facturas):
            # Agregar un elemento oculto para mantener el índice
            self.tabla_facturas.setItem(i, 0, QTableWidgetItem(str(i)))  # Guardar el índice
            
            # Agregar los datos de la factura
            self.tabla_facturas.setItem(i, 1, QTableWidgetItem(factura['fecha']))
            self.tabla_facturas.setItem(i, 2, QTableWidgetItem(factura['tipo']))
            self.tabla_facturas.setItem(i, 3, QTableWidgetItem(factura['descripcion']))
            self.tabla_facturas.setItem(i, 4, QTableWidgetItem(f"${factura['valor']:,.0f} COP".replace(',', '.')))
            
            # Resaltar filas con valores altos
            if factura['valor'] > 1000000:  # Resaltar facturas mayores a 1 millón
                for col in range(1, 5):  # Solo las columnas visibles
                    item = self.tabla_facturas.item(i, col)
                    if item:
                        item.setBackground(QColor(255, 230, 230))  # Rojo claro
    
    def eliminar_facturas_seleccionadas(self):
        """Eliminar las facturas seleccionadas de la lista"""
        # Obtener las filas seleccionadas (sin duplicados)
        filas_seleccionadas = set()
        for item in self.tabla_facturas.selectedItems():
            filas_seleccionadas.add(item.row())
        
        if not filas_seleccionadas:
            QMessageBox.warning(self, "Eliminar Facturas", "No hay facturas seleccionadas para eliminar.")
            return
        
        # Confirmar eliminación
        confirmacion = QMessageBox.question(
            self,
            "Confirmar Eliminación",
            f"¿Está seguro de que desea eliminar {len(filas_seleccionadas)} factura(s) seleccionada(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if confirmacion == QMessageBox.StandardButton.Yes:
            # Ordenar las filas de mayor a menor para evitar problemas con los índices al eliminar
            filas_ordenadas = sorted(filas_seleccionadas, reverse=True)
            
            # Eliminar las facturas de la lista
            for fila in filas_ordenadas:
                if 0 <= fila < len(self.facturas):
                    self.facturas.pop(fila)
            
            # Guardar los cambios
            if self.guardar_datos():
                # Actualizar la interfaz
                self.actualizar_lista_facturas()
                self.actualizar_resumen()
                self.statusBar().showMessage(f"Se eliminaron {len(filas_seleccionadas)} factura(s) correctamente.", 3000)
                logger.info(f"Se eliminaron {len(filas_seleccionadas)} facturas")
    
    def confirmar_limpiar_todo(self):
        """Mostrar diálogo de confirmación para limpiar todos los datos"""
        if not self.facturas:
            QMessageBox.information(self, "Limpiar Datos", "No hay datos para limpiar.")
            return
        
        confirmacion = QMessageBox.warning(
            self,
            "Confirmar Limpieza Total",
            "¿Está seguro de que desea eliminar TODAS las facturas?\n\n"
            "Esta acción NO se puede deshacer.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if confirmacion == QMessageBox.StandardButton.Yes:
            self.limpiar_todo()
    
    def limpiar_todo(self):
        """Eliminar todas las facturas"""
        # Guardar una copia de respaldo
        facturas_backup = self.facturas.copy()
        
        try:
            # Limpiar la lista de facturas
            self.facturas.clear()
            
            # Guardar los cambios
            if self.guardar_datos():
                # Actualizar la interfaz
                self.actualizar_lista_facturas()
                self.actualizar_resumen()
                self.statusBar().showMessage("Se eliminaron todas las facturas correctamente.", 3000)
                logger.info("Se eliminaron todas las facturas")
        except Exception as e:
            # Restaurar la copia de respaldo en caso de error
            self.facturas = facturas_backup
            error_msg = f"Error al limpiar los datos: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def importar_desde_json(self):
        """Importar facturas desde un archivo JSON"""
        # Abrir diálogo para seleccionar archivo
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo JSON",
            "",
            "Archivos JSON (*.json);;Todos los archivos (*)"
        )
        
        if not file_path:
            return  # Usuario canceló el diálogo
        
        try:
            # Leer el archivo JSON
            with open(file_path, 'r', encoding='utf-8') as f:
                facturas_importadas = json.load(f)
            
            if not isinstance(facturas_importadas, list):
                raise ValueError("El archivo no contiene una lista de facturas válida.")
            
            # Validar el formato de las facturas
            facturas_validas = []
            for i, factura in enumerate(facturas_importadas, 1):
                if not all(key in factura for key in ['fecha', 'tipo', 'descripcion', 'valor']):
                    logger.warning(f"Factura {i} omitida: formato incorrecto")
                    continue
                
                try:
                    # Validar y convertir los valores
                    fecha_valida = bool(datetime.strptime(factura['fecha'], '%d/%m/%Y'))
                    valor_valido = float(factura['valor']) > 0
                    
                    if fecha_valida and valor_valido:
                        # Asegurar que el valor sea un número
                        factura['valor'] = float(factura['valor'])
                        facturas_validas.append(factura)
                    else:
                        logger.warning(f"Factura {i} omitida: fecha o valor inválidos")
                except (ValueError, TypeError):
                    logger.warning(f"Factura {i} omitida: error en los datos")
            
            if not facturas_validas:
                QMessageBox.warning(self, "Importar Datos", "No se encontraron facturas válidas para importar.")
                return
            
            # Mostrar resumen de importación
            resumen = f"Se importarán {len(facturas_validas)} factura(s) de {len(facturas_importadas)} encontradas."
            if len(facturas_validas) < len(facturas_importadas):
                resumen += "\n\nAlgunas facturas se omitieron por tener formato incorrecto."
            
            confirmacion = QMessageBox.question(
                self,
                "Confirmar Importación",
                f"{resumen}\n\n¿Desea continuar con la importación?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if confirmacion == QMessageBox.StandardButton.Yes:
                # Hacer una copia de respaldo
                facturas_originales = self.facturas.copy()
                
                try:
                    # Agregar las facturas importadas a la lista existente
                    self.facturas.extend(facturas_validas)
                    
                    # Guardar los cambios
                    if self.guardar_datos():
                        # Actualizar la interfaz
                        self.actualizar_lista_facturas()
                        self.actualizar_resumen()
                        self.actualizar_filtros()
                        
                        # Mostrar mensaje de éxito
                        QMessageBox.information(
                            self,
                            "Importación Exitosa",
                            f"Se importaron {len(facturas_validas)} factura(s) correctamente."
                        )
                        
                        logger.info(f"Se importaron {len(facturas_validas)} facturas desde {file_path}")
                        self.statusBar().showMessage("Importación completada correctamente.", 3000)
                except Exception as e:
                    # Restaurar la copia de respaldo en caso de error
                    self.facturas = facturas_originales
                    error_msg = f"Error al importar las facturas: {str(e)}"
                    logger.error(error_msg, exc_info=True)
                    QMessageBox.critical(self, "Error", error_msg)
        
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Error", "El archivo seleccionado no es un JSON válido.")
        except Exception as e:
            error_msg = f"Error al importar desde JSON: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def cargar_preferencia_tema(self):
        """Cargar la preferencia de tema desde el archivo de configuración"""
        config_path = "config.json"
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    if 'tema_oscuro' in config:
                        self.tema_oscuro = config['tema_oscuro']
                        self.aplicar_tema()
        except Exception as e:
            logger.warning(f"No se pudo cargar la preferencia de tema: {str(e)}")
    
    def guardar_preferencia_tema(self):
        """Guardar la preferencia de tema en el archivo de configuración"""
        config_path = "config.json"
        try:
            config = {}
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    try:
                        config = json.load(f)
                    except json.JSONDecodeError:
                        config = {}
            
            config['tema_oscuro'] = self.tema_oscuro
            
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            logger.error(f"Error al guardar la preferencia de tema: {str(e)}")
    
    def cambiar_tema(self):
        """Alternar entre modo oscuro y claro"""
        self.tema_oscuro = not self.tema_oscuro
        # Actualizar el texto del botón según el tema
        self.btn_tema.setText("Modo Claro" if self.tema_oscuro else "Modo Oscuro")
        self.aplicar_tema()
        self.guardar_preferencia_tema()
    
    def aplicar_tema(self):
        """Aplicar el tema seleccionado a toda la aplicación"""
        if self.tema_oscuro:
            # Estilo para modo oscuro (negros y rosados)
            self.setStyleSheet("""
                /* Estilos generales */
                QMainWindow, QDialog, QWidget, QTabWidget::pane, QTabBar::tab:selected {
                    background-color: #121212;
                    color: #f8f9fa;
                }
                
                /* Pestañas */
                QTabBar::tab {
                    background: #1e1e1e;
                    color: #f8f9fa;
                    padding: 8px 20px;
                    border: 1px solid #ff4081;
                    border-bottom: none;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                    margin-right: 2px;
                }
                
                QTabBar::tab:selected {
                    background: #ff4081;
                    color: #121212;
                    font-weight: bold;
                    border-bottom: 1px solid #ff4081;
                }
                
                QTabBar::tab:!selected {
                    margin-top: 2px;
                    background: #2a2a2a;
                }
                
                /* Botones */
                QPushButton {
                    background-color: #2a2a2a;
                    color: #f8f9fa;
                    border: 1px solid #ff4081;
                    padding: 5px 15px;
                    border-radius: 4px;
                }
                
                QPushButton:hover {
                    background-color: #ff4081;
                    color: #121212;
                }
                
                QPushButton:pressed {
                    background-color: #c2185b;
                }
                
                /* Campos de entrada */
                QLineEdit, QTextEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox {
                    background-color: #1e1e1e;
                    color: #f8f9fa;
                    border: 1px solid #ff4081;
                    padding: 5px;
                    border-radius: 4px;
                    selection-background-color: #ff4081;
                }
                
                /* Tablas */
                QTableWidget {
                    background-color: #121212;
                    color: #f8f9fa;
                    gridline-color: #2a2a2a;
                    border: 1px solid #2a2a2a;
                    alternate-background-color: #1a1a1a;
                }
                
                QTableWidget::item {
                    padding: 5px;
                    border-bottom: 1px solid #2a2a2a;
                }
                
                QTableWidget::item:selected {
                    background-color: #ff4081;
                    color: #121212;
                }
                
                QHeaderView::section {
                    background-color: #1e1e1e;
                    color: #f8f9fa;
                    padding: 5px;
                    border: 1px solid #2a2a2a;
                    border-top: none;
                    border-bottom: 2px solid #ff4081;
                }
                
                /* Grupos */
                QGroupBox {
                    border: 1px solid #ff4081;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 15px;
                    color: #f8f9fa;
                }
                
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px;
                    color: #ff80ab;
                }
                
                /* Barra de estado */
                QStatusBar {
                    background-color: #1e1e1e;
                    color: #f8f9fa;
                    border-top: 1px solid #2a2a2a;
                }
                
                /* Diálogos */
                QMessageBox {
                    background-color: #121212;
                    color: #f8f9fa;
                }
                
                QMessageBox QLabel {
                    color: #f8f9fa;
                }
                
                QMessageBox QPushButton {
                    min-width: 80px;
                    background-color: #2a2a2a;
                    color: #f8f9fa;
                    border: 1px solid #ff4081;
                }
                
                QMessageBox QPushButton:hover {
                    background-color: #ff4081;
                    color: #121212;
                }
                
                /* Botones de acción */
                QPushButton#btn_limpiar_todo, QPushButton#btn_eliminar {
                    background-color: #c2185b;
                    color: white;
                    border: 1px solid #9c1352;
                }
                
                QPushButton#btn_limpiar_todo:hover, QPushButton#btn_eliminar:hover {
                    background-color: #e91e63;
                }
                
                QPushButton#btn_limpiar_todo:pressed, QPushButton#btn_eliminar:pressed {
                    background-color: #880e4f;
                }
                
                /* Resaltado de filas */
                QTableWidget::item[valor_alto="true"] {
                    background-color: #4a1a2a;
                    color: #ff80ab;
                }
                
                /* Estilos para menús desplegables */
                QMenu {
                    background-color: #1e1e1e;
                    color: #f8f9fa;
                    border: 1px solid #ff4081;
                    padding: 5px;
                }
                
                QMenu::item {
                    padding: 5px 15px;
                }
                
                QMenu::item:selected {
                    background-color: #ff4081;
                    color: #121212;
                }
                
                QMenu::item:disabled {
                    color: #6c757d;
                }
            """)
            # Cambiar ícono a luna (modo oscuro activado)
            self.btn_tema.setIcon(self.style().standardIcon(getattr(QStyle.StandardPixmap, 'SP_TitleBarNormalButton')))
        else:
            # Estilo para modo claro (blancos y azules)
            self.setStyleSheet("""
                /* Estilos generales */
                QMainWindow, QDialog, QWidget, QTabWidget::pane, QTabBar::tab:selected {
                    background-color: #f8f9fa;
                    color: #212529;
                }
                
                /* Pestañas */
                QTabBar::tab {
                    background: #e9ecef;
                    color: #495057;
                    padding: 8px 20px;
                    border: 1px solid #0d6efd;
                    border-bottom: none;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                    margin-right: 2px;
                }
                
                QTabBar::tab:selected {
                    background: #0d6efd;
                    color: white;
                    font-weight: bold;
                    border-bottom: 1px solid #0d6efd;
                }
                
                QTabBar::tab:!selected {
                    margin-top: 2px;
                    background: #e9ecef;
                }
                
                /* Botones */
                QPushButton {
                    background-color: #e9ecef;
                    color: #212529;
                    border: 1px solid #ced4da;
                    padding: 5px 15px;
                    border-radius: 4px;
                }
                
                QPushButton:hover {
                    background-color: #0d6efd;
                    color: white;
                    border-color: #0b5ed7;
                }
                
                QPushButton:pressed {
                    background-color: #0b5ed7;
                }
                
                /* Campos de entrada */
                QLineEdit, QTextEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox {
                    background-color: white;
                    color: #212529;
                    border: 1px solid #ced4da;
                    padding: 5px;
                    border-radius: 4px;
                    selection-background-color: #0d6efd;
                    selection-color: white;
                }
                
                /* Tablas */
                QTableWidget {
                    background-color: white;
                    color: #212529;
                    gridline-color: #dee2e6;
                    border: 1px solid #dee2e6;
                    alternate-background-color: #f8f9fa;
                }
                
                QTableWidget::item {
                    padding: 5px;
                    border-bottom: 1px solid #dee2e6;
                }
                
                QTableWidget::item:selected {
                    background-color: #0d6efd;
                    color: white;
                }
                
                QHeaderView::section {
                    background-color: #f1f3f5;
                    color: #212529;
                    padding: 5px;
                    border: 1px solid #dee2e6;
                    border-top: none;
                    border-bottom: 2px solid #0d6efd;
                }
                
                /* Grupos */
                QGroupBox {
                    border: 1px solid #dee2e6;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 15px;
                    color: #212529;
                }
                
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px;
                    color: #0d6efd;
                }
                
                /* Barra de estado */
                QStatusBar {
                    background-color: #e9ecef;
                    color: #212529;
                    border-top: 1px solid #dee2e6;
                }
                
                /* Diálogos */
                QMessageBox {
                    background-color: white;
                    color: #212529;
                }
                
                QMessageBox QLabel {
                    color: #212529;
                }
                
                QMessageBox QPushButton {
                    min-width: 80px;
                    background-color: #e9ecef;
                    color: #212529;
                    border: 1px solid #ced4da;
                }
                
                QMessageBox QPushButton:hover {
                    background-color: #0d6efd;
                    color: white;
                }
                
                /* Botones de acción */
                QPushButton#btn_limpiar_todo, QPushButton#btn_eliminar {
                    background-color: #dc3545;
                    color: white;
                    border: 1px solid #bb2d3b;
                }
                
                QPushButton#btn_limpiar_todo:hover, QPushButton#btn_eliminar:hover {
                    background-color: #bb2d3b;
                }
                
                QPushButton#btn_limpiar_todo:pressed, QPushButton#btn_eliminar:pressed {
                    background-color: #b02a37;
                }
                
                /* Resaltado de filas */
                QTableWidget::item[valor_alto="true"] {
                    background-color: #fff3bf;
                    color: #212529;
                }
            """)
            # Cambiar ícono a sol (modo claro activado)
            self.btn_tema.setIcon(self.style().standardIcon(getattr(QStyle.StandardPixmap, 'SP_TitleBarMaxButton')))
        
        # Actualizar el tooltip del botón
        modo = "oscuro" if self.tema_oscuro else "claro"
        self.btn_tema.setToolTip(f"Cambiar a modo {'claro' if self.tema_oscuro else 'oscuro'}")
        
        # Forzar actualización de la interfaz
        self.update()
        for widget in self.findChildren(QWidget):
            widget.update()

def main():
    app = QApplication(sys.argv)
    
    # Establecer estilo visual
    app.setStyle('Fusion')
    
    # Crear y mostrar ventana principal
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()