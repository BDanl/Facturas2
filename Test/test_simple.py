import sqlite3
import logging
import traceback

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_database():
    """Prueba simple de la base de datos."""
    try:
        # Crear una base de datos en memoria
        logger.info("Creando base de datos en memoria...")
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row  # Para acceder a las columnas por nombre
        cursor = conn.cursor()
        
        # Crear la tabla tipos_gasto
        logger.info("Creando tabla tipos_gasto...")
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS tipos_gasto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            descripcion TEXT,
            color TEXT
        )''')
        
        # Insertar datos de prueba
        logger.info("Insertando datos de prueba...")
        tipos_prueba = [
            ('Mercado', 'Compras de supermercado', '#FF6B6B'),
            ('Transporte', 'Transporte público/privado', '#4ECDC4'),
            ('Entretenimiento', 'Cine, salidas, etc.', '#45B7D1')
        ]
        
        cursor.executemany(
            'INSERT INTO tipos_gasto (nombre, descripcion, color) VALUES (?, ?, ?)',
            tipos_prueba
        )
        conn.commit()
        
        # Consultar los datos
        logger.info("Consultando datos...")
        cursor.execute('SELECT * FROM tipos_gasto ORDER BY nombre')
        filas = cursor.fetchall()
        
        # Mostrar resultados
        logger.info(f"Se encontraron {len(filas)} tipos de gasto:")
        for fila in filas:
            logger.info(f"- ID: {fila[0]}, Nombre: {fila[1]}, Color: {fila[3]}")
        
        # Cerrar la conexión
        conn.close()
        logger.info("Prueba completada exitosamente.")
        return True
        
    except Exception as e:
        logger.error(f"Error en la prueba: {str(e)}")
        logger.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    if test_database():
        logger.info("¡Todas las pruebas se completaron exitosamente!")
    else:
        logger.error("Algunas pruebas fallaron.")
        exit(1)
