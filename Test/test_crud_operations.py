import sys
import logging
import traceback
import sqlite3
from datetime import datetime, timedelta
from database import Database

# Enable SQLite query logging
def trace_callback(query):
    logger.debug(f"SQL: {query}")
    return query

# Configure SQLite to log all queries
sqlite3.enable_callback_tracebacks(True)

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,  # Cambiado a DEBUG para más detalles
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('test_debug_output.txt', mode='w', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

# Configurar el nivel de logging para SQLite
logging.getLogger('sqlite3').setLevel(logging.DEBUG)

def print_header(title):
    """Imprime un encabezado para las pruebas."""
    logger.info("\n" + "="*80)
    logger.info(f" {title} ".center(80, '='))
    logger.info("="*80 + "\n")

def log_sql_queries(func):
    """Decorador para registrar consultas SQL."""
    def wrapper(*args, **kwargs):
        logger.debug(f"Ejecutando consulta SQL en {func.__name__}")
        try:
            result = func(*args, **kwargs)
            logger.debug(f"Consulta SQL exitosa en {func.__name__}")
            return result
        except Exception as e:
            logger.error(f"Error en consulta SQL en {func.__name__}: {str(e)}")
            logger.error(traceback.format_exc())
            raise
    return wrapper

class DatabaseWrapper:
    """Wrapper para la clase Database con logging mejorado."""
    
    def __init__(self, db):
        self.db = db
    
    def __getattr__(self, name):
        attr = getattr(self.db, name)
        if callable(attr):
            return log_sql_queries(attr)
        return attr

def test_crud_operations():
    """Prueba las operaciones CRUD en la base de datos."""
    db = None
    try:
        # 1. Inicialización
        print_header("1. INICIALIZACIÓN DE LA BASE DE DATOS")
        logger.info("Creando instancia de Database...")
        db = DatabaseWrapper(Database(":memory:"))
        logger.info("✓ Base de datos en memoria inicializada correctamente")
        
        # Verificar que las tablas se crearon correctamente
        with db._get_connection() as conn:
            cursor = conn.cursor()
            logger.debug("Ejecutando consulta para listar tablas...")
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tablas = cursor.fetchall()
            logger.info(f"Tablas en la base de datos: {[t[0] for t in tablas] if tablas else 'Ninguna'}")
            
            if tablas and any('tipos_gasto' in t for t in tablas[0]):
                logger.debug("La tabla 'tipos_gasto' existe, obteniendo su estructura...")
                cursor.execute("PRAGMA table_info(tipos_gasto)")
                columnas = cursor.fetchall()
                logger.info("Estructura de la tabla 'tipos_gasto':")
                for col in columnas:
                    logger.info(f"  {col[1]} ({col[2]})")
            else:
                logger.error("✗ La tabla 'tipos_gasto' no existe")
                # Intentar crear la tabla manualmente para diagnóstico
                try:
                    logger.info("Intentando crear la tabla 'tipos_gasto' manualmente...")
                    cursor.execute('''
                    CREATE TABLE IF NOT EXISTS tipos_gasto (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nombre TEXT UNIQUE NOT NULL,
                        descripcion TEXT,
                        color TEXT
                    )''')
                    conn.commit()
                    logger.info("✓ Tabla 'tipos_gasto' creada manualmente")
                except Exception as e:
                    logger.error(f"✗ Error al crear la tabla manualmente: {str(e)}")
                    raise
        
        # 2. Prueba de creación (Create)
        print_header("2. PRUEBA DE CREACIÓN (CREATE)")
        
        # 2.1 Agregar tipos de gasto
        logger.info("Agregando tipos de gasto...")
        tipos_esperados = [
            ("Mercado", "Compras de supermercado", "#FF6B6B"),
            ("Transporte", "Transporte público/privado", "#4ECDC4"),
            ("Entretenimiento", "Cine, salidas, etc.", "#45B7D1")
        ]
        
        for nombre, descripcion, color in tipos_esperados:
            db.agregar_factura(
                fecha="01/01/2025",
                tipo=nombre,
                descripcion=f"{descripcion} - prueba",
                valor=1000.0
            )
        logger.info("✓ Tipos de gasto agregados a través de facturas")
        
        # 2.2 Verificar tipos de gasto
        tipos = db.obtener_tipos_gasto()
        logger.info(f"Tipos de gasto en la base de datos: {len(tipos)}")
        for i, tipo in enumerate(tipos, 1):
            logger.info(f"  {i}. {tipo['nombre']} - {tipo['descripcion']}")
        
        # 2.3 Agregar facturas de prueba
        logger.info("\nAgregando facturas de prueba...")
        factura_ids = []
        for i in range(1, 6):
            fecha = (datetime.now() - timedelta(days=i)).strftime("%d/%m/%Y")
            factura_id = db.agregar_factura(
                fecha=fecha,
                tipo=tipos_esperados[i % len(tipos_esperados)][0],
                descripcion=f"Factura de prueba {i}",
                valor=1000.0 * i
            )
            factura_ids.append(factura_id)
            logger.info(f"  Factura {i} agregada con ID: {factura_id}")
        
        # 3. Prueba de lectura (Read)
        print_header("3. PRUEBA DE LECTURA (READ)")
        
        # 3.1 Obtener todas las facturas
        facturas = db.obtener_facturas()
        logger.info(f"Total de facturas en la base de datos: {len(facturas)}")
        for i, factura in enumerate(facturas, 1):
            logger.info(f"  {i}. ID: {factura['id']}, Fecha: {factura['fecha']}, "
                      f"Tipo: {factura['tipo']}, Valor: ${factura['valor']:,.2f}")
        
        # 3.2 Obtener resumen por tipo
        logger.info("\nObteniendo resumen por tipo de gasto...")
        try:
            resumen_tipos = db.obtener_resumen_por_tipo()
            if not resumen_tipos:
                logger.warning("No se encontraron datos para el resumen por tipo")
            else:
                for i, resumen in enumerate(resumen_tipos, 1):
                    tipo = resumen.get('tipo', 'Sin tipo')
                    total = resumen.get('total', 0) or 0  # Handle None values
                    cantidad = resumen.get('cantidad', 0)
                    logger.info(f"  {i}. {tipo}: ${total:,.2f} ({cantidad} facturas)")
        except Exception as e:
            logger.error(f"Error al obtener resumen por tipo: {str(e)}")
            logger.error(traceback.format_exc())
        
        # 3.3 Obtener resumen mensual
        logger.info("\nObteniendo resumen mensual...")
        try:
            resumen_mensual = db.obtener_resumen_mensual()
            if not resumen_mensual:
                logger.warning("No se encontraron datos para el resumen mensual")
            else:
                for i, resumen in enumerate(resumen_mensual, 1):
                    mes = resumen.get('mes', 'Sin fecha')
                    total = resumen.get('total', 0) or 0  # Handle None values
                    logger.info(f"  {i}. {mes}: ${total:,.2f}")
        except Exception as e:
            logger.error(f"Error al obtener resumen mensual: {str(e)}")
            logger.error(traceback.format_exc())
        
        # 4. Prueba de actualización (Update)
        print_header("4. PRUEBA DE ACTUALIZACIÓN (UPDATE)")
        
        if facturas:
            factura_id = facturas[0]['id']
            logger.info(f"Actualizando factura ID: {factura_id}")
            
            # Actualizar la factura
            actualizado = db.actualizar_factura(
                factura_id=factura_id,
                fecha="31/12/2024",
                tipo=tipos_esperados[0][0],
                descripcion="Factura actualizada",
                valor=9999.99
            )
            
            if actualizado:
                logger.info("✓ Factura actualizada correctamente")
                
                # Verificar la actualización
                factura_actualizada = next((f for f in db.obtener_facturas() 
                                          if f['id'] == factura_id), None)
                if factura_actualizada:
                    logger.info(f"  Fecha: {factura_actualizada['fecha']}")
                    logger.info(f"  Descripción: {factura_actualizada['descripcion']}")
                    logger.info(f"  Valor: ${factura_actualizada['valor']:,.2f}")
            else:
                logger.error("✗ No se pudo actualizar la factura")
        
        # 5. Prueba de eliminación (Delete)
        print_header("5. PRUEBA DE ELIMINACIÓN (DELETE)")
        
        if len(facturas) > 1:
            factura_id = facturas[1]['id']
            logger.info(f"Eliminando factura ID: {factura_id}")
            
            # Eliminar la factura
            eliminado = db.eliminar_factura(factura_id)
            
            if eliminado:
                logger.info("✓ Factura eliminada correctamente")
                
                # Verificar la eliminación
                factura_eliminada = next((f for f in db.obtener_facturas() 
                                        if f['id'] == factura_id), None)
                if not factura_eliminada:
                    logger.info("  La factura ya no existe en la base de datos")
                else:
                    logger.error("✗ La factura no se eliminó correctamente")
            else:
                logger.error("✗ No se pudo eliminar la factura")
        
        # 6. Verificación de integridad de datos
        print_header("6. VERIFICACIÓN DE INTEGRIDAD DE DATOS")
        
        # Verificar que los tipos de gasto no se duplicaron
        tipos = db.obtener_tipos_gasto()
        nombres_tipos = [t['nombre'] for t in tipos]
        if len(nombres_tipos) == len(set(nombres_tipos)):
            logger.info("✓ No hay tipos de gasto duplicados")
        else:
            logger.error("✗ Se encontraron tipos de gasto duplicados")
        
        # Verificar que las claves foráneas son válidas
        facturas = db.obtener_facturas()
        tipos_validos = [t['nombre'] for t in tipos]
        facturas_invalidas = [f for f in facturas if f['tipo'] not in tipos_validos]
        
        if not facturas_invalidas:
            logger.info("✓ Todas las facturas tienen tipos de gasto válidos")
        else:
            logger.error(f"✗ Se encontraron {len(facturas_invalidas)} facturas con tipos de gasto inválidos")
        
        logger.info("\n¡Todas las pruebas se completaron exitosamente!")
        return True
        
    except Exception as e:
        logger.error(f"✗ Error durante las pruebas: {str(e)}")
        logger.error(traceback.format_exc())
        return False
    finally:
        if db:
            # Limpiar la base de datos al finalizar
            if hasattr(db, '_conn') and db._conn:
                db._conn.close()
                logger.info("Conexión a la base de datos cerrada")

if __name__ == "__main__":
    if test_crud_operations():
        logger.info("¡Todas las pruebas CRUD se completaron exitosamente!")
        sys.exit(0)
    else:
        logger.error("Algunas pruebas CRUD fallaron.")
        sys.exit(1)
