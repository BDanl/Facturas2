import sys
import os
import traceback
from pathlib import Path
from database import Database
import logging

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,  # Cambiado a DEBUG para más detalles
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def test_database_operations():
    """Prueba las operaciones CRUD en la base de datos."""
    try:
        # Crear una base de datos en memoria para pruebas
        db_path = ":memory:"
        logger.info(f"Creando base de datos en memoria: {db_path}")
        
        # Imprimir el directorio actual para depuración
        logger.info(f"Directorio de trabajo actual: {os.getcwd()}")
        
        # Crear la base de datos
        db = Database(db_path)
        logger.info("Base de datos creada exitosamente")
        
        # Verificar si las tablas existen
        logger.info("Verificando tablas en la base de datos...")
        try:
            with db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tablas = cursor.fetchall()
                
                # Convertir los resultados a una lista de nombres de tablas
                nombres_tablas = [t[0] for t in tablas]  # Usar índice numérico en lugar de nombre de columna
                logger.info(f"Tablas en la base de datos: {nombres_tablas}")
                
                # Verificar la estructura de la tabla tipos_gasto
                if 'tipos_gasto' in nombres_tablas:
                    logger.info("La tabla 'tipos_gasto' existe en la base de datos.")
                    cursor.execute("PRAGMA table_info(tipos_gasto)")
                    columnas = cursor.fetchall()
                    logger.info("Estructura de la tabla tipos_gasto:")
                    for col in columnas:
                        # Usar índices numéricos para mayor compatibilidad
                        logger.info(f"  {col[1]} ({col[2]})")  # col[1] = nombre, col[2] = tipo
                else:
                    logger.warning("La tabla 'tipos_gasto' NO existe en la base de datos.")
                    
        except Exception as e:
            logger.error(f"Error al verificar la estructura de la base de datos: {str(e)}")
            logger.error(traceback.format_exc())
            raise
        
        logger.info("=== Iniciando pruebas de la base de datos ===")
        
        # 1. Probar obtener tipos de gasto
        logger.info("\n1. Probando obtener_tipos_gasto()...")
        try:
            tipos = db.obtener_tipos_gasto()
            logger.info(f"✓ Tipos de gasto cargados exitosamente: {len(tipos)} tipos")
            for i, tipo in enumerate(tipos, 1):
                logger.debug(f"  Tipo {i}: {tipo}")
        except Exception as e:
            logger.error(f"✗ Error en obtener_tipos_gasto(): {str(e)}")
            logger.error(traceback.format_exc())
            raise
        
        # 2. Probar agregar una factura
        factura_id = db.agregar_factura(
            fecha="07/08/2025",
            tipo="Mercado",
            descripcion="Compra de supermercado",
            valor=150000.0
        )
        logger.info(f"Factura agregada con ID: {factura_id}")
        
        # 3. Obtener la factura recién agregada
        facturas = db.obtener_facturas()
        logger.info(f"Total de facturas: {len(facturas)}")
        if facturas:
            logger.info(f"Última factura: {facturas[0]}")
        
        # 4. Probar actualizar la factura
        if facturas:
            actualizado = db.actualizar_factura(
                factura_id=facturas[0]['id'],
                fecha="08/08/2025",
                tipo="Mercado",
                descripcion="Compra de supermercado (actualizada)",
                valor=160000.0
            )
            logger.info(f"Factura actualizada: {actualizado}")
        
        # 5. Probar obtener resumen por tipo
        resumen_tipos = db.obtener_resumen_por_tipo()
        logger.info("Resumen por tipo:")
        for item in resumen_tipos:
            total = item.get('total', 0) or 0  # Handle None values
            logger.info(f"- {item.get('tipo', 'Desconocido')}: ${total:,.2f}")
        
        # 6. Probar obtener resumen mensual
        resumen_mensual = db.obtener_resumen_mensual()
        logger.info("Resumen mensual:")
        for item in resumen_mensual:
            total = item.get('total', 0) or 0  # Handle None values
            mes = item.get('mes', '??')
            anio = item.get('anio', '????')
            logger.info(f"- {mes}/{anio}: ${total:,.2f}")
        
        # 7. Probar eliminar la factura
        if facturas:
            eliminado = db.eliminar_factura(facturas[0]['id'])
            logger.info(f"Factura eliminada: {eliminado}")
            
            # Verificar que se eliminó
            facturas_despues = db.obtener_facturas()
            logger.info(f"Total de facturas después de eliminar: {len(facturas_despues)}")
        
        logger.info("=== Pruebas completadas exitosamente ===")
        
    except Exception as e:
        logger.error(f"Error durante las pruebas: {str(e)}", exc_info=True)
        return False
    
    return True

if __name__ == "__main__":
    if test_database_operations():
        logger.info("Todas las pruebas se completaron exitosamente.")
    else:
        logger.error("Algunas pruebas fallaron.")
        sys.exit(1)
