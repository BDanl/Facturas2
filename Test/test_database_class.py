import sys
import os
import logging
import traceback
from database import Database

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)  # Forzar salida a stdout
    ]
)
logger = logging.getLogger(__name__)

def print_header(title):
    """Imprime un encabezado para las pruebas."""
    logger.info("\n" + "="*80)
    logger.info(f" {title} ".center(80, '='))
    logger.info("="*80 + "\n")

def test_database_class():
    """Prueba la clase Database."""
    try:
        # 1. Inicialización
        print_header("1. PRUEBA DE INICIALIZACIÓN")
        db_path = ":memory:"
        logger.info(f"Creando base de datos en: {db_path}")
        db = Database(db_path)
        logger.info("✓ Base de datos inicializada correctamente")
        
        # 2. Prueba de obtener_tipos_gasto
        print_header("2. PRUEBA DE OBTENER TIPOS DE GASTO")
        try:
            tipos = db.obtener_tipos_gasto()
            logger.info(f"✓ Se obtuvieron {len(tipos)} tipos de gasto")
            for i, tipo in enumerate(tipos, 1):
                logger.info(f"  {i}. {tipo['nombre']} - {tipo['descripcion']}")
        except Exception as e:
            logger.error(f"✗ Error en obtener_tipos_gasto: {str(e)}")
            logger.error(traceback.format_exc())
            return False
        
        # 3. Prueba de agregar_factura
        print_header("3. PRUEBA DE AGREGAR FACTURA")
        try:
            factura_id = db.agregar_factura(
                fecha="07/08/2025",
                tipo="Mercado",
                descripcion="Compra de supermercado",
                valor=150000.0
            )
            logger.info(f"✓ Factura agregada con ID: {factura_id}")
            
            # Verificar que la factura se agregó
            facturas = db.obtener_facturas()
            logger.info(f"Total de facturas en la base de datos: {len(facturas)}")
            if facturas:
                logger.info(f"Última factura: {facturas[0]}")
            
        except Exception as e:
            logger.error(f"✗ Error en agregar_factura: {str(e)}")
            logger.error(traceback.format_exc())
            return False
        
        logger.info("\n¡Todas las pruebas se completaron exitosamente!")
        return True
        
    except Exception as e:
        logger.error(f"✗ Error inesperado: {str(e)}")
        logger.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    # Configurar el nivel de logging para la aplicación
    logging.getLogger().setLevel(logging.DEBUG)
    
    # Ejecutar pruebas
    if test_database_class():
        logger.info("¡Todas las pruebas se completaron exitosamente!")
        sys.exit(0)
    else:
        logger.error("Algunas pruebas fallaron.")
        sys.exit(1)
