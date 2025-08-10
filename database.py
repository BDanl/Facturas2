import sqlite3
import json
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional

# Configurar logging
logger = logging.getLogger(__name__)

class Database:
    def __init__(self, db_path: str = 'facturas.db'):
        """Inicializa la conexión a la base de datos SQLite."""
        self.db_path = db_path
        self._conn = None
        self._is_memory_db = db_path == ':memory:'
        
        # Para bases de datos en memoria, creamos una conexión persistente
        if self._is_memory_db:
            self._conn = sqlite3.connect(':memory:')
            self._conn.row_factory = sqlite3.Row
            logger.info("Conexión a base de datos en memoria creada")
        
        self._create_tables()
    
    def _get_connection(self):
        """Obtiene una conexión a la base de datos."""
        if self._is_memory_db and self._conn is not None:
            return self._conn
        
        # Para bases de datos de archivo, creamos una nueva conexión cada vez
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def __del__(self):
        """Cierra la conexión a la base de datos al destruir la instancia."""
        if self._is_memory_db and self._conn is not None:
            self._conn.close()
            logger.info("Conexión a base de datos en memoria cerrada")
    
    def _create_tables(self):
        """Crea las tablas necesarias si no existen."""
        logger.info("Iniciando creación de tablas...")
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                logger.info("Conexión a la base de datos establecida")
                
                # Tabla de tipos de gastos
                logger.info("Creando tabla 'tipos_gasto' si no existe...")
                cursor.execute('''
                CREATE TABLE IF NOT EXISTS tipos_gasto (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE NOT NULL,
                    descripcion TEXT,
                    color TEXT
                )''')
                logger.info("Tabla 'tipos_gasto' creada o verificada")
                
                # Tabla de facturas
                logger.info("Creando tabla 'facturas' si no existe...")
                cursor.execute('''
                CREATE TABLE IF NOT EXISTS facturas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fecha DATE NOT NULL,
                    tipo_id INTEGER NOT NULL,
                    descripcion TEXT NOT NULL,
                    valor REAL NOT NULL,
                    fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (tipo_id) REFERENCES tipos_gasto (id)
                )''')
                logger.info("Tabla 'facturas' creada o verificada")
                
                # Índices para mejorar el rendimiento de las consultas
                logger.info("Creando índices...")
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_facturas_fecha ON facturas(fecha)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_facturas_tipo ON facturas(tipo_id)')
                logger.info("Índices creados o verificados")
                
                # Insertar tipos de gastos por defecto si no existen
                logger.info("Insertando tipos de gasto por defecto...")
                self._insert_default_tipos_gasto(cursor)
                
                # Verificar que las tablas se crearon correctamente
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tablas = cursor.fetchall()
                logger.info(f"Tablas en la base de datos: {[t[0] for t in tablas]}")
                
                conn.commit()
                logger.info("Cambios confirmados en la base de datos")
                
        except Exception as e:
            logger.error(f"Error al crear las tablas: {str(e)}")
            logger.error(traceback.format_exc())
            raise
    
    def _insert_default_tipos_gasto(self, cursor):
        """Inserta los tipos de gastos por defecto."""
        default_tipos = [
            ('Mercado', 'Compras de supermercado', '#FF6B6B'),
            ('Transporte', 'Transporte público/privado', '#4ECDC4'),
            ('Entretenimiento', 'Cine, salidas, etc.', '#45B7D1'),
            ('Servicios', 'Luz, agua, internet, etc.', '#96CEB4'),
            ('Salud', 'Gastos médicos', '#FFEEAD'),
            ('Educación', 'Cursos, libros, etc.', '#D4A373'),
            ('Gastos Básicos', 'Gastos básicos del hogar', '#9B5DE5'),
            ('Ocio', 'Actividades de ocio y diversión', '#FF9F1C'),
            ('Reparaciones', 'Reparaciones y mantenimiento', '#2EC4B6'),
            ('Préstamo', 'Pagos de préstamos', '#E71D36'),
            ('Ahorro', 'Ahorros e inversiones', '#2EC4B6'),
            ('Predial', 'Impuesto predial', '#6A4C93'),
            ('Gastos fijos', 'Gastos fijos mensuales', '#9B5DE5'),
            ('Otros', 'Otros gastos', '#8B8C89')
        ]
        
        for tipo in default_tipos:
            cursor.execute(
                'INSERT OR IGNORE INTO tipos_gasto (nombre, descripcion, color) VALUES (?, ?, ?)',
                tipo
            )
    
    def migrar_desde_json(self, json_path: str) -> int:
        """
        Migra los datos desde un archivo JSON a la base de datos SQLite.
        
        Args:
            json_path: Ruta al archivo JSON con los datos a migrar.
            
        Returns:
            int: Número de facturas migradas.
        """
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                facturas = json.load(f)
            
            with self._get_connection() as conn:
                cursor = conn.cursor()
                count = 0
                
                for factura in facturas:
                    # Obtener o crear el tipo de gasto
                    cursor.execute(
                        'SELECT id FROM tipos_gasto WHERE nombre = ?',
                        (factura['tipo'],)
                    )
                    tipo = cursor.fetchone()
                    
                    if not tipo:
                        cursor.execute(
                            'INSERT INTO tipos_gasto (nombre) VALUES (?)',
                            (factura['tipo'],)
                        )
                        tipo_id = cursor.lastrowid
                    else:
                        tipo_id = tipo['id']
                    
                    # Convertir la fecha al formato YYYY-MM-DD
                    try:
                        fecha = datetime.strptime(factura['fecha'], '%d/%m/%Y').strftime('%Y-%m-%d')
                    except (ValueError, KeyError):
                        # Si hay un error con la fecha, usar la fecha actual
                        fecha = datetime.now().strftime('%Y-%m-%d')
                    
                    # Insertar la factura
                    cursor.execute('''
                        INSERT INTO facturas (fecha, tipo_id, descripcion, valor)
                        VALUES (?, ?, ?, ?)
                    ''', (
                        fecha,
                        tipo_id,
                        factura.get('descripcion', ''),
                        float(factura['valor'])
                    ))
                    
                    count += 1
                
                conn.commit()
                return count
                
        except Exception as e:
            logger.error(f"Error al migrar datos desde JSON: {str(e)}")
            raise
    
    def obtener_facturas(self, fecha_inicio: str = None, fecha_fin: str = None) -> List[Dict[str, Any]]:
        """
        Obtiene las facturas de la base de datos.
        
        Args:
            fecha_inicio: Fecha de inicio en formato YYYY-MM-DD (opcional).
            fecha_fin: Fecha de fin en formato YYYY-MM-DD (opcional).
            
        Returns:
            List[Dict]: Lista de diccionarios con los datos de las facturas.
        """
        query = '''
            SELECT 
                f.id,
                f.fecha,
                tg.nombre as tipo,
                f.descripcion,
                f.valor,
                tg.color
            FROM facturas f
            JOIN tipos_gasto tg ON f.tipo_id = tg.id
        '''
        
        params = []
        
        if fecha_inicio and fecha_fin:
            query += ' WHERE f.fecha BETWEEN ? AND ?'
            params.extend([fecha_inicio, fecha_fin])
        
        query += ' ORDER BY f.fecha DESC'
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            
            # Convertir los resultados a una lista de diccionarios
            facturas = []
            for row in cursor.fetchall():
                factura = dict(row)
                # Convertir la fecha al formato DD/MM/YYYY
                try:
                    fecha = datetime.strptime(factura['fecha'], '%Y-%m-%d').strftime('%d/%m/%Y')
                    factura['fecha'] = fecha
                except (ValueError, KeyError):
                    pass
                
                facturas.append(factura)
            
            return facturas
    
    def agregar_factura(self, fecha: str, tipo: str, descripcion: str, valor: float) -> int:
        """
        Agrega una nueva factura a la base de datos.
        
        Args:
            fecha: Fecha en formato DD/MM/YYYY.
            tipo: Nombre del tipo de gasto.
            descripcion: Descripción del gasto.
            valor: Monto del gasto.
            
        Returns:
            int: ID de la factura insertada.
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # Obtener o crear el tipo de gasto
                cursor.execute(
                    'SELECT id FROM tipos_gasto WHERE nombre = ?',
                    (tipo,)
                )
                tipo_row = cursor.fetchone()
                
                if not tipo_row:
                    cursor.execute(
                        'INSERT INTO tipos_gasto (nombre) VALUES (?)',
                        (tipo,)
                    )
                    tipo_id = cursor.lastrowid
                else:
                    tipo_id = tipo_row['id']
                
                # Convertir la fecha al formato YYYY-MM-DD
                try:
                    fecha_db = datetime.strptime(fecha, '%d/%m/%Y').strftime('%Y-%m-%d')
                except ValueError:
                    fecha_db = datetime.now().strftime('%Y-%m-%d')
                
                # Insertar la factura
                cursor.execute('''
                    INSERT INTO facturas (fecha, tipo_id, descripcion, valor)
                    VALUES (?, ?, ?, ?)
                ''', (
                    fecha_db,
                    tipo_id,
                    descripcion,
                    float(valor)
                ))
                
                factura_id = cursor.lastrowid
                conn.commit()
                return factura_id
                
        except Exception as e:
            logger.error(f"Error al agregar factura: {str(e)}")
            raise
    
    def actualizar_factura(self, factura_id: int, fecha: str, tipo: str, descripcion: str, valor: float) -> bool:
        """
        Actualiza una factura existente.
        
        Args:
            factura_id: ID de la factura a actualizar.
            fecha: Nueva fecha en formato DD/MM/YYYY.
            tipo: Nuevo tipo de gasto.
            descripcion: Nueva descripción.
            valor: Nuevo valor.
            
        Returns:
            bool: True si la actualización fue exitosa, False en caso contrario.
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # Obtener o crear el tipo de gasto
                cursor.execute(
                    'SELECT id FROM tipos_gasto WHERE nombre = ?',
                    (tipo,)
                )
                tipo_row = cursor.fetchone()
                
                if not tipo_row:
                    cursor.execute(
                        'INSERT INTO tipos_gasto (nombre) VALUES (?)',
                        (tipo,)
                    )
                    tipo_id = cursor.lastrowid
                else:
                    tipo_id = tipo_row['id']
                
                # Convertir la fecha al formato YYYY-MM-DD
                try:
                    fecha_db = datetime.strptime(fecha, '%d/%m/%Y').strftime('%Y-%m-%d')
                except ValueError:
                    fecha_db = datetime.now().strftime('%Y-%m-%d')
                
                # Actualizar la factura
                cursor.execute('''
                    UPDATE facturas
                    SET fecha = ?,
                        tipo_id = ?,
                        descripcion = ?,
                        valor = ?,
                        fecha_actualizacion = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (
                    fecha_db,
                    tipo_id,
                    descripcion,
                    float(valor),
                    factura_id
                ))
                
                conn.commit()
                return cursor.rowcount > 0
                
        except Exception as e:
            logger.error(f"Error al actualizar factura: {str(e)}")
            return False
    
    def eliminar_factura(self, factura_id: int) -> bool:
        """
        Elimina una factura de la base de datos.
        
        Args:
            factura_id: ID de la factura a eliminar.
            
        Returns:
            bool: True si la eliminación fue exitosa, False en caso contrario.
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM facturas WHERE id = ?', (factura_id,))
                conn.commit()
                return cursor.rowcount > 0
                
        except Exception as e:
            logger.error(f"Error al eliminar factura: {str(e)}")
            return False
    
    def obtener_tipos_gasto(self) -> List[Dict[str, Any]]:
        """
        Obtiene la lista de tipos de gasto, excluyendo 'coche' y 'coches'.
        
        Returns:
            List[Dict]: Lista de diccionarios con los tipos de gasto.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, nombre, descripcion, color 
                FROM tipos_gasto 
                WHERE LOWER(nombre) NOT IN ('coche', 'coches')
                ORDER BY nombre
            ''')
            return [dict(row) for row in cursor.fetchall()]
    
    def obtener_resumen_por_tipo(self, fecha_inicio: str = None, fecha_fin: str = None) -> List[Dict[str, Any]]:
        """
        Obtiene un resumen de gastos por tipo.
        
        Args:
            fecha_inicio: Fecha de inicio en formato YYYY-MM-DD (opcional).
            fecha_fin: Fecha de fin en formato YYYY-MM-DD (opcional).
            
        Returns:
            List[Dict]: Lista de diccionarios con el resumen por tipo.
        """
        query = '''
            SELECT 
                tg.nombre as tipo,
                tg.color,
                COUNT(f.id) as cantidad,
                SUM(f.valor) as total
            FROM tipos_gasto tg
            LEFT JOIN facturas f ON tg.id = f.tipo_id
        '''
        
        params = []
        where_clause = []
        
        if fecha_inicio and fecha_fin:
            where_clause.append('f.fecha BETWEEN ? AND ?')
            params.extend([fecha_inicio, fecha_fin])
        
        if where_clause:
            query += ' WHERE ' + ' AND '.join(where_clause)
        
        query += ' GROUP BY tg.id, tg.nombre, tg.color ORDER BY total DESC'
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    
    def obtener_resumen_mensual(self, anio: int = None) -> List[Dict[str, Any]]:
        """
        Obtiene un resumen de gastos por mes.
        
        Args:
            anio: Año para el resumen (opcional, si no se especifica usa el año actual).
            
        Returns:
            List[Dict]: Lista de diccionarios con el resumen mensual.
        """
        if anio is None:
            anio = datetime.now().year
        
        query = '''
            SELECT 
                strftime('%Y-%m', f.fecha) as mes,
                SUM(f.valor) as total
            FROM facturas f
            WHERE strftime('%Y', f.fecha) = ?
            GROUP BY strftime('%Y-%m', f.fecha)
            ORDER BY mes
        '''
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (str(anio),))
            return [dict(row) for row in cursor.fetchall()]


def migrar_datos_desde_json(json_path: str, db_path: str = 'facturas.db') -> int:
    """
    Función de conveniencia para migrar datos desde un archivo JSON a SQLite.
    
    Args:
        json_path: Ruta al archivo JSON con los datos a migrar.
        db_path: Ruta donde se creará o actualizará la base de datos SQLite.
        
    Returns:
        int: Número de facturas migradas.
    """
    db = Database(db_path)
    return db.migrar_desde_json(json_path)


if __name__ == '__main__':
    # Ejemplo de uso
    import sys
    
    if len(sys.argv) > 1:
        json_path = sys.argv[1]
        db_path = sys.argv[2] if len(sys.argv) > 2 else 'facturas.db'
        
        try:
            count = migrar_datos_desde_json(json_path, db_path)
            print(f"Se migraron {count} facturas exitosamente a {db_path}")
        except Exception as e:
            print(f"Error al migrar datos: {str(e)}")
            sys.exit(1)
    else:
        print("Uso: python database.py <ruta_al_archivo_json> [ruta_salida_db]")
        sys.exit(1)
