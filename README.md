# Facturas2

Sistema de gestión de facturas desarrollado en Python.

## Descripción

Aplicación para la gestión de facturas que permite:
- Crear, editar y eliminar facturas
- Gestionar clientes
- Generar informes
- Otras funcionalidades de gestión de facturación

## Requisitos

- Python 3.x
- Dependencias listadas en `requirements.txt` (si aplica)

## Instalación

1. Clonar el repositorio:
   ```bash
   git clone [URL_DEL_REPOSITORIO]
   cd Facturas2
   ```

2. Crear un entorno virtual (recomendado):
   ```bash
   python -m venv venv
   .\venv\Scripts\activate  # En Windows
   source venv/bin/activate  # En macOS/Linux
   ```

3. Instalar dependencias (si existen):
   ```bash
   pip install -r requirements.txt
   ```

## Uso

1. Ejecutar la aplicación:
   ```bash
   python facturas2.py
   ```

2. Siga las instrucciones en pantalla para utilizar la aplicación.

## Creación del Ejecutable

Para crear un ejecutable de la aplicación que pueda ser distribuido y usado sin necesidad de tener Python instalado:

### Requisitos previos

1. Instalar PyInstaller:
   ```bash
   pip install pyinstaller
   ```

### Generar el ejecutable

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta el siguiente comando:
   ```powershell
   py -m PyInstaller --onefile --windowed --name="GestorFacturas" --add-data="database.py;." --add-data="config.ini;." --distpath="dist_standalone" --workpath="build_temp" facturas2.py
   ```

### Ubicación del ejecutable

El archivo ejecutable se creará en:
```
Facturas2/dist_standalone/GestorFacturas.exe
```

### Actualizar el ejecutable después de cambios

Si realizas modificaciones al código:

1. Guarda todos los cambios
2. (Opcional) Para evitar problemas, elimina las carpetas de compilación anteriores:
   ```powershell
   Remove-Item -Recurse -Force build, build_temp, dist_standalone -ErrorAction SilentlyContinue
   ```
3. Vuelve a ejecutar el comando de generación del ejecutable

### Notas importantes

- La primera vez que se ejecute, creará automáticamente una carpeta `FacturasApp` en tu directorio de usuario (`%USERPROFILE%\FacturasApp`) para almacenar:
  - Base de datos SQLite (`facturas.db`)
  - Archivo de configuración (`config.ini`)
  - Archivos de log (`facturas_qt.log`)
  - Cualquier archivo JSON de respaldo de migración

- El ejecutable es completamente independiente y no requiere instalación de Python ni dependencias adicionales.

- Para distribuir la aplicación, solo necesitas compartir el archivo `GestorFacturas.exe` del directorio `dist_standalone`.
- No es necesario tener Python instalado en la computadora donde se ejecute el programa.

## Estructura del Proyecto

```
Facturas2/
├── facturas2.py      # Archivo principal de la aplicación
├── README.md         # Este archivo
└── .gitignore        # Archivo para ignorar archivos en git
```

## Contribución

Las contribuciones son bienvenidas. Por favor, abra un issue primero para discutir los cambios que le gustaría realizar.

## Licencia

Este proyecto está bajo la [Licencia MIT](LICENSE).
