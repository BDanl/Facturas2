import subprocess
import time
import sys
import os

def run():
    print("Iniciando la aplicación...")
    process = subprocess.Popen([sys.executable, "facturas2.py"])
    last_mtime = 0
    
    try:
        while True:
            current_mtime = os.path.getmtime("facturas2.py")
            if current_mtime > last_mtime:
                print("\n--- Archivo modificado, reiniciando... ---\n")
                process.terminate()
                process.wait()
                process = subprocess.Popen([sys.executable, "facturas2.py"])
                last_mtime = current_mtime
            time.sleep(1)
    except KeyboardInterrupt:
        process.terminate()
        process.wait()

if __name__ == "__main__":
    run()