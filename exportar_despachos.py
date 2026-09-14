"""
Script para exportar despachos desde DispatchTrack (LastMile) y guardarlos
dentro del repositorio (carpeta 'reportes/').

Disenado para correr automaticamente en GitHub Actions.

IMPORTANTE: la llave API NO se escribe en este archivo.
Se lee desde un secreto de GitHub llamado DISPATCHTRACK_API_KEY.
En GitHub: Settings -> Secrets and variables -> Actions
-> New repository secret -> Name: DISPATCHTRACK_API_KEY
-> Secret: (pega aqui tu llave real de DispatchTrack)
"""

import requests
import time
import os
from datetime import datetime

# ============================================================
# La llave API se toma del secreto de GitHub. NO se escribe aqui.
# ============================================================
API_KEY = os.environ["DISPATCHTRACK_API_KEY"]

# Carpeta destino dentro del repositorio
CARPETA_DESTINO = "reportes"
DESTINO = os.path.join(CARPETA_DESTINO, "despachos_diarios.xls")

# ============================================================
# Normalmente no hace falta tocar nada de aqui para abajo
# ============================================================

BASE_URL = "https://uyusa.dispatchtrack.com/api/external/v1/data_exports"

# La exportacion tarda alrededor de 5 minutos en quedar lista.
# Se revisa cada 5 minutos para no agotar el limite diario de consultas.
MAX_ESPERA_MINUTOS = 20
INTERVALO_REVISION_SEGUNDOS = 300  # 5 minutos


def log(mensaje):
    linea = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {mensaje}"
    print(linea, flush=True)


def main():
    headers = {
        "Content-Type": "application/json",
        "X-AUTH-TOKEN": API_KEY,
    }

    log("Iniciando exportacion de despachos...")

    # Paso 1: pedirle a DispatchTrack que genere una nueva exportacion
    try:
        resp = requests.post(BASE_URL, headers=headers, json={"export_items": False}, timeout=30)
        resp.raise_for_status()
        export_id = resp.json()["response"]["id"]
        log(f"Exportacion creada con id: {export_id}")
    except Exception as e:
        log(f"ERROR al crear la exportacion: {e}")
        raise

    # Paso 2: revisar cada 5 minutos hasta que este lista
    intentos_max = int((MAX_ESPERA_MINUTOS * 60) / INTERVALO_REVISION_SEGUNDOS)
    file_url = None

    for intento in range(intentos_max):
        time.sleep(INTERVALO_REVISION_SEGUNDOS)
        try:
            check = requests.get(f"{BASE_URL}/{export_id}", headers=headers, timeout=30)
            check.raise_for_status()
            info = check.json()["response"]
            estado = info.get("status")
            log(f"Revision {intento + 1}: status = {estado}")
            if estado == "success":
                file_url = info.get("file_url")
                break
        except Exception as e:
            log(f"ERROR al revisar estado: {e}")

    if not file_url:
        log(f"La exportacion no quedo lista tras {MAX_ESPERA_MINUTOS} minutos. Se detiene el script.")
        raise SystemExit(1)

    # Paso 3: descargar el archivo
    try:
        archivo = requests.get(file_url, timeout=60)
        archivo.raise_for_status()
    except Exception as e:
        log(f"ERROR al descargar el archivo: {e}")
        raise

    # Paso 4: guardarlo en la carpeta destino del repositorio
    os.makedirs(CARPETA_DESTINO, exist_ok=True)
    try:
        with open(DESTINO, "wb") as f:
            f.write(archivo.content)
        log(f"Archivo guardado correctamente en: {DESTINO}")
    except Exception as e:
        log(f"ERROR al guardar el archivo: {e}")
        raise


if __name__ == "__main__":
    main()
