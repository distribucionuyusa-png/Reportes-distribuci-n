"""
Script para exportar despachos desde DispatchTrack (LastMile) y guardarlos
dentro del repositorio (carpeta 'reportes/').

Disenado para correr automaticamente en GitHub Actions.

La llave API NO se escribe aqui: se lee del secreto de GitHub
DISPATCHTRACK_API_KEY (Settings -> Secrets and variables -> Actions).
"""

import requests
import time
import os
from datetime import datetime

# Llave tomada del secreto de GitHub. NO se escribe en este archivo.
API_KEY = os.environ["DISPATCHTRACK_API_KEY"]

# Carpeta destino dentro del repositorio
CARPETA_DESTINO = "reportes"
DESTINO = os.path.join(CARPETA_DESTINO, "despachos_diarios.xls")

BASE_URL = "https://uyusa.dispatchtrack.com/api/external/v1/data_exports"

# La exportacion tarda ~5 minutos en quedar lista.
# Por eso se espera 4 minutos antes de la primera revision (revisar antes
# solo gastaria consultas de la API sin necesidad) y luego se revisa cada
# minuto hasta que este lista. Asi la corrida termina apenas el archivo esta
# listo, sin dormir de mas.
ESPERA_INICIAL_SEGUNDOS = 240      # 4 minutos
INTERVALO_REVISION_SEGUNDOS = 60   # luego, cada minuto
MAX_ESPERA_MINUTOS = 15            # tope de seguridad


def log(mensaje):
    linea = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {mensaje}"
    print(linea, flush=True)


def main():
    headers = {
        "Content-Type": "application/json",
        "X-AUTH-TOKEN": API_KEY,
    }

    log("Iniciando exportacion de despachos...")

    # Paso 1: pedir a DispatchTrack que genere la exportacion
    try:
        resp = requests.post(BASE_URL, headers=headers, json={"export_items": False}, timeout=30)
        resp.raise_for_status()
        export_id = resp.json()["response"]["id"]
        log(f"Exportacion creada con id: {export_id}")
    except Exception as e:
        log(f"ERROR al crear la exportacion: {e}")
        raise

    # Paso 2: esperar 4 min y luego revisar cada minuto hasta que este lista
    log(f"Esperando {ESPERA_INICIAL_SEGUNDOS // 60} minutos antes de la primera revision...")
    time.sleep(ESPERA_INICIAL_SEGUNDOS)

    transcurrido = ESPERA_INICIAL_SEGUNDOS
    file_url = None

    while transcurrido <= MAX_ESPERA_MINUTOS * 60:
        try:
            check = requests.get(f"{BASE_URL}/{export_id}", headers=headers, timeout=30)
            check.raise_for_status()
            info = check.json()["response"]
            estado = info.get("status")
            log(f"Revision a los {transcurrido // 60} min: status = {estado}")
            if estado == "success":
                file_url = info.get("file_url")
                break
        except Exception as e:
            log(f"ERROR al revisar estado: {e}")

        time.sleep(INTERVALO_REVISION_SEGUNDOS)
        transcurrido += INTERVALO_REVISION_SEGUNDOS

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

    # Paso 4: guardar en la carpeta destino del repositorio
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
