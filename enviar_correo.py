# -*- coding: utf-8 -*-
"""
enviar_correo.py
----------------
Envia el maestro acumulado (historico/historico_despachos.xlsx) como adjunto
por correo, usando SMTP de Gmail. Power Automate recibe ese correo en el buzon
de Microsoft 365 y guarda el adjunto en OneDrive.

Lee tres variables de entorno (se configuran como secretos en GitHub):
  MAIL_USERNAME  -> la cuenta Gmail que envia (ej. distribucion.uyusa@gmail.com)
  MAIL_PASSWORD  -> contrasena de aplicacion de Gmail (no la contrasena normal)
  MAIL_TO        -> el correo de Microsoft 365 que recibe (buzon que vigila el flujo)

El asunto es fijo ("Historico despachos acumulado") para que el flujo lo filtre.
"""

import os
import ssl
import smtplib
from email.message import EmailMessage

ASUNTO = "Historico despachos acumulado"
ARCHIVO = "historico/historico_despachos.xlsx"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    user = os.environ["MAIL_USERNAME"]
    pwd = os.environ["MAIL_PASSWORD"]
    to = os.environ["MAIL_TO"]

    if not os.path.exists(ARCHIVO):
        raise SystemExit(f"No se encontro el maestro: {ARCHIVO}")

    msg = EmailMessage()
    msg["Subject"] = ASUNTO
    msg["From"] = user
    msg["To"] = to
    msg.set_content(
        "Adjunto el maestro acumulado de despachos (generado automaticamente)."
    )
    with open(ARCHIVO, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype=XLSX.split("/", 1)[1],
            filename="historico_despachos.xlsx",
        )

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
        s.login(user, pwd)
        s.send_message(msg)
    print(f"Correo enviado a {to} (adjunto {ARCHIVO})")


if __name__ == "__main__":
    main()
