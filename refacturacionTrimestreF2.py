from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import pandas as pd
import locale
from datetime import datetime
from dateutil.relativedelta import relativedelta



# 🚀 CONFIGURACIÓN FINAL
MODO_TEST = True  # si esta en False es para numeros Reales si ponemos True para prueba
NUMERO_TEST = "5493516570658"  # Tu número real (ej: 5493511234567)
LIMITAR_A = 0  # 0 = sin límite

# Establecer idioma español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try: 
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
    except:
        pass

# LEER EXCEL
archivo = "clientesActualizadoCopia.xlsx"
df_estado = pd.read_excel(archivo, sheet_name="Estado de cuenta")
df_polizas = pd.read_excel(archivo, sheet_name="Polizas", dtype={"telefono": str})
df_estado.columns = df_estado.columns.str.strip().str.lower()
df_polizas.columns = df_polizas.columns.str.strip().str.lower()

def limpiar_dni(dni):
    return ''.join(filter(str.isdigit, str(dni)))

def limpiar_telefono(numero):
    if pd.isna(numero):
        return ""

    # Si viene como float, cortamos decimales
    if isinstance(numero, float):
        numero = str(int(numero))
    else:
        numero = str(numero).strip()

    numero = ''.join(filter(str.isdigit, numero))

    if numero.startswith("351") and len(numero) == 10:
        return "549" + numero

    if numero.startswith("0351"):
        return "549" + numero[1:]

    if numero.startswith("15") and len(numero) >= 9:
        return "549351" + numero[2:]

    if numero.startswith("549") and len(numero) >= 12:
        return numero

    if len(numero) == 10:
        return "549" + numero

    return numero



def obtener_mes_espanol(fecha, compania):
    try:
        fecha = pd.to_datetime(fecha)
        if "holando" in compania.lower():
            return fecha.strftime("%B").capitalize()
        else:
            fecha_sumada = fecha + relativedelta(months=1)
            return fecha_sumada.strftime("%B").capitalize()
    except:
        return "próximos días"

# Armado de mensajes
errores = []
mensajes = []

for i in range(min(len(df_polizas), len(df_estado))):
    fila_poliza = df_polizas.iloc[i]
    fila_estado = df_estado.iloc[i]

    error = ""

    dni = limpiar_dni(fila_poliza.get("dni", ""))
    telefono = limpiar_telefono(fila_poliza.get("telefono", ""))
    riesgo = str(fila_poliza.get("riesgo", "")).strip().lower()
    nombre = str(fila_estado.get("apellido y nombre", "")).title()
    compania = str(fila_estado.get("compañía", "")).strip()
    fecha = fila_estado.get("flyer", "")
    refacturacion = str(fila_estado.get("refacturación", "")).strip()
    estado = str(fila_estado.get("estado", "")).strip()

    if refacturacion.lower() != "trimestral":
        continue
    if estado.upper() != "SI":
        continue

    if not dni:
        error += "DNI vacío; "
    if not telefono:
        error += "Teléfono vacío o no numérico; "
    if not riesgo:
        error += "Riesgo vacío; "
    if not nombre:
        error += "Nombre vacío; "
    if not compania:
        error += "Compañía vacía; "
    if pd.isna(fecha):
        error += "Fecha flyer vacía; "
    if not refacturacion:
        error += "Refacturación vacía; "

    mes = obtener_mes_espanol(fecha, compania)

    mensaje = (
        f"Hola {nombre}, este mensaje originalmente sería enviado al número *{telefono}*, "
        f"Te recordamos que en el mes de *{mes}* se debitará tu póliza de *{compania}*, "
        f"correspondiente al seguro de *{riesgo.title()}*. "
        f"La refacturación de esta póliza es *{refacturacion.lower()}*.\n "
        f"¡Gracias por confiar en nosotros!"
    )

    mensajes.append({
        "index": i + 1,
        "telefono": telefono,
        "compañia": compania,
        "mensaje": mensaje,
        "error": error.strip()
    })
    # Crear lista de pendientes
pendientes = []

for i in range(len(df_estado)):
    fila_estado = df_estado.iloc[i]
    fila_poliza = df_polizas.iloc[i] if i < len(df_polizas) else {}

    estado = str(fila_estado.get("estado", "")).strip().upper()
    refacturacion = str(fila_estado.get("refacturación", "")).strip().lower()

    if estado != "SI" and refacturacion == "trimestral":
        pendientes.append({
            "index": i + 1,
            "apellido y nombre": fila_estado.get("apellido y nombre", ""),
            "dni": fila_poliza.get("dni", ""),
            "telefono": fila_poliza.get("telefono", ""),
            "compañía": fila_estado.get("compañía", ""),
            "riesgo": fila_poliza.get("riesgo", ""),
            "estado": estado,
            "refacturación": refacturacion,
            "motivo": "Pendientes"
        })

# Guardar Excel de pendientes
if pendientes:
    df_pendientes = pd.DataFrame(pendientes)
    df_pendientes.to_excel("refacturacionTrimestral/pendientes.xlsx", index=False)
    print("📄 Archivo generado: estado_pendientes.xlsx")


# GUARDAR ARCHIVO DE VERIFICACIÓN
df_verificacion = pd.DataFrame(mensajes)
df_verificacion.to_excel("refacturacionTrimestral/enviados.xlsx", index=False)
print("📄 Archivo generado: mensajes_autorizados.xlsx")




if mensajes:
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com/")
    input("📲 Escaneá el código QR y presioná ENTER para continuar...")

    for m in mensajes:
        if LIMITAR_A and m["index"] > LIMITAR_A:
            break

        telefono_real = m["telefono"]
        mensaje = m["mensaje"]
        destino = NUMERO_TEST if MODO_TEST else telefono_real
        print(f"📞 Enviando a: {destino}")  # DEBUG

        try:
            # Buscar contacto en barra de búsqueda
            search_box = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )

            search_box.clear()
            search_box.click()
            search_box.send_keys(destino)
            search_box.send_keys(Keys.ENTER)

            # Esperar que cargue el chat
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
            )
            time.sleep(1)

            # Escribir y enviar el mensaje
            input_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            input_box.click()
            input_box.send_keys(mensaje)
            input_box.send_keys(Keys.ENTER)

            print(f"✅ Mensaje #{m['index']} {'(modo TEST)' if MODO_TEST else ''} enviado a {destino}")
            time.sleep(3)

        except Exception as e:
            print(f"❌ Error al enviar mensaje #{m['index']} a {destino}: {e}")