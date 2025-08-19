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
import random

# 🚀 CONFIGURACIÓN FINAL
MODO_TEST = False
NUMERO_TEST = "5493516570658"
LIMITAR_A = 0

# Establecer idioma español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
    except:
        pass

# INPUT DE FILTRO
print("\nSeleccioná una opción de envío:")
print("1 - Solo Holando con forma de pago CBU")
print("2 - Solo Holando con forma de pago Tarjeta")
print("3 - Holando con Cupón + todas las demás compañías (todas las formas de pago)")
opcion = input("Ingresá 1, 2 o 3: ").strip()

# LEER EXCEL
archivo = "clientes.xlsx"
df_estado = pd.read_excel(archivo, sheet_name="Estado de cuenta")
df_polizas = pd.read_excel(archivo, sheet_name="Polizas", dtype={"telefono": str})
df_estado.columns = df_estado.columns.str.strip().str.lower()
df_polizas.columns = df_polizas.columns.str.strip().str.lower()

def limpiar_dni(dni):
    return ''.join(filter(str.isdigit, str(dni)))

def limpiar_telefono(numero):
    if pd.isna(numero):
        return ""
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
        return fecha.strftime("%B").capitalize()
    except:
        return "próximos días"

# Carpeta con fecha actual
fecha_actual = datetime.now().strftime("%Y-%m-%d")
os.makedirs(f"refacturacionTrimestral/{fecha_actual}", exist_ok=True)

errores = []
mensajes = []
pendientes = []

for i in range(len(df_estado)):
    fila_estado = df_estado.iloc[i]
    fila_poliza = df_polizas.iloc[i] if i < len(df_polizas) else {}

    dni = limpiar_dni(fila_poliza.get("dni", ""))
    telefono = limpiar_telefono(fila_poliza.get("telefono", ""))
    riesgo = str(fila_poliza.get("riesgo", "")).strip().lower()
    nombre = str(fila_estado.get("apellido y nombre", "")).title()
    compania = str(fila_estado.get("compañía", "")).strip()
    fecha = fila_estado.get("flyer", "")
    refacturacion = str(fila_estado.get("refacturación", "")).strip()
    estado = str(fila_estado.get("estados", "")).strip().upper()
    forma_pago = str(fila_estado.get("forma de pago", "")).strip().lower()
    compania_lower = compania.lower()

    # --- FILTROS ---
    if refacturacion.lower() != "trimestral" or estado != "SI":
        continue

    if opcion == "1" and not (compania_lower == "holando" and forma_pago == "cbu"):
        continue
    if opcion == "2" and not (compania_lower == "holando" and forma_pago == "tarjeta"):
        continue
    if opcion == "3":
        if "galeno" in compania_lower:  # Excluir Galeno siempre
            continue
        if not ((compania_lower == "holando" and forma_pago == "cupon") or (compania_lower != "holando")):
            continue

    # --- ERRORES DE DATOS ---
    error = ""
    if not dni: error += "DNI vacío; "
    if not telefono: error += "Teléfono vacío o no numérico; "
    if not riesgo: error += "Riesgo vacío; "
    if not nombre: error += "Nombre vacío; "
    if not compania: error += "Compañía vacía; "
    if pd.isna(fecha): error += "Fecha flyer vacía; "

    # --- MENSAJE ---
    mes = obtener_mes_espanol(fecha, compania)
    if forma_pago == "cupon":
        try:
            fecha_vencimiento = pd.to_datetime(fecha).strftime("%d/%m/%Y")
        except:
            fecha_vencimiento = "próximos días"
        mensaje = (
            f"Hola {nombre}, "
            f"Te recordamos que el día *{fecha_vencimiento}* se vencerá tu póliza de *{compania}*, "
            f"correspondiente al seguro de *{riesgo.title()}*."
            f"❗Tu póliza está al día!\n"
            f"✅¡Gracias por confiar en nosotros!"
        )
    else:
        mensaje = (
            f"Hola {nombre}, "
            f"Te recordamos que en el mes de *{mes}* se debitará tu póliza de *{compania}*, "
            f"correspondiente al seguro de *{riesgo.title()}*."
            f"❗Tu póliza está al día!\n"
            f"✅¡Gracias por confiar en nosotros!"
        )

    # --- AGREGAR A LISTA ---
    mensajes.append({
        "index": i + 1,
        "telefono": telefono,
        "compañia": compania,
        "riesgo": riesgo,
        "nombre": nombre,
        "dni": dni,
        "forma_pago": forma_pago,
        "refacturacion": refacturacion,
        "estado": estado,
        "mensaje": mensaje,
        "error": error.strip()
    })

    # --- PENDIENTES ---
    if estado != "SI" and refacturacion == "trimestral":
        pendientes.append({
            "index": i + 1,
            "apellido y nombre": nombre,
            "dni": dni,
            "telefono": telefono,
            "compañía": compania,
            "riesgo": riesgo,
            "estado": estado,
            "refacturación": refacturacion,
            "motivo": "Pendientes"
        })

# Guardar Excel
if pendientes:
    df_pendientes = pd.DataFrame(pendientes)
    df_pendientes.to_excel(f"refacturacionTrimestral/{fecha_actual}/pendientes.xlsx", index=False)

df_verificacion = pd.DataFrame(mensajes)
df_verificacion.to_excel(f"refacturacionTrimestral/{fecha_actual}/enviados.xlsx", index=False)

log_envios = []
if mensajes:
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com/")
    input("📲 Escaneá el código QR y presioná ENTER para continuar...")

try:
    for m in mensajes:
        if LIMITAR_A and m["index"] > LIMITAR_A:
            break
        destino = NUMERO_TEST if MODO_TEST else m["telefono"]
        print(f"📲 Enviando a: {destino}")

        try:
            search_box = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
            search_box.clear()
            search_box.click()
            search_box.send_keys(destino)
            search_box.send_keys(Keys.ENTER)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
            )
            time.sleep(1)

            input_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            input_box.click()
            input_box.send_keys(m["mensaje"])
            input_box.send_keys(Keys.ENTER)

            print(f"✅ Mensaje #{m['index']} enviado a {destino}")
            if m["index"] % random.randint(9, 14) == 0: ## Este fue el ultimo cambio que hice
                # Pausa más larga cada 10 mensajes
                tiempo_espera = random.uniform(60, 120)  # pausa más larga cada 10 mensajes // final
            else:
                tiempo_espera = random.uniform(10, 25)
            time.sleep(tiempo_espera)

            log_envios.append({
                **m,
                "fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "estado_envio": "OK",
                "error_envio": ""
            })

        except Exception as e:
            print(f"❌ Error al enviar mensaje #{m['index']} a {destino}: {e}")
            log_envios.append({
                **m,
                "fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "estado_envio": "ERROR",
                "error_envio": str(e)
            })

finally:
    pd.DataFrame(log_envios).to_excel(f"refacturacionTrimestral/{fecha_actual}/log_enviosTrimestrales.xlsx", index=False)
    print("📄 Log de envíos guardado.")
