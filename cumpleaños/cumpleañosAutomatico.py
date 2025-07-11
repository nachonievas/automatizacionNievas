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

# 🚀 CONFIGURACIÓN
MODO_TEST = False  # True para pruebas, False para envío real
NUMERO_TEST = "5493516570658"
LIMITAR_A = 0  # 0 = sin límite
EXCEL = "clientesActualizadoCopia.xlsx"

# 🌐 Configurar idioma español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
    except:
        pass

# 🔧 Función para limpiar teléfono
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

# 📊 Leer archivo Excel
df_polizas = pd.read_excel(EXCEL, sheet_name="Polizas")
df_polizas.columns = df_polizas.columns.str.strip()

# Asegurar formato correcto
df_polizas['Dia de Nac'] = pd.to_numeric(df_polizas['Dia de Nac'], errors='coerce')
df_polizas['Mes de Nac'] = pd.to_numeric(df_polizas['Mes de Nac'], errors='coerce')

# 📅 Filtrar cumpleaños del día
hoy = datetime.now()
cumples_hoy = df_polizas[
    (df_polizas['Dia de Nac'] == hoy.day) &
    (df_polizas['Mes de Nac'] == hoy.month)
]

# ✉️ Armar mensajes
mensajes = []
for _, fila in cumples_hoy.iterrows():
    nombre = str(fila.get("Apellido y Nombre", "")).title()
    telefono = limpiar_telefono(fila.get("Telefono", ""))
    if not nombre or not telefono:
        continue

    mensaje = (
        f" ¡Feliz cumpleaños, {nombre}! "
        f"Desde *Grupo Nievas Seguros* te deseamos un día lleno de alegría.\n"
    )

    mensajes.append({
        "nombre": nombre,
        "telefono": telefono,
        "mensaje": mensaje
    })

# 🚀 Enviar mensajes por WhatsApp Web
if mensajes:
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com/")
    input("📲 Escaneá el código QR y presioná ENTER para continuar...")

    for i, m in enumerate(mensajes, start=1):
        if LIMITAR_A and i > LIMITAR_A:
            break

        destino = NUMERO_TEST if MODO_TEST else m["telefono"]
        print(f"📞 Enviando a: {destino}")

        try:
            # Buscar y abrir chat
            search_box = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
            search_box.clear()
            search_box.click()
            search_box.send_keys(destino)
            search_box.send_keys(Keys.ENTER)

            # Esperar caja de mensaje
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
            )
            time.sleep(1)

            # Enviar mensaje
            input_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            input_box.click()
            input_box.send_keys(m["mensaje"])
            input_box.send_keys(Keys.ENTER)

            print(f"✅ Mensaje #{i} {'(TEST)' if MODO_TEST else ''} enviado a {destino}")
            time.sleep(3)

        except Exception as e:
            print(f"❌ Error al enviar a {destino}: {e}")
else:
    print("📭 No hay cumpleaños para hoy.")
