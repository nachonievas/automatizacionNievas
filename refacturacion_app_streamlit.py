
import streamlit as st
import pandas as pd
from datetime import datetime
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from dateutil.relativedelta import relativedelta
import locale

# Configurar idioma
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
    except:
        pass

st.title("Refacturación Automática por WhatsApp")

archivo = st.file_uploader("📁 Cargá el archivo Excel", type=["xlsx"])
modo_test = st.checkbox("Modo TEST (solo se envía a tu número)", value=True)
numero_test = st.text_input("Número de prueba (formato 549351xxxxxxx)", "5493516570658")
tipo_refacturacion = st.selectbox("Seleccioná el tipo de refacturación", ["Mensual", "Trimestral"])

iniciar = st.button("🚀 Iniciar envío de mensajes")

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
        if "holando" in compania.lower():
            return fecha.strftime("%B").capitalize()
        else:
            return (fecha + relativedelta(months=1)).strftime("%B").capitalize()
    except:
        return "próximos días"

if iniciar and archivo is not None:
    df_estado = pd.read_excel(archivo, sheet_name="Estado de cuenta")
    df_polizas = pd.read_excel(archivo, sheet_name="Polizas", dtype={"Telefono": str})
    df_estado.columns = df_estado.columns.str.strip().str.lower()
    df_polizas.columns = df_polizas.columns.str.strip().str.lower()

    refacturacion_objetivo = tipo_refacturacion.lower()
    mensajes = []
    pendientes = []

    for i in range(min(len(df_polizas), len(df_estado))):
        fila_poliza = df_polizas.iloc[i]
        fila_estado = df_estado.iloc[i]

        refacturacion = str(fila_estado.get("refacturación", "")).strip().lower()
        estado = str(fila_estado.get("estado", "")).strip().upper()

        if refacturacion != refacturacion_objetivo:
            continue

        if estado != "SI":
            pendientes.append({
                "index": i + 1,
                "apellido y nombre": fila_estado.get("apellido y nombre", ""),
                "dni": fila_poliza.get("dni", ""),
                "telefono": fila_poliza.get("telefono", ""),
                "compañía": fila_estado.get("compañía", ""),
                "riesgo": fila_poliza.get("riesgo", ""),
                "estado": estado,
                "refacturación": refacturacion,
                "motivo": "Estado distinto de SI"
            })
            continue

        nombre = str(fila_estado.get("apellido y nombre", "")).title()
        compania = str(fila_estado.get("compañía", "")).strip()
        riesgo = str(fila_poliza.get("riesgo", "")).strip().lower()
        telefono = limpiar_telefono(fila_poliza.get("telefono", ""))
        fecha = fila_estado.get("flyer", "")
        cuota = fila_estado.get("cuota", 0)
        suma_asegurada = fila_estado.get("suma asegurada2", 0)

        mes = obtener_mes_espanol(fecha, compania)
        cuota_fmt = f"${cuota:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        suma_fmt = f"${suma_asegurada:,.0f}".replace(",", ".")

        mensaje = (
            f"Hola {nombre}, te recordamos que en el mes de *{mes}* se debitará tu póliza de *{compania}*, "
            f"correspondiente al seguro de *{riesgo.title()}*."
            f"La refacturación de esta póliza es *{refacturacion_objetivo}*."
            f" Cuota: *{cuota_fmt}*"
            f" Suma asegurada: *{suma_fmt}*"
            f"¡Gracias por confiar en nosotros!"
        )

        mensajes.append({
            "index": i + 1,
            "telefono": telefono,
            "compañia": compania,
            "mensaje": mensaje
        })

    st.success(f"Mensajes listos para enviar: {len(mensajes)}")
    st.warning(f"Pendientes (estado ≠ SI): {len(pendientes)}")

    if len(mensajes) > 0:
        driver = webdriver.Chrome()
        driver.get("https://web.whatsapp.com/")
        st.info("Escaneá el código QR en el navegador y luego volvé aquí.")
        input("Presioná ENTER para continuar una vez escaneado el QR...")

        for m in mensajes:
            destino = numero_test if modo_test else m["telefono"]
            st.write(f"Enviando a: {destino}")
            try:
                search_box = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
                )
                time.sleep(2)
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
                time.sleep(2)

            except Exception as e:
                st.error(f"Error al enviar a {destino}: {e}")
