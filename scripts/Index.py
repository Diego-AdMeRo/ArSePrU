import os
import json
from time import sleep
from Sheets import conexion_sheets
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys


def cargar_informacion(ruta):
    with open(ruta, "r", encoding="UTF-8") as info:
        return json.load(info)


def guardar_informacion(ruta, informacion):
    with open(ruta, "w", encoding="UTF-8") as info:
        info.write(json.dumps(informacion, ensure_ascii=False))


def ingreso_informacion(fila, controlador):
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.ID, "movenextbtn"))).click()

    # Selección de Carrera y Estudiante
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8377X107808']/option[contains(.,'{informacion['Estudiante']['Carrera']}')]"))).click()

    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8377X108245']/option[contains(.,'{informacion['Estudiante']['Nombre']}')]"))).click()

    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.ID, "movenextbtn"))).click()

    # Ingreso de Parametros
    # Semana
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107799']/option[contains(.,'{fila[0]}')]"))).click()

    # Proyecto
    tipo = fila[1]
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107811']/option[contains(.,'{tipo}')]"))).click()

    if tipo == "Otro":
        WebDriverWait(controlador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107819']/option[contains(.,'{fila[2]}')]"))).click()

    # Actividad General
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107800']/option[contains(.,'{fila[3]}')]"))).click()

    # Tiempo
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107803']/option[contains(.,'{fila[4]}')]"))).click()

    # Herramientas
    tipo = fila[5]
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107802']/option[contains(.,'{tipo}')]"))).click()
    if tipo == "Otro:":
        WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='othertext752585X8378X107802']"))).send_keys(fila[6])

    # Entregable
    tipo = fila[7]
    WebDriverWait(controlador, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[@id='answer752585X8378X107801']/option[contains(.,'{tipo}')]"))).click()

    if tipo == "Otro:":
        WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='othertext752585X8378X107801']"))).send_keys(fila[8])

    # Descripción de Actividad
    WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='answer752585X8378X107807']"))).send_keys(fila[9])

    sleep(10)

    WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='movenextbtn']"))).click()

    WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='javatbd752585X8379X107805A1']"))).click()
        #//*[@id="answer752585X8379X107805sA1"]

    WebDriverWait(controlador, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='movesubmitbtn']"))).click()


if __name__ == "__main__":

    # 1. Almacenamiento de Ruta de Informacion
    rutaInformacion = os.path.join(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))), "docs/encuesta.json")

    # 2. Cargado de datos (credenciales, url's, etc.)
    informacion = cargar_informacion(ruta=rutaInformacion)

    # 3. Obtención de Información
    conexion = conexion_sheets(
        informacion["Google Sheets"]["Token"], informacion["Google Sheets"]["Credentials"])
    sheets = conexion.spreadsheets()
    query = sheets.values().get(
        spreadsheetId=informacion["Google Sheets"]["Url"], range=f"Semana {informacion['Google Sheets']['Semana']}!A{informacion['Google Sheets']['Linea']}:J").execute()
    datos = query.get('values', [])

    # 4. Configuración de Driver Chrome
    opciones = webdriver.ChromeOptions()
    opciones.add_experimental_option('excludeSwitches', ['enable-logging'])

    controlador = webdriver.Chrome(
        informacion["WebDriver"]["Path"], options=opciones)
    controlador.maximize_window()

    # 5. Llenado de Información
    fila, continuar, pestanias = 0, True, controlador.window_handles

    while fila < len(datos) and continuar:
        controlador.get(informacion["Plataforma"]["Url"])

        ingreso_informacion(datos[fila], controlador)

        controlador.execute_script('''window.open("about:blank", "_blank");''')
        fila += 1
        controlador.switch_to.window(controlador.window_handles[fila])
        sleep(10)

    informacion["Google Sheets"]["Semana"] = informacion["Google Sheets"]["Semana"] + 1
    guardar_informacion(rutaInformacion, informacion)
    controlador.close()
