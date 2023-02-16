from ast import If
import configparser
from re import T
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd
from datetime import date
import pyodbc
from fast_to_sql import fast_to_sql as fts

#---------------------------------------------------------------------------
# Conecta a SQL
def SQL_conexion (server, database):
    SQLConn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server} ;"
                     "Server=" + server + ";"
                     "Database=" + database + ";"
                     "Trusted_Connection=yes;")

    return SQLConn  #.cursor()
#-----------------------------------------------------
def get_matriculados (URL, driver):

    driver.get(URL)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="example_wrapper"]/div[3]/div[2]')))

    # Crea el dataframe para guardar los datos
    Matriculados = pd.DataFrame (columns=['Matricula','Nombre','Actividad','Telefono','Mail','Sitio','Localidad', 'Provincia'])
    selector = driver.find_elements(by=By.XPATH, value='//*[@id="i-tipo"]/option')

    # Recorre lista de matriculados, actividad por actividad
    for i in range(len(selector)):
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="example_wrapper"]/div[3]/div[2]')))  
        Select (driver.find_element(by=By.NAME, value='cmb_filtro')).select_by_index(i)
        driver.find_element(by=By.XPATH, value='//*[@id="subhead-matriculados"]/div[2]/form/div[3]/button').click()
        Select (driver.find_element(by=By.NAME, value='example_length')).select_by_index(3)
        time.sleep(5)
        element = driver.find_elements(by=By.XPATH, value='//*[@id="example"]/tbody/tr')
        pages = len(driver.find_elements(by=By.XPATH, value='//*[@id="example_paginate"]/span/a'))
        actual = 1
        while True:

            for i in element:
                fila = i.find_elements(by=By.TAG_NAME, value="td")
                Matriculados = pd.concat ([Matriculados,pd.Series([fila[0].text,fila[1].text,fila[2].text,fila[3].text ,fila[4].text,fila[5].text,fila[6].text,fila[7].text],index = Matriculados.columns).to_frame().T])
                print(fila[0].text,fila[1].text,fila[2].text,fila[3].text ,fila[4].text,fila[5].text,fila[6].text,fila[7].text)

            if actual < pages:
                actual += 1
                btn = driver.find_element(by=By.ID, value='example_next')
                statusbtn = btn.get_attribute("class")
                if (statusbtn == 'paginate_button next') :
                    print('next page')
                    driver.find_element(by=By.XPATH, value='/html/body/div/div/div[3]/div/div[2]/div/div[4]/div[5]/a[2]').click()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="example_wrapper"]/div[3]/div[2]')))
                    element = driver.find_elements(by=By.XPATH, value='//*[@id="example"]/tbody/tr')
                else:
                    break
            else:
                break
    
    return Matriculados


#---------------------------------------------------------------------------------
def graba_sql (df, Conn):

    cursor = Conn.cursor()

    cursor.execute("SET ANSI_WARNINGS  OFF")
    cursor.commit()


    create_statement = fts.fast_to_sql(df, 'dbo.martillerosCts', Conn, if_exists='replace')
    Conn.commit()


#-----------------------------------------------------
cp = configparser.ConfigParser()
cp.read("config.ini")
URL = cp["DEFAULT"]["URL"]
Server_Origen = cp["DEFAULT"]["server_origen"]
Base_Origen= cp["DEFAULT"]["base_origen"]

Conn = SQL_conexion(Server_Origen, Base_Origen)

# Configura Chromedriver
options = webdriver.ChromeOptions()
options.headless = False
preferences = { "download.directory_upgrade": True,
                "safebrowsing_for_trusted_sources_enabled": False,
                "safebrowsing.enabled": False,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "download.prompt_for_download": False }
options.add_argument("--ignore-certificate-errors")
options.add_experimental_option("prefs", preferences)
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_argument('--kiosk-printing')
options.add_argument('--disable-gpu')
options.add_argument('--disable-software-rasterizer')
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)

df= get_matriculados(URL, driver)

graba_sql(df, Conn)
