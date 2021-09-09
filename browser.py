from os import name
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import random
import os
import urllib
import pydub
import speech_recognition as sr


def open_browser( dis_nav_web, win_size, user_agent, domain):
    option = webdriver.ChromeOptions()
    
    option.add_argument(dis_nav_web)
    option.add_argument(win_size)
    option.add_argument(user_agent)      
    global driver
    driver = webdriver.Chrome(options=option)
    driver.get(domain)

def bypass():
    def delay(waiting_time=5):
        driver.implicitly_wait(waiting_time)

    # switch to recaptcha frame
    frames = driver.find_elements_by_tag_name("iframe")
    driver.switch_to.frame(frames[0])
    delay()

    # click on checkbox to activate recaptcha
    driver.find_element_by_class_name("recaptcha-checkbox-border").click()

    # switch to recaptcha audio control frame
    driver.switch_to.default_content()
    frames = driver.find_element_by_xpath("/html/body/div[2]/div[4]").find_elements_by_tag_name("iframe")
    driver.switch_to.frame(frames[0])
    delay()

    # click on audio challenge
    driver.find_element_by_id("recaptcha-audio-button").click()

    # switch to recaptcha audio challenge frame
    driver.switch_to.default_content()
    frames = driver.find_elements_by_tag_name("iframe")
    driver.switch_to.frame(frames[-1])
    delay()

    # get the mp3 audio file
    src = driver.find_element_by_id("audio-source").get_attribute("src")
    print("[INFO] Audio src: %s" % src)

    # download the mp3 audio file from the source
    urllib.request.urlretrieve(src, os.path.normpath(os.getcwd() + "\\sample.mp3"))
    delay()

    # load downloaded mp3 audio file as .wav
    try:
        sound = pydub.AudioSegment.from_mp3(os.path.normpath(os.getcwd() + "\\sample.mp3"))
        sound.export(os.path.normpath(os.getcwd() + "\\sample.wav"), format="wav")
        sample_audio = sr.AudioFile(os.path.normpath(os.getcwd() + "\\sample.wav"))
    except Exception:
        print("[-] Please run program as administrator or download ffmpeg manually, "
                "http://blog.gregzaal.com/how-to-install-ffmpeg-on-windows/")

    # translate audio to text with google voice recognition
    r = sr.Recognizer()
    with sample_audio as source:
        audio = r.record(source)
    key = r.recognize_google(audio)
    print("[INFO] Recaptcha Passcode: %s" % key)

    # key in results and submit
    driver.find_element_by_id("audio-response").send_keys(key.lower())
    driver.find_element_by_id("audio-response").send_keys(Keys.ENTER)
    driver.switch_to.default_content()
    delay()

def switch_frames():
    driver.switch_to.default_content()

class Element:
    def __init__(self, excel_ruc, excel_dv, excel_tipo, excel_timbrado, excel_doc, excel_nd1, excel_nd2, excel_nd3, excel_fecha):
        self.excel_ruc = excel_ruc 
        self.excel_dv = excel_dv
        self.excel_tipo = excel_tipo
        self.excel_timbrado = excel_timbrado
        self.excel_doc = excel_doc
        self.excel_nd1 = excel_nd1
        self.excel_nd2 = excel_nd2
        self.excel_nd3 = excel_nd3
        self.excel_fecha = excel_fecha

    def text_keys(self):
        time.sleep(random.randint(1, 2))
        #CARGAR RUC
        ruc = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/div[2]/div[1]/div/input')
        ruc.click()
        ruc.send_keys(self.excel_ruc) 
        
        time.sleep(random.randint(1, 2))
        #CARGAR DV
        dv = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/div[2]/div[2]/input')
        dv.click()
        dv.send_keys(self.excel_dv)

        time.sleep(random.randint(1, 2))
        # CARGAR NOMBRE
        nombre = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/div[3]/div/input')
        nombre.click()
    
    #TIPO DE CONSULTA
    def select_keys(self):
        self.tipo = Select(driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/div[5]/div/select'))
    
    # TIPO DOCUMENTO AUTOIMPRESORES / PREIMPRESOS
        if self.excel_tipo == 'AUTOIMPRESORES / PREIMPRESOS':
            self.tipo.select_by_visible_text('AUTOIMPRESORES / PREIMPRESOS')
            
            time.sleep(random.randint(1, 2))
            timbrado = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[1]/input')
            timbrado.click()
            timbrado.send_keys(self.excel_timbrado)

            doc = Select(driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[2]/select'))
            if self.excel_doc == 'FACTURA':
                doc.select_by_visible_text('FACTURA')
            elif self.excel_doc == 'AUTOFACTURAS':
                doc.select_by_visible_text('AUTOFACTURAS')
            elif self.excel_doc == 'BOLETA DE VENTA':
                doc.select_by_visible_text('BOLETA DE VENTA')
            elif self.excel_doc == 'BOLETOS DE TRANSPORTE AEREO':
                doc.select_by_visible_text('BOLETOS DE TRANSPORTE AEREO')
            elif self.excel_doc == 'NOTA DE REMISION':
                doc.select_by_visible_text('NOTA DE REMISION')
            elif self.excel_doc == 'NOTA DE CREDITO':
                doc.select_by_visible_text('NOTA DE CREDITO')
            elif self.excel_doc == 'NOTA DE DEBITO':
                doc.select_by_visible_text('NOTA DE DEBITO')
            elif self.excel_doc == 'COMPROBANTES DE RETENCION':
                doc.select_by_visible_text('COMPROBANTES DE RETENCION')
            elif self.excel_doc == 'BOLETA RESIMPLE':
                doc.select_by_visible_text('BOLETA RESIMPLE')
            elif self.excel_doc == 'BOLETOS DEL TRANSPORTE PUBLICO':
                doc.select_by_visible_text('BOLETOS DEL TRANSPORTE PUBLICO')
            elif self.excel_doc == 'BOLETOS DE LOTERIAS, JUEGOS DE AZAR':
                doc.select_by_visible_text('BOLETOS DE LOTERIAS, JUEGOS DE AZAR')
            elif self.excel_doc == 'ENTRADAS A ESPECTACULOS PUBLICOS':
                doc.select_by_visible_text('ENTRADAS A ESPECTACULOS PUBLICOS')
            else:
                print('Error: TIPO DE DOCUMENTO')
                exit()
            
            time.sleep(random.randint(1, 2))
            nd1 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[1]')
            nd1.click()
            nd1.send_keys(self.excel_nd1)

            nd2 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[2]')
            nd2.click()
            nd2.send_keys(self.excel_nd2)

            nd3 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[3]')
            nd3.click()
            nd3.send_keys(self.excel_nd3)

            time.sleep(random.randint(1, 2))
            fecha = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[2]/div/input')
            fecha.send_keys(self.excel_fecha)
        
    # TIPO DOCUMENTO MÁQUINAS REGISTRADORAS    
        elif self.excel_tipo == 'MÁQUINAS REGISTRADORAS':
            self.tipo.select_by_visible_text('MÁQUINAS REGISTRADORAS')
            
            timbrado = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div/div[1]/input')
            timbrado.click()
            timbrado.send_keys(self.excel_timbrado)
            
            time.sleep(random.randint(1, 2))
            fecha_exp_ticket = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div/div[2]/div/input')
            fecha_exp_ticket.send_keys(self.excel_fecha)   
    
    # TIPO DOCUMENTO COMPROBANTES VIRTUALES
        elif self.excel_tipo == 'COMPROBANTES VIRTUALES':
            self.tipo.select_by_visible_text('COMPROBANTES VIRTUALES')
            
            time.sleep(random.randint(1, 2))
            timbrado = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[1]/input')
            timbrado.click()
            timbrado.send_keys(self.excel_timbrado)

            doc = Select(driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[2]/select'))
            if self.excel_doc == 'FACTURA VIRTUAL':
                doc.select_by_visible_text('FACTURA VIRTUAL')
            elif self.excel_doc == 'COMPROBANTE DE RETENCION VIRTUA':
                doc.select_by_visible_text('COMPROBANTE DE RETENCION VIRTUA')
            else:
                print('Error: TIPO DE DOCUMENTO')

            time.sleep(random.randint(1, 2))        
            nd1 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[1]')
            nd1.click()
            nd1.send_keys(self.excel_nd1)

            nd2 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[2]')
            nd2.click()
            nd2.send_keys(self.excel_nd2)
            
            nd3 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[3]')
            nd3.click()
            nd3.send_keys(self.excel_nd3)

            time.sleep(random.randint(1, 2))
            fecha = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[2]/div/input')
            fecha.send_keys(self.excel_fecha)
    
    # TIPO DOCUMENTO DOCUMENTOS ELECTRONICOSS
        elif self.excel_tipo == 'DOCUMENTOS ELECTRONICOS':
            self.tipo.select_by_visible_text('DOCUMENTOS ELECTRONICOS')

            time.sleep(random.randint(1, 2))
            timbrado = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[1]/input')
            timbrado.click()
            timbrado.send_keys(self.excel_timbrado)

            doc = Select(driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[1]/div[2]/select'))
            if self.excel_doc == 'NOTA DE DEBITO ELECTRONICA':
                doc.select_by_visible_text('NOTA DE DEBITO ELECTRONICA')
            elif self.excel_doc == 'FACTURA ELECTRONICA':
                doc.select_by_visible_text('FACTURA ELECTRONICA')
            elif self.excel_doc == 'NOTA DE CREDITO ELECTRONICA':
                doc.select_by_visible_text('NOTA DE CREDITO ELECTRONICA')
            elif self.excel_doc == 'AUTOFACTURA ELECTRONICA':
                doc.select_by_visible_text('AUTOFACTURA ELECTRONICA')
            elif self.excel_doc == 'NOTA DE REMISION ELECTRONICA':
                doc.select_by_visible_text('NOTA DE REMISION ELECTRONICA')
            elif self.excel_doc == 'BOLETA RESIMPLE ELECTRONICA':
                doc.select_by_visible_text('BOLETA RESIMPLE ELECTRONICA')
            elif self.excel_doc == 'BOLETA DE VENTA ELECTRONICA':
                doc.select_by_visible_text('BOLETA DE VENTA ELECTRONICA')
            else:
                print('Error: TIPO DE DOCUMENTO')
            
            time.sleep(random.randint(1, 2))
            nd1 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[1]')
            nd1.click()
            nd1.send_keys(self.excel_nd1)

            nd2 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[2]')
            nd2.click()
            nd2.send_keys(self.excel_nd2)

            nd3 = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[1]/div/input[3]')
            nd3.click()
            nd3.send_keys(self.excel_nd3)

            time.sleep(random.randint(1, 2))
            fecha = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/div/div[2]/div/div/form/section[1]/div[2]/div[2]/div/input')
            fecha.send_keys(self.excel_fecha)
    
        else:
            print('Error: TIPO DE CONSULTA')

class Comprobante:
    def __init__(self, set_val, validaciones):
        self.set_val = set_val
        self.validaciones = validaciones

    def ruc_error(self):
        try:
            notification = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div/div')
            self.set_val = notification.text
            return self.set_val
        except:
            pass

    def consult_comp(self):
        try:
            aceptar = driver.find_element_by_css_selector('#botones > div > div > button')
            aceptar.click()
            time.sleep(1)
            comprobante = driver.find_element_by_xpath('//*[@id="pico-3"]/div/div/div/div[1]/p')
            self.set_val = comprobante.text
            
            aceptar = driver.find_element_by_xpath('//*[@id="pico-3"]/div/div/div/div[2]/button')
            aceptar.click()
            
            return self.set_val
        except:
            pass

            driver.close()
    
    def create_list(self):
        self.validaciones.append(self.set_val)

    def terminar_sesion(self):
        driver.quit()