import os
import urllib
import pydub
import speech_recognition as sr
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from browser import driver

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
    driver.find_element_by_id("recaptcha-demo-submit").click()
    delay()