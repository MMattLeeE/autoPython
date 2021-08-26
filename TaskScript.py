from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client as winClient
import time
import threading

driver = webdriver.Chrome('./chromedriver')
certSelect = winClient.Dispatch('WScript.Shell')

def threaded_goToSite():
    # Using chromedriver.exe create an instance of chrome
    # Open a website
    driver.get('https://citrixaccesspiv.va.gov/Citrix/StoreWeb/')
    print('opening window and getting cac card cert prompt')

def threaded_pressEnter():
    print('start press enter')
    print('wait 5 sec')
    time.sleep(5)
    #certSelect.AppActivate("Chrome")
    print('pressing ENTER...')
    certSelect.SendKeys("{ENTER}")

#Calling the website and pressing 10 times in the same time
thread2 = threading.Thread(target = threaded_pressEnter)
thread2.start()

thread = threading.Thread(target = threaded_goToSite)
thread.start()