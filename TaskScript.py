from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client as winClient
import time
import threading
from config import CAC_PIN, WEBSITE_URL
import glob
import os


certSelect = winClient.Dispatch('WScript.Shell')
driver = webdriver.Chrome('./chromedriver')

def sleepCt(num, pt=False):
    if pt:
        print(f'wait {num} sec')
    for i in range(num,0,-1):
        time.sleep(1)
        print(i)
    
def goToSite():
    # Using chromedriver.exe create an instance of chrome
    # Open a website
    print('going to site...')
    driver.get(WEBSITE_URL)

def cacAuth():
    sleepCt(2,True)
    print('pressing ENTER...')
    certSelect.SendKeys("{ENTER}")
    sleepCt(4,True)
    print('entering pin...')
    certSelect.SendKeys(CAC_PIN)
    certSelect.SendKeys("{ENTER}")

# starting the website in a separate thread as the 
# cac card prompt freezes the thread
thread = threading.Thread(target = cacAuth)
thread.start()

goToSite()
sleepCt(7,True)
certSelect.SendKeys("{ENTER}")

sleepCt(2,True)
driver.find_element_by_id('desktopsBtn').click()
driver.find_element_by_xpath("//img[@alt='VACO Server Farm Desktop']").click()

# once file is downloaded, start up the desktop
# finding the most recent file in the downloads file for chrome
sleepCt(3,True)
list_of_files = glob.glob('C:/Users/Matt/Downloads/*.ica') 
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)

os.startfile(latest_file)