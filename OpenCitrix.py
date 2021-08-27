from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client as winClient
import time
import threading
from config import CAC_PIN, WEBSITE_URL
import glob
import os

def main():
    # count down timers to print
    yell = False

    certSelect = winClient.Dispatch('WScript.Shell')
    driver = webdriver.Chrome(r'C:\Users\Matt\Desktop\autoPython-main\chromedriver.exe')

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
        sleepCt(3,yell)
        print('pressing ENTER...')
        certSelect.SendKeys("{ENTER}")
        sleepCt(4,yell)
        print('entering pin...')
        certSelect.SendKeys(CAC_PIN)
        certSelect.SendKeys("{ENTER}")

    # starting the website in a separate thread as the 
    # cac card prompt freezes the thread
    thread = threading.Thread(target = cacAuth)
    thread.start()

    goToSite()
    sleepCt(7,yell)
    certSelect.SendKeys("{ENTER}")

    sleepCt(3,yell)
    driver.find_element_by_id('desktopsBtn').click()
    sleepCt(1)
    driver.find_element_by_xpath("//img[@alt='VACO Server Farm Desktop']").click()
    sleepCt(7,yell)

    # once file is downloaded, start up the desktop
    # finding the most recent file in the downloads file for chrome
    list_of_files = glob.glob('C:/Users/Matt/Downloads/*.ica') 
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)

    os.startfile(latest_file)
    
    sleepCt(3,yell)

    driver.close

if __name__ == "__main__":
    main()