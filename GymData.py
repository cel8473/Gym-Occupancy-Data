from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime as dt
from datetime import timedelta as td
import xlsxwriter

#Initialize workbook
today = dt.now() 
timeString = today.strftime("%m_%d_%Y")
gymBook = xlsxwriter.Workbook('GymOccupancy' + timeString + '.xlsx')
gymSheet = gymBook.add_worksheet()
i = 1
while(i <= 96): #24 hrs every 15 min is 96 times
    #Pull Occupancy from RIT gym website 
    fifteenMinute = int(dt.now().strftime("%M")) % 15
    if(fifteenMinute == 0):
    #Pull Occupancy from RIT gym website 
        try:
            PATH = "C:\Program Files (x86)\chromedriver.exe"
            driver = webdriver.Chrome(PATH)
            driver.get("https://recreation.rit.edu/facilityoccupancy")

            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"occupancy-cf65cbcd-c559-4c6c-83e6-1d4fc886b543-sm\"]/div[2]/p[3]")))
            element=driver.find_element_by_xpath("//*[@id=\"occupancy-cf65cbcd-c559-4c6c-83e6-1d4fc886b543-sm\"]/div[2]/p[3]/strong")
            currentOccupancy=element.text
            time.sleep(10)
            driver.quit()
        except:
            print("The program failed, check selenium or ChromeDriver updates")
            driver.quit()
            input("Press Enter to Continue")
    #Put the occupancy in the sheet   
        print(currentOccupancy)
        column = 'A' + str(i)
        gymSheet.write(column, currentOccupancy)
        i+=1
        time.sleep(720) #Wait 12 min, there is no need to check for a while

#Close and save .xslx file
print("Closing book")
gymBook.close()
