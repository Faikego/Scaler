import pendulum
import datetime
from datetime import datetime,timedelta,time,date
import time
import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
end_date='01.08.2023'
def date_comparison(x_date,y_date):
    def dot_seeker(dot,ifer):
        if ifer == 1:
            dot_finder=dot.rfind('.')
            dot_findes=dot.find('.')
            dot = dot[dot_findes+1:dot_finder]
            try:
                dot=int(dot)
            except ValueError:
                dot=dot[1:]
        elif ifer == 0:
            dot_finder=dot.find('.')
            dot=dot[:dot_finder]
        try:
            dot = int(dot)
        except ValueError:
            dot = dot[1:]
        return dot
    day_x=dot_seeker(x_date,0)
    day_y=dot_seeker(y_date,0)
    month_x=dot_seeker(x_date,1)
    month_y=dot_seeker(y_date,1)
    if month_y>month_x:
        returner = True
        return returner
    elif month_y==month_x:
        if day_y>day_x:
            returner = True
            return returner
        elif day_y==day_x:
            returner = True
            return returner
        else:
            returner = False
            return returner
    else:
        returner = False
        return returner
status_list=[]
full_price_list=[]
create_date_list=[]
create_number_list=[]
service = Service(executable_path='C:\chromedriver\chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)
driver.get('https://business.kazanexpress.ru/seller/4449/invoices/send')
time.sleep(15)
field_finder=driver.find_element(By.XPATH, "//*[@id='username']")
field_finder.send_keys('f4llno@yandex.ru')
field_finder=driver.find_element(By.XPATH, "//*[@id='password']")
field_finder.send_keys('Pravoe_del0')
time.sleep(1)
field_finder=driver.find_element(By.XPATH, "//*[@id='signin']/section/div/section[2]/form/button").click()
time.sleep(7)
driver.get('https://business.kazanexpress.ru/seller/4449/invoices/send')
time.sleep(3)
field_finder=driver.find_element(By.XPATH, "//*[@id='status-cell-0-73-0']/div").click()
time.sleep(3)
create_number_list.append(driver.find_element(By.XPATH, "//*[@id='openInvoice']/div/div/div[2]/span[2]").text)
create_date=driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[4]').text
create_date_list.append(create_date)
status_list.append(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[6]').text)
full_price_list.append(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[8]').text)
print(create_number_list)
print(create_date_list)
print(status_list)
print(full_price_list)
solution=date_comparison(create_date,end_date)
print(solution)

# login XPATH - //*[@id="username"]
# password XPATH - //*[@id="password"]
# start button //*[@id="signin"]/section/div/section[2]/form/button
