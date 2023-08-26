import pendulum
import datetime
from datetime import datetime,timedelta,time,date
import time
import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import openpyxl
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
def main():
    i=0
    end_date = '24.08.2023'
    service = Service(executable_path='\chromedriver.exe')
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    driver.get('https://business.kazanexpress.ru/seller/4449/invoices/send')
    time.sleep(15)
    field_finder=driver.find_element(By.XPATH, "//*[@id='username']")
    field_finder.send_keys('f4llno@yandex.ru')
    field_finder=driver.find_element(By.XPATH, "//*[@id='password']")
    field_finder.send_keys('Pravoe_del0')
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='signin']/section/div/section[2]/form/button").click()
    time.sleep(7)
    driver.get('https://business.kazanexpress.ru/seller/4449/invoices/send')
    time.sleep(1)
    work_table=openpyxl.load_workbook(filename='scaler_table.xlsx')
    worksheet=work_table['Scaler_Place']
    abc=['A','B','C','D','E','F','G','H','Z']
    Litera=''
    while True:
        element=driver.find_element(By.XPATH, "//*[@id='status-cell-"+str(i)+"-73-0']/div")
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.find_element(By.XPATH, "//*[@id='status-cell-" + str(i) + "-73-0']/div").click()
        time.sleep(1)
        create_number=(driver.find_element(By.XPATH, "//*[@id='openInvoice']/div/div/div[2]/span[2]").text)
        create_date=driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[4]').text
        status=driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[6]').text
        if status=='Принята на складе':
            send_to_house=(driver.find_element(By.XPATH, '//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[3]/div').text)
            get_on_house=(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div').text)
            purchase_price=(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[5]/div/span/em[1]').text)
        else:
            send_to_house=(driver.find_element(By.XPATH, '//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[3]/div').text)
            purchase_price=(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[4]/div/span/em[1]').text)
            get_on_house=0
        full_price=(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[2]/span[8]').text)
        product_name=(driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[1]/td[2]/div').text)
        if create_date==end_date:
            return
        else:
            vals = []
            i = 0
            x = ''
            while x != None:
                i = i + 1
                x = worksheet['A' + str(i)].value
                vals.append(x)
            for Inspector in vals:
                if Inspector==create_number:
                    worksheet['A'+str(vals.index(Inspector)+1)]=create_number
                    worksheet['B' + str(vals.index(Inspector) + 1)]=create_date
                    worksheet['C' + str(vals.index(Inspector) + 1)]=status
                    worksheet['D' + str(vals.index(Inspector) + 1)]=full_price
                    worksheet['E' + str(vals.index(Inspector) + 1)]=product_name
                    worksheet['F' + str(vals.index(Inspector) + 1)]=send_to_house
                    worksheet['G' + str(vals.index(Inspector) + 1)]=get_on_house
                    worksheet['H' + str(vals.index(Inspector) + 1)]=purchase_price
                elif Inspector==None:
                    worksheet['A'+str(vals.index(Inspector)+1)]=create_number
                    worksheet['B' + str(vals.index(Inspector) + 1)]=create_date
                    worksheet['C' + str(vals.index(Inspector) + 1)]=status
                    worksheet['D' + str(vals.index(Inspector) + 1)]=full_price
                    worksheet['E' + str(vals.index(Inspector) + 1)]=product_name
                    worksheet['F' + str(vals.index(Inspector) + 1)]=send_to_house
                    worksheet['G' + str(vals.index(Inspector) + 1)]=get_on_house
                    worksheet['H' + str(vals.index(Inspector) + 1)]=purchase_price
            # while Litera!='Z':
            #     Litera=abc[Inspector]
            work_table.save('scaler_table.xlsx')
            i=i+1
            driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/header/div[2]').click()
if __name__=='__main__':
    main()

# //*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[2]/td[5]/div/span/em[1]
# //*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[3]/td[5]/div/span/em[1]
# //*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[2]/td[4]/div/span/em[1]
# //*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[2]/td[2]/div
# //*[@id="openInvoice"]/div/div/div[3]/div[2]/div/table/tbody/tr[3]/td[2]/div