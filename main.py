import pendulum
import datetime
from datetime import datetime,timedelta,time,date
import time
import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import openpyxl
from tkinter import *
from tkinter import ttk
import tkinter as tk

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
def write_file(create_number,create_date,status,full_price,product_name,send_to_house,get_on_house,purchase_price):
    work_table=openpyxl.load_workbook(filename='scaler_table.xlsx')
    worksheet=work_table['Scaler_Place']
    vals = []
    Inspector = 0
    x = ''
    while x != None:
        Inspector = Inspector + 1
        x = worksheet['A' + str(Inspector)].value
        vals.append(x)
    for Inspector in vals:
        if Inspector == create_number:
            worksheet['B' + str(vals.index(Inspector) + 1)] = create_date
            worksheet['C' + str(vals.index(Inspector) + 1)] = status
            worksheet['D' + str(vals.index(Inspector) + 1)] = full_price
            worksheet['E' + str(vals.index(Inspector) + 1)] = product_name
            worksheet['F' + str(vals.index(Inspector) + 1)] = send_to_house
            worksheet['G' + str(vals.index(Inspector) + 1)] = get_on_house
            worksheet['H' + str(vals.index(Inspector) + 1)] = purchase_price
            work_table.save('scaler_table.xlsx')
            return
        elif Inspector == None:
            worksheet['A' + str(vals.index(Inspector) + 1)] = create_number
            worksheet['B' + str(vals.index(Inspector) + 1)] = create_date
            worksheet['C' + str(vals.index(Inspector) + 1)] = status
            worksheet['D' + str(vals.index(Inspector) + 1)] = full_price
            worksheet['E' + str(vals.index(Inspector) + 1)] = product_name
            worksheet['F' + str(vals.index(Inspector) + 1)] = send_to_house
            worksheet['G' + str(vals.index(Inspector) + 1)] = get_on_house
            worksheet['H' + str(vals.index(Inspector) + 1)] = purchase_price
            work_table.save('scaler_table.xlsx')
def main():
    i=0
    url=comber.get()
    if url=="TOPS":
        url='https://business.kazanexpress.ru/seller/4449/invoices/send'
    elif url=="Стельки":
        url='https://business.kazanexpress.ru/seller/65366/invoices/send'
    elif url=='Триколор':
        url='https://business.kazanexpress.ru/seller/10020/invoices/send'
    elif url=='Джибитсы':
        url='https://business.kazanexpress.ru/seller/51310/invoices/send'
    elif url=='Discont OFF':
        url='https://business.kazanexpress.ru/seller/10238/invoices/send'
    end_date = entry_date.get()
    try:
        service = Service(executable_path='chromedriver.exe')
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(options=options)
    except:
        driver = webdriver.ChromiumEdge()
    driver.implicitly_wait(90)
    driver.get(url)
    #time.sleep(15)
    field_finder=driver.find_element(By.XPATH, "//*[@id='username']")
    field_finder.send_keys('f4llno@yandex.ru')
    field_finder=driver.find_element(By.XPATH, "//*[@id='password']")
    field_finder.send_keys('Scaler_Password')
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='signin']/section/div/section[2]/form/button").click()
    time.sleep(2)
    driver.get(url)
    #time.sleep(1.6)
    while True:
        element = driver.find_element(By.XPATH, "//*[@id='status-cell-" + str(i) + "-73-0']/div")
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.find_element(By.XPATH, "//*[@id='status-cell-" + str(i) + "-73-0']/div").click()
        #time.sleep(0.8)
        create_number=(driver.find_element(By.XPATH, "//*[@id='openInvoice']/div/div/div[2]/span[2]").text)
        #print(create_number)
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
            write_file(create_number,create_date,status,full_price,product_name,send_to_house,get_on_house,purchase_price)
            i=i+1
            driver.find_element(By.XPATH,'//*[@id="openInvoice"]/div/header/div[2]').click()
# def scaler_window():
magazines = ['TOPS', 'Стельки', 'Триколор', 'Джибитсы', 'Discont OFF']
window = tk.Tk()
window.title('Scaler')
window ['bg'] = 'gray10'
date_label = tk.Label(text="Введите дату (ДД.ММ.ГГГГ)",background='gray10',foreground='white')
#date_label ['bg'] = 'gray10'
date_label.pack()
entry_date = tk.Entry()
entry_date.pack()
magazine_label = tk.Label(text="Выберите магазин",background='gray10',foreground='white')
#magazine_label ['bg'] = 'gray10'
magazine_label.pack()
comber = ttk.Combobox(values=magazines)
comber.pack()
button = tk.Button(text='Начать выполнение', command=main)
button ['bg'] = 'white'
button.pack()
window.mainloop()
# if __name__=='__main__':
#     scaler_window()
