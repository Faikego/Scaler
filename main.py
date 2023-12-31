from datetime import time
import time
import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from tkinter import ttk
import tkinter as tk
from file_worker import write_file,library_converter,get_key,sorter_created,sorter_other,dot_seeker
from changer import changer

def main(): #Главный скрипт по парсингу актов, запускается по кнопке
    def parser():
        i = 0
        while True:
            element = driver.find_element(By.XPATH, "//*[@id='status-cell-" + str(i) + "-69-0']/div")
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            element.click()
            time.sleep(0.8)
            # info_elements= driver.find_element(By.CSS_SELECTOR, "#openInvoice > div > div > div.invoice-info").text
            # print(info_elements)
            info_elements_list= driver.find_elements(By.CLASS_NAME, 'invoice-info__title')
            if len(info_elements_list) == 4:
                create_number = (driver.find_element(By.CSS_SELECTOR, "#openInvoice > div > div > div.invoice-info > span:nth-child(2)").text)
                print(create_number)
                create_date = driver.find_element(By.XPATH, '//*[@id="openInvoice"]/div/div/div[2]/span[4]').text
                status = driver.find_element(By.CSS_SELECTOR, '#openInvoice > div > div > div.invoice-info > span:nth-child(6)').text
                print(status)
            else:
                create_number = (driver.find_element(By.CSS_SELECTOR, "#openInvoice > div > div > div.invoice-info > span:nth-child(2)").text)
                print(create_number)
                create_date = driver.find_element(By.XPATH, '//*[@id="openInvoice"]/div/div/div[2]/span[4]').text
                status = driver.find_element(By.CSS_SELECTOR, '#openInvoice > div > div > div.invoice-info > span:nth-child(8)').text
                print(status)
            if end_month == dot_seeker(create_date) and status != 'Создана':
                print('Работа закончена...')
                return
            #if status == 'Принята на складе' and
            if status == 'Принята на складе':
                element_library=[]
                elements=driver.find_elements(By.CSS_SELECTOR,'.modal-card [data-test-id="table__body-row"] .data-container')
                for element in elements:
                    element_library.append(element.text)
                temp_get_on_house,temp_send_to_house=sorter_other(element_library)
                send_to_house = library_converter(temp_send_to_house)
                get_on_house = library_converter(temp_get_on_house)
                product_name= element_library[0]
                write_file(create_number, create_date, status, product_name, send_to_house, get_on_house, comber.get())
            elif status == 'Отменена':
                False
            else:
                element_library = []
                elements = driver.find_elements(By.CSS_SELECTOR,'.modal-card [data-test-id="table__body-row"] .data-container')
                for element in elements:
                    element_library.append(element.text)
                temp_send_to_house = sorter_created(element_library)
                send_to_house = library_converter(temp_send_to_house)
                get_on_house = 0
                product_name = element_library[0]
                if get_on_house == 0:
                    write_file(create_number, create_date, status, product_name, send_to_house, get_on_house,comber.get())
                else:
                    write_file(create_number, create_date, status, product_name, send_to_house, get_on_house,comber.get())
            i = i + 1
            driver.find_element(By.XPATH, '//*[@id="openInvoice"]/div/header/div[2]').click()
            time.sleep(0.3)
    if checker_debugger_var.get()==1: #Проверка на нажатие "Debugging" уменьшает время ожидания от интерфейса до 15 секунд
        waiting=15
    elif checker_debugger_var.get()==0:
        waiting=900
    months_dict = {'12': 'Январь', '01': 'Февраль', '02': 'Март', '03': 'Апрель', '04': "Май", '05': "Июнь", "06": "Июль",
              "07": 'Август', '08': 'Сентябрь', '09': "Октябрь", '10': 'Ноябрь', '11': 'Декабрь'}
    url_end=changer(comber.get())
    print(url_end)
    end_date = comber_date.get()
    end_month = get_key(months_dict,str(end_date))
    try:
        service = Service(executable_path='chromedriver.exe')
        options = webdriver.ChromeOptions()
        if checker_graphic_var.get()==1: #Проверка на нажатие кнопки "Выключить графику"
            options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)
    except:
        driver = webdriver.ChromiumEdge()
    driver.implicitly_wait(waiting)
    driver.get(url_end)
    login = login_entry.get()
    password = password_entry.get()
    field_finder=driver.find_element(By.XPATH, "//*[@id='username']")
    field_finder.send_keys(login)
    field_finder=driver.find_element(By.XPATH, "//*[@id='password']")
    field_finder.send_keys(password)
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='signin']/section/div/section[2]/form/button").click()
    time.sleep(9)
    driver.implicitly_wait(30)
    driver.get(url_end)
    parser()

#Объявляются листы с магазинами и месяцами (используются везде)
magazines = ['TOPS', 'Стельки', 'Триколор', 'Джибитсы', 'Discont OFF']
months=['Январь','Февраль','Март','Апрель',"Май","Июнь","Июль",'Август','Сентябрь',"Октябрь",'Ноябрь','Декабрь']
#Ниже инициализируется окно программы
window = tk.Tk()
window.title('Scaler')
window ['bg'] = 'gray10'
login_text = tk.Label(text="Введите логин",background='gray10',foreground='white')
login_entry = tk.Entry()
login_text.pack()
login_entry.pack()
password_text = tk.Label(text="Введите пароль",background='gray10',foreground='white')
password_entry = tk.Entry()
password_text.pack()
password_entry.pack()
date_label = tk.Label(text="Выберите месяц начала парсинга",background='gray10',foreground='white')
date_label.pack()
comber_date = ttk.Combobox(values=months)
comber_date.pack()
magazine_label = tk.Label(text="Выберите магазин",background='gray10',foreground='white')
magazine_label.pack()
comber = ttk.Combobox(values=magazines)
comber.pack()
button = tk.Button(text='Начать выполнение', command=main)
button ['bg'] = 'white'
button.pack()
checker_internet_var=tk.IntVar()
checker_graphic_var=tk.IntVar()
checker_graphic=ttk.Checkbutton(text='Отключить графику',variable=checker_graphic_var)
checker_graphic.pack()
checker_debugger_var=tk.IntVar()
checker_debugging=ttk.Checkbutton(text='Debugging',variable=checker_debugger_var)
checker_debugging.pack()
window.mainloop()

