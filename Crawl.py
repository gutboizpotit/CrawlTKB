from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import PySimpleGUI as psg
import requests
from pandas import DataFrame
from getpass import getpass

# open_web

def open_web(user, password, field):
    driver = webdriver.Chrome('chromedriver.exe')
    url = "https://qldt.ptit.edu.vn/default.aspx?page=dangnhap"
    driver.get(url)
    sleep(3)
    return login(driver=driver, user=user, passw=password, field=field)

def login(driver, user, passw, field):
    username = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ctl00_txtTaiKhoa')
    username.send_keys(user)
    password = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ctl00_txtMatKhau')
    password.send_keys(passw)
    login = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ctl00_btnDangNhap')
    login.click()
    sleep(3)
    if field == 'TKB': return CrawlTKB(driver=driver)
    else: return CrawlDiem(driver=driver)

def CrawlTKB(driver):
    try:
        tkb_button = driver.find_element(By.ID, 'ctl00_menu_lblThoiKhoaBieu')
        tkb_button.click()
        sleep(3)
    except:
        return psg.popup("Sai thông tin đăng nhập")
    tuan = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ctl00_ddlTuan')
    Selecttuan = Select(driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ctl00_ddlTuan'))
    Selecttuan.select_by_visible_text(values["Tuan"])
    sleep(3)
    date = values["Tuan"]
    ftkb = open('TKB.csv','a', encoding='utf-8')
    ftkb.write( str(date) + '\n' )
    ftkb.write( "," + "Thứ 2" + "," + "Thư 3" + "," + "Thứ 4" + "," + "Thứ 5" + "," + "Thứ 6" + "," + "Thứ 7" + "," + "Chủ Nhật\n")
    sohang = len(driver.find_elements(By.XPATH, "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/div/div[3]/table[1]/tbody/tr"))
    socot = len(driver.find_elements(By.XPATH, "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/div/div[3]/table[1]/tbody/tr[1]/td"))
    for hang in range(1,sohang):
        tiet = driver.find_element(By.XPATH,"/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/div/div[3]/table[1]/tbody/tr[" + str(hang) + "]/td[1]").text
        ftkb.write( tiet + "," )
        for cot in range(2,socot):
            xp = "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/div/div[3]/table[1]/tbody/tr[" + str(hang) + "]/td[" + str(cot) + "]/table"
            if( len(driver.find_elements(By.XPATH , xp)) > 0 ):
                mon = driver.find_element(By.XPATH,"/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/div/div[3]/table[1]/tbody/tr[" + str(hang) + "]/td[" + str(cot) + "]").text
                lstmon = mon.split("\n") 
                for s in lstmon:
                    ftkb.write( s + ' ' )
            ftkb.write(",")
        ftkb.write("\n")
    psg.popup("Đẫ xong")
                
def CrawlDiem(driver):
    try:
        diem_button = driver.find_element(By.ID, 'ctl00_menu_lblXemDiem')
        diem_button.click()
        sleep(3)
    except:
        return psg.popup("Sai thông tin đăng nhập")
    fd = open('Diem.csv', 'a', encoding='utf-8')
    hocky = values["Hocky"]
    if( hocky == "Học kỳ 1 - Năm học 2020-2021" ):
        somon = 6
        st = 3
    elif( hocky == "Học kỳ 2 - Năm học 2020-2021" ): 
        somon = 8
        st = 14
    elif( hocky == "Học kỳ 1 - Năm học 2021-2022" ):
        somon = 7
        st = 27
    else:
        somon = 7
        st = 39
    sohang = len(driver.find_elements(By.XPATH, "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/table/tbody/tr[3]/td/div/table/tbody/tr"))
    socot = len(driver.find_elements(By.XPATH, "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/table/tbody/tr[3]/td/div/table/tbody/tr[1]/td"))
    fd.write(hocky + "\n")
    for cot in range(1,socot):
        xp = "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/table/tbody/tr[3]/td/div/table/tbody/tr[1]/td[" + str(cot) + "]"
        dulieu = driver.find_element(By.XPATH , xp).text
        fd.write(dulieu + ",")
    fd.write("\n")
    for hang in range(st,st+somon):
        for cot in range(1,socot):
            xp = "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/table/tbody/tr[3]/td/div/table/tbody/tr[" + str(hang) + "]/td[" + str(cot) + "]"
            fd.write(driver.find_element(By.XPATH , xp).text)
            fd.write(",")
        fd.write("\n")
    for hang in range(st+somon,st+somon+4):
        xp = "/html/body/form/div[3]/div/table/tbody/tr[2]/td/div[3]/table/tbody/tr[3]/td/div/table/tbody/tr[" + str(hang) + "]/td[1]"
        fd.write(driver.find_element(By.XPATH , xp).text)
        fd.write("\n")
    psg.popup("Đã xong")


layout = [ [psg.Text("THÔNG TIN ĐĂNG NHÂP QLDT" , background_color="Blue", text_color="Black", justification="center")],
        [psg.Text("Mã sinh viên",size=15), psg.Input(key="User", size=20)],
        [psg.Text("Mật khẩu", size=15), psg.Input(key="Pass", size=20, password_char='*')],
        [psg.Text("Tuần học", size=15), psg.Combo(['Tuần 01 [Từ 15/08/2022 -- Đến 21/08/2022]', 
        'Tuần 02 [Từ 22/08/2022 -- Đến 28/08/2022]', 'Tuần 03 [Từ 29/08/2022 -- Đến 04/09/2022]', 
        'Tuần 04 [Từ 05/09/2022 -- Đến 11/09/2022]', 'Tuần 05 [Từ 12/09/2022 -- Đến 18/09/2022]', 
        'Tuần 06 [Từ 19/09/2022 -- Đến 25/09/2022]', 'Tuần 07 [Từ 26/09/2022 -- Đến 02/10/2022]',
        'Tuần 08 [Từ 03/10/2022 -- Đến 09/10/2022]','Tuần 09 [Từ 10/10/2022 -- Đến 16/10/2022]',
        'Tuần 10 [Từ 17/10/2022 -- Đến 23/10/2022]','Tuần 11 [Từ 24/10/2022 -- Đến 30/10/2022]',
        'Tuần 12 [Từ 31/10/2022 -- Đến 06/11/2022]','Tuần 13 [Từ 07/11/2022 -- Đến 13/11/2022]',
        'Tuần 14 [Từ 14/11/2022 -- Đến 20/11/2022]','Tuần 15 [Từ 21/11/2022 -- Đến 27/11/2022]',
        'Tuần 16 [Từ 28/11/2022 -- Đến 04/12/2022]'], key = "Tuan")],
        [psg.Text("Học kỳ", size=15), psg.Combo(['Học kỳ 1 - Năm học 2020-2021', 'Học kỳ 2 - Năm học 2020-2021', 'Học kỳ 1 - Năm học 2021-2022', 'Học kỳ 2 - Năm học 2021-2022'] , key = "Hocky" )],
        [psg.Button("Xuất TKB", button_color="Black", size=20 ) , psg.Button("Xuất Điểm", button_color="Black",size=20)]
        ]
window = psg.Window("MÀN HÌNH ĐĂNG NHẬP" , layout)

while True :
    event , values = window.read()
    if event == "Xuất TKB":
        open_web(user=values["User"], password=values["Pass"],field='TKB')
    if event == "Xuất Điểm":
        open_web(user=values["User"], password=values["Pass"],field='Diem')
    if event == psg.WIN_CLOSED:
        break