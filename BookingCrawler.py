# -*- coding: utf-8 -*-
"""
Created on Dec 2022
@author: Y.C
"""
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import requests #Method:get or post
import time
import os
import shutil
import pandas as pd 
import openpyxl
def nextpage(): #換頁
    while True:
        hotel_link = driver.find_elements(By.CLASS_NAME ,"e13098a59f" )
        for i in hotel_link :
            hotel_list.append(i.get_attribute("href"))
        time.sleep(5)
        next_page_button = driver.find_element(By.XPATH,'//*[@id="search_results_table"]/div[2]/div/div/div/div[5]/div[2]/nav/div/div[3]/button')
        #最後一頁的next_page_button的class跟前面不同
        if next_page_button.get_attribute("class") =='fc63351294 a822bdf511 e3c025e003 fa565176a8 f7db01295e f9d6150b8e':
            break
        else:
            next_page_button.click()

def price_adjust(a): #原始價格,目前價格
    b = a.split(" ")
    if len(b)>3:
        number1=""
        for i in b[2]:
            if i.isdigit():
                number1+=i
        number2=""
        for j in b[5]:
            if j.isdigit():
                number2+=j
        return int(number1),int(number2)
    else:
        number3=""
        for k in b[2]:
            if k.isdigit():
                number3+=k
        return number3,number3

def distance_adjust(a): #距離
    distance=''
    if '捷運台北101' in a:
       a = a[a.index('站')::]
    if a[-2:]=='公里':
        for k in a:
            if k.isdigit():
                distance+=k
        distance = int(distance) * 100
    else:   
        for i in a:
            if i.isdigit():
                distance+=i
    return int(distance)


options = Options()
options.add_argument("--disable-notifications")

driver = webdriver.Chrome("D:\chromedriver\chromedriver.exe")
driver.implicitly_wait(3)
#booking網站
url = "https://www.booking.com/searchresults.zh-tw.html?ss=%E5%9C%8B%E7%88%B6%E7%B4%80%E5%BF%B5%E9%A4%A8&ssne=%E5%9C%8B%E7%88%B6%E7%B4%80%E5%BF%B5%E9%A4%A8&ssne_untouched=%E5%9C%8B%E7%88%B6%E7%B4%80%E5%BF%B5%E9%A4%A8&efdco=1&label=yho748jc-1DCAEoggI46AdIMFgDaOcBiAEBmAEwuAEXyAEM2AED6AEB-AECiAIBqAIDuALpn4-dBsACAdICJDQ5NmZiZGU4LTRjNWYtNGJhMy1hMjIyLTAxMGRlYmIyYTkxZtgCBOACAQ&sid=874c44ecd046bd3cf82995f6c85a9c6d&aid=397646&lang=zh-tw&sb=1&src_elem=sb&src=searchresults&dest_id=233880&dest_type=landmark&checkin=2023-02-11&checkout=2023-02-12&group_adults=2&no_rooms=1&group_children=0&sb_travel_purpose=leisure"
driver.get(url) #讓selenuim前往的網址

start_time=time.ctime()

'''建立hotel_list'''
hotel_list = []
#換頁button
nextpage()
#檢查是否有重複項
hotel_list_nondup=[]
for i in hotel_list :
    if i not in hotel_list_nondup :
        hotel_list_nondup.append(i)

count = len(hotel_list_nondup) 
hotel_list = hotel_list_nondup ###不重複旅館list
print("總共：",count,"間旅館")

"""寫進worksheet(openpxl)"""
wb = openpyxl.Workbook()
sheet1 = wb.active
sheet1.append(["旅館名稱","人數","原始價格","目前價格"])
sheet2 = wb.create_sheet('SHEET2') 
sheet2.append(["旅館名稱","總分","員工素質","設施","清潔程度","舒適程度","性價比","住宿地點","免費 WiFi","最近大眾距離","星級"])
name ='Booking'+time.ctime()[4:7]+time.ctime()[8:10]+'.xlsx'  #檔名

#url = 'https://www.booking.com/hotel/tw/hotel-june.zh-tw.html?aid=397646&label=yho748jc-1FCAEoggI46AdIMFgDaOcBiAEBmAEwuAEXyAEM2AEB6AEB-AECiAIBqAIDuALpn4-dBsACAdICJDQ5NmZiZGU4LTRjNWYtNGJhMy1hMjIyLTAxMGRlYmIyYTkxZtgCBeACAQ&sid=6bf999ccbb49e6108244b347cd5055a2&all_sr_blocks=62689504_356590570_0_0_0;checkin=2023-02-10;checkout=2023-02-12;dest_id=233880;dest_type=landmark;dist=0;group_adults=2;group_children=0;hapos=218;highlighted_blocks=62689504_356590570_0_0_0;hpos=18;matching_block_id=62689504_356590570_0_0_0;no_rooms=1;req_adults=2;req_children=0;room1=A%2CA;sb_price_type=total;sr_order=popularity;sr_pri_blocks=62689504_356590570_0_0_0__522500;srepoch=1672296331;srpvid=cb802f6a236a0042;type=total;ucfs=1&#hotelTmpl'
hotelnumber=1
for url in hotel_list:
    print("第",hotelnumber,'/',count,"間")
    hotelnumber+=1
    driver.get(url) 
    time.sleep(3)

    """旅館名稱(SHEET1,SHEET2)"""
    A = driver.find_element(By.XPATH,'//*[@id="hp_hotel_name"]/div/div/h2')
    """房間容納人數、原價、目前價格(SHEET1)"""
    roomtypelist=[]
    lista=[]
    roomtable = driver.find_element(By.ID,'hprt-table')
    
    lista = roomtable.find_elements(By.CLASS_NAME,'bui-u-sr-only')
    peoplelist=lista[0:len(lista):2]
    pricelist=lista[1:len(lista)+1:2]
    
    peoplelistA= []
    pricelist_Origin = []
    pricelist_Discount = []
    
    for i,j in zip(peoplelist,pricelist):
        #print(i.text[-1])
        peoplelistA.append(i.text[-1])
        peopleA = i.text[-1]
        #print(j.text)
        a_1,a_2 = price_adjust(j.text)
        pricelist_Origin.append(a_1)
        pricelist_Discount.append(a_2)
        price_Origin = a_1
        price_Discount = a_2 
        
        sheet1.append([A.text,peopleA,a_1,a_2])
        
    
    """項目評價、最近大眾交通工具、星級、評論(SHEET2)"""
    total_list=[] #本頁要加入sheet的list  list.extend(score_class,comment_list)
    score_class =[]  #名稱1、總評1、評分7、交通1、星級1
    ####評價####
    #名稱
    score_class.append(A.text)
    #總評價
    score_total = driver.find_element(By.CLASS_NAME , "b5cd09854e" ).text
    print("總分:",score_total)
    score_class.append(score_total)
    #各項評價
    score_name_list = ["員工素質","設施","清潔程度","舒適程度","性價比","住宿地點","免費 WiFi"]
    score1 = driver.find_elements(By.CLASS_NAME,"ee746850b6")
    for i in range(len(score1)) :
        print(score1[i].text,end = '/')
        ###print([i for i in score1[i].text if (i.isdigit() or i=='.') ])
        if score1[i].text in score_name_list :
            score_class.append(score1[i+1].text)
    if len(score_class)<9:
        score_class.append('')
    ####交通####
    #右方欄位有相同class: 附近有什麼、鄰近餐廳、景點、自然美景、鄰近交通...(鄰近交通在第五個欄位)
    traffic = driver.find_elements(By.CLASS_NAME , "d31796cb42" )
    time.sleep(1)
    try:
        traffic_5 = traffic[4]  #找第五個欄位
        traffic_first = traffic_5.find_elements(By.CLASS_NAME , "fc63351294" ) 
        distance = distance_adjust(traffic_first[0].text) #第一項為距離旅館最近的大眾運輸
    except IndexError:
        try:
            driver.refresh()
            time.sleep(3)
            traffic_5 = traffic[4]
            traffic_first = traffic_5.find_elements(By.CLASS_NAME , "fc63351294" )
            distance = distance_adjust(traffic_first[0].text)
        except IndexError:
            distance = ''
    score_class.append(distance)
    print("最近大眾:",distance)
    
    ####飯店星級####
    #星星圖案個數代表星級 ; class = "b6dc9a9e69 adc357e4f1 fe621d6382"
    #找不到代表 : 無星級飯店
    try :
        star = driver.find_elements(By.CLASS_NAME , "adc357e4f1" )
        star_count = len(star)
    except NoSuchElementException:
        star_count = 0
    score_class.append(star_count)
    print("星級:",star_count)
    
    
    comment_list =[] #各旅館經常出現之評論關鍵字
    ####評論####
    #class="b0fee5870b" 
    #class="bui-input-checkbutton__item
    try:
        comment_buttom = driver.find_element(By.CLASS_NAME,'b0fee5870b')
        comment_buttom.click()
        time.sleep(3)
        #comment_div = driver.find_element(By.CLASS_NAME,'bui-group bui-group--inline')
        comment_text = driver.find_elements(By.CLASS_NAME,'bui-input-checkbutton__item')
        for i in comment_text:
            print(i.text)
            if i.text != "":
                comment_list.append(i.text)
        comment_list = list(set(comment_list)) #避免重複
    except NoSuchElementException:
        print('nosuchelement')
        pass
    
    #本次sheet2新增內容
    total_list = score_class + comment_list 
    print(total_list)
    sheet2.append(total_list)


wb.save(name)

finish_time=time.ctime()

print("Finish")
print(start_time,'開始')
print(finish_time,'結束')
