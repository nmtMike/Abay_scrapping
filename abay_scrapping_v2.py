from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from os.path import exists
import os
import pandas as pd
import time
import easygui
import sys

abay_routes = {'SGN':'Tp_Ho_Chi_Minh', 'HAN':'Ha_Noi', 'DAD':'Da_Nang', 'UIH':'Quy_Nhon', 'PQC':'Phu_Quoc'}

# if scrapping_info not exists or null, tell user to input information
if not exists('scrapping_info.xlsx'):
    easygui.msgbox("Please input the scrapping_info", title="Input alert")
    pd.DataFrame(columns=['from_date', 'to_date', 'route']).to_excel('scrapping_info.xlsx', index=False)
    sys.exit()
if pd.read_excel('scrapping_info.xlsx').shape[0] == 0:
    easygui.msgbox("Please input the scrapping_info", title="Input alert")
    sys.exit()


# define a function to create a list of day_sector to scrape
# also make a needed directory
def route_pricing(from_date:str, to_date:str, route:str):
    # create a folder in current directory
    parent_dir = os.getcwd() + '\\' + 'abay_scrapping'
    try:
        os.mkdir(parent_dir)
    except FileExistsError:
        pass

    scrapping_dates = pd.date_range(start=from_date, end=to_date)

    # create string name of folder container
    dep1 = abay_routes[route[:3]]
    arr1 = abay_routes[route[4:7]]
    dep2 = arr1
    arr2 = dep1
    sector1 = dep1 + '-' + arr1
    sector2 = dep2 + '-' + arr2
    
    # create directory
    sub_dir1 = route[:7]
    sub_dir2 = route[4:]
    dir_1 = os.path.join(parent_dir, sub_dir1)
    dir_2 = os.path.join(parent_dir, sub_dir2)
    
    try:
        os.mkdir(dir_1)
    except FileExistsError:
        pass
    
    try:
        os.mkdir(dir_2)
    except FileExistsError:
        pass
    
    # create a dataframe
    to_scrape1 = pd.DataFrame({'dates':scrapping_dates, 'sector':sector1})
    to_scrape1['dir'] = dir_1
    to_scrape2 = pd.DataFrame({'dates':scrapping_dates, 'sector':sector2})
    to_scrape2['dir'] = dir_2
    to_scrape = pd.concat([to_scrape1, to_scrape2])

    # flattern and return into a list
    return to_scrape

# define web scrapping function
def abay_scrapping(ngay, thang, nam, ten, location):
    website = f'https://www.abay.vn/Ve-May-Bay-Gia-Re-{ten}-{ngay}Thang{thang}'
    driver.get(website)
    time.sleep(1) # make sure the driver has time to load contain
    
    flights_table = driver.find_element(By.XPATH, '//table[@class="f-result"]')
    flight_rows = flights_table.find_elements(By.XPATH, '//tr[@class="i-result"]')

    f_name = []
    f_time = []
    f_baggage_meal = []
    f_price = []
    final_baggage_meal = []

    for row in flight_rows:
        f_name.append(row.find_element(By.XPATH, './td[2]').text)
        f_time.append(row.find_element(By.XPATH, './td[3]').text)
        f_baggage_meal.append(row.find_element(By.XPATH, './td[4]').find_elements(By.TAG_NAME, 'img'))
        f_price.append(row.find_element(By.XPATH, './td[5]').text)

    # create a list contain baggage and meal information
    for bag_meal in f_baggage_meal:
        tmp_str = ''
        for element in bag_meal:
            tmp_str += '-' + element.get_attribute('src').split('/')[-1]
        final_baggage_meal.append(tmp_str)

    f_date = [f'{nam}-{thang}-{ngay}'] * len(f_name)
    df = pd.DataFrame({'name': f_name, 'time': f_time, 'baggage_meal': final_baggage_meal, 'price': f_price, 'date' : f_date})
    df['price'] = df['price'].str[:-1] # delete the "d" currency value

    df['bag'] = df['baggage_meal'].str.contains('hanhly')*1
    df['meal'] = df['baggage_meal'].str.contains('suatan')*1
    column_order = ['name', 'time', 'bag', 'meal', 'price', 'date']
    df = df.reindex(columns=column_order)
    
    if ngay < 10:
        ngay = '0' + str(ngay)
    if thang < 10:
        thang = '0' + str(thang)

    csv_path = str(location) + r'\price' + str(ten) + str(nam) + str(thang) + str(ngay) + '.csv'
    df.to_csv(csv_path, index=False)


# define a function to apply on a dataframe
def route_pricing_apply(row):
    to_scrape_all.append(route_pricing(from_date=row['from_date'], to_date=row['to_date'], route=row['route']))


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
scrapping_info = pd.read_excel('scrapping_info.xlsx')

to_scrape_all = []
scrapping_info.apply(route_pricing_apply, axis=1)

to_scrape_all = pd.concat(to_scrape_all).drop_duplicates().values.tolist()

err_list = []
for scrape in to_scrape_all:
    i = 0
    while True:
        i += 1
        if i == 5:
            err_list.append(scrape)
            break
        try:
            abay_scrapping(scrape[0].day, scrape[0].month, scrape[0].year, scrape[1], scrape[2])
        except:
            continue
        break
    
driver.quit()

pd.DataFrame(err_list, columns=['date', 'sector', 'dir']).to_excel('error_list.xlsx', index=False)
sys.exit()