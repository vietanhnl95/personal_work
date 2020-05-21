import os
import sys

from selenium import webdriver
from bs4 import BeautifulSoup

import pandas as pd
import time


#get msdn list from excel file
list_msdn = pd.read_excel('thong_tin_doanh_nghiep.xlsx', sheet_name='Sheet1')
list_msdn = list(list_msdn['ma_dn'].apply(lambda x: x.replace('a', '')))

url = 'https://dichvuthongtin.dkkd.gov.vn/inf/Forms/Products/ProductCatalog.aspx?h=d2d4'
browser = webdriver.Chrome(os.getcwd() + '/chromedriver')
result_list = []

for msdn in list_msdn:
    print('running msdn: ', msdn)
    browser.get(url)

    #click on "vai tro ca nhan" segment
    vai_tro_ca_nhan = browser.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[2]/div[1]/div[2]/div/div[1]/div[4]/div[2]/div/ul/li[3]/a')
    vai_tro_ca_nhan.click()
    time.sleep(1)

    #find the input field for msdn
    input_table = browser.find_elements_by_tag_name('table')[0]
    msdn_tr = input_table.find_elements_by_tag_name('tr')[3]
    msdn_input = msdn_tr.find_element_by_tag_name('input')
    msdn_input.send_keys(msdn)

    #find "tim kiem" button
    search_btn = browser.find_elements_by_id('ctl00_C_UC_PERS_LIST1_BtnFilter')
    search_btn[0].click()
    time.sleep(1)

    #make soup with bs4
    soup = BeautifulSoup(browser.page_source)

    #find all results row in result
    trs = soup.find_all('table')[1].tbody.find_all('tr')

    #loop through each result row and extract information
    for tr in trs[1:]:
        result_dict = {}
        tds = tr.find_all('td')
        result_dict['ma_so_doanh_nghiep'] = msdn
        result_dict['giay_chung_thuc_ca_nhan'] = tds[1].a.string
        result_dict['so_giay_chung_thuc_ca_nhan'] = tds[2].a.string
        result_dict['ho_ten'] = tds[3].a.string
        result_dict['gioi_tinh'] = tds[4].a.string
        result_dict['ngay_sinh'] = tds[5].a.string
        result_list.append(result_dict)

# cast dictionary to DataFrame for easier view
df = pd.DataFrame(result_list)
print(df)

# export to excel
df.to_excel('result_file.xlsx')
