from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import  pandas as pd
import time


def can_convert_to_int(value):
    try:
        int(value)
        return False
    except ValueError:
        return True

config = pd.read_excel("config.xlsx")
urllist = config['url']
url = urllist[0]

namelist = config['name']
name = namelist[0]
passwordlist = config['password']
password = passwordlist[0]


driver = webdriver.Chrome()
driver.get(url)
WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div[1]/div/input')))
au = driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div[1]/div/input')
au.send_keys(str(name))
driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[2]/div[2]/div[2]/div/input').send_keys(password)
driver.find_element(By.XPATH,'//*[@id="submit"]').click()


WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div[1]/div[1]/div/div[5]/ul/div/div/div[1]/button')))

driver.find_element(By.XPATH,'//*[@id="app"]/div[1]/div[1]/div/div[5]/ul/div/div/div[1]/button').click()


WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,'//*[@id="editor-b"]/div[3]/div[1]/div/span[2]/div')))
time.sleep(2)
se = driver.find_element(By.XPATH,'//*[@id="editor-b"]/div[3]/div[1]/div/span[2]/div')
se.click()

WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[1]/div[1]/ul/li[8]')))


el = driver.find_element(By.XPATH,'/html/body/div[2]/div[1]/div[1]/ul/li[7]')
el.click()

WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,'//*[@id="editor-b"]/div[4]/div[1]/div[3]/table/tbody/tr[201]/td[4]/div/div[1]')))

ds = driver.find_elements(By.XPATH,'//*[@id="wordsTag"]')

df = pd.read_excel("全部可编辑行excel表格.xlsx")

for idx,row in df.iterrows():
    
    行数 = row['行数']
    原文 = row['原文']
    译文 = str(row['译文'])
    print(行数,原文,译文)
    bo = '//*[@id="editor-b"]/div[4]/div[1]/div[3]/table/tbody/tr[' + str(行数) + ']/td[8]/div/div/div/button'
    if can_convert_to_int(行数):
        continue
    if len(译文)==0:
        continue
    istext =ds[int(行数)-1].get_attribute('contenteditable')
    print(istext)
    if istext == 'false':
        continue
    print('******',行数,原文,译文)
    ds[int(行数)-1].clear()
    if 译文 == 'nan':
        continue
    ds[int(行数)-1].send_keys(译文)
    WebDriverWait(driver,1000).until(EC.presence_of_element_located((By.XPATH,bo)))
    driver.find_element(By.XPATH,bo).click()


