#coding:utf-8
import lxml.html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import nose.tools as nose
import pandas as pd
import os
import shutil
from os import listdir
import time

profile = webdriver.FirefoxProfile()

profile.set_preference('browser.download.folderList',2) # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', 'C:\\Users\\seisaku\\Desktop\\Indeed\buffer')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
driver = webdriver.Firefox(profile)
wait=WebDriverWait(driver,60)
driver.get("https://secure.indeed.com/account/login?service=mob&hl=ja_JP&co=JP&continue=https%3A%2F%2Fjp.indeed.com%2F%3F_ga%3D2.167660222.1622849172.1514262980-517429143.1514262980&tmpl=dmobile")
wait.until(ec.presence_of_all_elements_located)
driver.find_element_by_id('signin_email').clear()
driver.find_element_by_id('signin_email').send_keys("")
driver.find_element_by_id('signin_password').clear()
driver.find_element_by_id('signin_password').send_keys("")
driver.find_element_by_xpath('//form[@id="loginform"]/button[@class="sg-btn sg-btn-primary btn-signin"]').click()
wait.until(ec.presence_of_all_elements_located)
driver.get("https://ads.indeed.com/master_summary")
pagelist=driver.find_elements_by_xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a')
for k in range(len(pagelist)):
	flag=0
	page=driver.find_elements_by_xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a')
	pagename=lxml.html.fromstring(driver.page_source).xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a/text()')[k].replace('株式会社リクルーティング・デザイン for  ','').replace('株式会社リクルーティング・デザイン for ','').replace('株式会社リクルーティング・デザイン for　','').replace('株式会社リクルーティング・デザイン for','')
	page[k].click()
	wait.until(ec.presence_of_all_elements_located)

#	キャンペーンCSVダウンロード
	try:
		driver.find_element_by_xpath('//div[@id="filter_options"]/a[@class="margin-right"]').click()
	except:
		pass

	wait.until(ec.presence_of_all_elements_located)

	os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer")

	for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer"):
		df= pd.read_csv(f)
		df.to_excel(pagename+'.xlsx',sheet_name=pagename,index=None,encoding="CP932")
		shutil.move('C:\\Users\\seisaku\\Desktop\\Indeed\\buffer\\'+pagename+'.xlsx','C:\\Users\\seisaku\\Desktop\\Indeed\\weekly\\')
		os.remove('C:\\Users\\seisaku\\Desktop\\Indeed\\buffer\\'+f)
	driver.back()
	
driver.quit()

os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\weekly")
excel_writer = pd.ExcelWriter('alldata.xlsx')
for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\weekly"): 
	os.rename(f,'1.xlsx')
	df=pd.read_excel('1.xlsx',sheet_name=f.replace('.xlsx','') ,encoding="shift-jis")
	print(str(f))
	df.to_excel(excel_writer,sheet_name=f.replace('.xlsx',''),index=False)
	os.rename('1.xlsx',f)
excel_writer.save()
