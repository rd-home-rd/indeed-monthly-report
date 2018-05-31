#coding:utf-8
import lxml.html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys
import nose.tools as nose
import pandas as pd
import os
import shutil
from os import listdir
import datetime
import time
from dateutil.relativedelta import relativedelta
from calendar import monthrange
import xlsxwriter

writer = pd.ExcelWriter('alldata_analytics_history.xlsx', engine='xlsxwriter')
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.preferences.instantApply",True)
profile.set_preference('browser.download.folderList',2) # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', 'C:\\Users\\tai.RD-WOODS\\Desktop\\csv_indeed2')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ' data:text/csv')

profile.set_preference("browser.helperApps.alwaysAsk.force", False);
profile.set_preference('browser.download.manager.useWindow', False);
driver = webdriver.Firefox(firefox_profile=profile)
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
	page=driver.find_elements_by_xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a')
	pagename=lxml.html.fromstring(driver.page_source).xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a/text()')[k].replace('株式会社リクルーティング・デザイン for  ','').replace('株式会社リクルーティング・デザイン for ','').replace('株式会社リクルーティング・デザイン for　','').replace('株式会社リクルーティング・デザイン for','')
	page[k].click()
	wait.until(ec.presence_of_all_elements_located)
	if lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/div[@class='page_indented_content']/div[@id='result_options']/div[@class='result_option'][1]/p")==[]:
		monthlist=[]
		sponsorlist=[]
		clickpercentlist=[]
		clicklist=[]
		sponsorapplypercent=[]
		sponsorapply=[]
		costlist=[]
		clickcost=[]
		costperapply=[]
		date_str='2017/01/01'
		date_formatted = datetime.datetime.strptime(date_str, "%Y/%m/%d")
		for i in range(int(datetime.datetime.today().strftime('%m'))+12):
			driver.get("https://ads.indeed.com/job/ads?endDate="+date_formatted.strftime("%Y-%m")+'-'+str(monthrange(date_formatted.year,date_formatted.month)[1])+"&startDate="+date_formatted.strftime("%Y-%m")+'-01')
			time.sleep(1)
			try:
				if lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][1]/text()")[0] !='0':
					monthlist.append(date_formatted.strftime("%m")+u'月')
					sponsorlist.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][1]/text()")[0])
					clickpercentlist.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][4]/text()")[0].replace('\n','').replace(' ',''))
					clicklist.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][2]/text()")[0])
					sponsorapplypercent.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][5]/text()")[0].replace('-','0').replace('\n','').replace(' ',''))
					sponsorapply.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][3]/text()")[0].replace('-','0'))
					costlist.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][6]/text()")[0].replace('￥','').replace('-','0').replace('\n','').replace(' ',''))
					clickcost.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][7]/text()")[0].replace('￥','').replace('-','0').replace('\n','').replace(' ',''))
					costperapply.append(lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/table[@id='sjc_table']/tbody/tr[@class='footer']/td[@class='right'][8]/text()")[0].replace('￥','').replace('-','0').replace('\n','').replace(' ',''))					
				date_formatted+=relativedelta(months=1)
				if date_formatted.month==datetime.datetime.today().month and date_formatted.year==datetime.datetime.today().year:
					for l in range(i+2):
						driver.back()
					break
			except:
				driver.back()
				driver.back()
				break	
			time.sleep(2)
		
		df=pd.DataFrame(data={'月':monthlist ,'表示回数':sponsorlist,'クリック率':clickpercentlist,'クリック数':clicklist,'応募率':sponsorapplypercent,'応募数':sponsorapply,'合計費用':costlist,'クリック単価':clickcost,'応募単価':costperapply})
		print(pagename)
		print(df)
		df.to_excel(writer,sheet_name=pagename,index=False)
	else:
		driver.back()
driver.close()
writer.save()
