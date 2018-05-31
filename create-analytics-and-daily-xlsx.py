#coding:utf-8
import lxml.html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
import nose.tools as nose
import pandas as pd
import os
import datetime
import time
import shutil
from os import listdir
from dateutil.relativedelta import relativedelta
from calendar import monthrange

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.preferences.instantApply",True)
profile.set_preference('browser.download.folderList',2) # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', 'C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ' data:text/csv')
profile.set_preference("browser.helperApps.alwaysAsk.force", False);
profile.set_preference('browser.download.manager.useWindow', False);
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
start=datetime.date.today()-relativedelta(months=1)
for k in range(len(pagelist)):
	page=driver.find_elements_by_xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a')
	pagename=lxml.html.fromstring(driver.page_source).xpath('//table[@id="hjhome"]/tbody/tr[@class="datarow unique"]/td[@class="left first"]/a/text()')[k].replace('株式会社リクルーティング・デザイン for  ','').replace('株式会社リクルーティング・デザイン for ','').replace('株式会社リクルーティング・デザイン for　','').replace('株式会社リクルーティング・デザイン for','')
	page[k].click()
	wait.until(ec.presence_of_all_elements_located)
	if lxml.html.fromstring(driver.page_source).xpath("/html/body/div[@id='page_frame']/div[@id='page_content_wrapper']/div[@id='page_content']/div[@class='page_indented_content']/div[@id='result_options']/div[@class='result_option'][1]/p")==[]:
		driver.get("https://ads.indeed.com/job/ads?startDate="+start.strftime("%Y-%m-01")+"&endDate="+start.strftime("%Y-%m")+'-'+str(monthrange(start.year,start.month)[1]))
	#	日別CSVダウンロード
		try:
			driver.find_element_by_xpath('//div[@id="table_header"]/ul[@class="report_tab"]/li[@class="unselected"]/a').click()
		except:
			pass
		wait.until(ec.presence_of_all_elements_located)
		try:
			driver.find_element_by_xpath('//div[@id="filter_options"]/a[@class="margin-right"]').click()
		except:
			pass
		wait.until(ec.presence_of_all_elements_located)
		os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2")
		for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2"):
			df= pd.read_csv(f)
			df.to_excel(pagename+'.xlsx',sheet_name=pagename,index=False,encoding="CP932")
			shutil.move('C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2\\'+pagename+'.xlsx','C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\daily')
			os.remove('C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2\\'+f)
		
		driver.back()
		driver.get("https://analytics.indeed.com/analytics/jobperf?startDate="+start.strftime("%Y-%m-01")+"&endDate="+start.strftime("%Y-%m")+'-'+str(monthrange(start.year,start.month)[1]))

		time.sleep(5)
		if lxml.html.fromstring(driver.page_source).xpath("//div[@id='body-container']/div[@class='ia-BodyWrapper']/div[@class='ia-BodyMain']/div/div[@class='icl-Grid'][3]/div[@class='icl-Grid-col icl-u-xs-span10']/div/div[@class='ia-PerfReportToolBar']/div[@class='ia-PerfReportToolBarItem'][3]/button[@class='icl-Button--secondary icl-Button--sm ia-PerfReportToolBarItem-dropbtn']")==[]:
			driver.find_element_by_tag_name("body").send_keys(Keys.F5)
			print("F5")
			time.sleep(5)
		wait.until(ec.element_to_be_clickable((By.XPATH, "//div[@id='body-container']/div[@class='ia-BodyWrapper']/div[@class='ia-BodyMain']/div/div[@class='icl-Grid'][3]/div[@class='icl-Grid-col icl-u-xs-span10']/div/div[@class='ia-PerfReportToolBar']/div[@class='ia-PerfReportToolBarItem'][3]/button[@class='icl-Button--secondary icl-Button--sm ia-PerfReportToolBarItem-dropbtn']")))
		webdriver.ActionChains(driver).move_to_element(driver.find_element_by_xpath("//div[@id='body-container']/div[@class='ia-BodyWrapper']/div[@class='ia-BodyMain']/div/div[@class='icl-Grid'][3]/div[@class='icl-Grid-col icl-u-xs-span10']/div/div[@class='ia-PerfReportToolBar']/div[@class='ia-PerfReportToolBarItem'][3]/button[@class='icl-Button--secondary icl-Button--sm ia-PerfReportToolBarItem-dropbtn']")).perform()
		wait.until(ec.element_to_be_clickable((By.XPATH,"//div[@id='body-container']/div[@class='ia-BodyWrapper']/div[@class='ia-BodyMain']/div/div[@class='icl-Grid'][3]/div[@class='icl-Grid-col icl-u-xs-span10']/div/div[@class='ia-PerfReportToolBar']/div[@class='ia-PerfReportToolBarItem'][3]/div[@class='ia-PerfReportToolBarItem-content']/a[2]")))
		webdriver.ActionChains(driver).click(driver.find_element_by_xpath("//div[@id='body-container']/div[@class='ia-BodyWrapper']/div[@class='ia-BodyMain']/div/div[@class='icl-Grid'][3]/div[@class='icl-Grid-col icl-u-xs-span10']/div/div[@class='ia-PerfReportToolBar']/div[@class='ia-PerfReportToolBarItem'][3]/div[@class='ia-PerfReportToolBarItem-content']/a[2]")).perform()
		wait.until(ec.presence_of_all_elements_located)
		time.sleep(2)
		os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2")
		for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2"):
			os.rename(f,'1.csv')
			df=pd.read_csv('1.csv')
			os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\analytics")
			df.to_excel(pagename+'.xlsx',sheet_name=pagename,index=None,encoding="CP932")
			os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2")
			os.remove('C:\\Users\\seisaku\\Desktop\\Indeed\\buffer_2\\1.csv')
		driver.back()
		driver.back()
	driver.back()
driver.quit()
os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\daily")
excel_writer = pd.ExcelWriter('alldata_daily.xlsx')
for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\daily"): 
	os.rename(f,'1.xlsx')
	df=pd.read_excel('1.xlsx',sheet_name=f.replace('.xlsx',''), encoding="shift-jis")
	print(str(f))
	df.to_excel(excel_writer,sheet_name=f.replace('.xlsx',''),index=False)
	os.rename('1.xlsx',f)
excel_writer.save()
os.chdir("C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\analytics")
excel_writer = pd.ExcelWriter('alldata_analytics.xlsx')
for f in listdir("C:\\Users\\seisaku\\Desktop\\Indeed\\monthly\\analytics"): 
	os.rename(f,'1.xlsx')
	df=pd.read_excel('1.xlsx',sheet_name=f.replace('.xlsx',''), encoding="shift-jis")
	print(str(f))
	df.to_excel(excel_writer,sheet_name=f.replace('.xlsx',''),index=False)
	os.rename('1.xlsx',f)
excel_writer.save()
