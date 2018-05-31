#coding:utf-8
import xlsxwriter
import pandas

from itertools import product
import types
from dateutil.relativedelta import relativedelta
from calendar import monthrange
from datetime import date
import datetime
import xlrd
xls = xlrd.open_workbook(r'alldata_analytics_history.xlsx', on_demand=True)
sheets=xls.sheet_names()
for x in sheets:
	dataframe1=pandas.read_excel('alldata_analytics.xlsx',sheet_name=x,encoding="shift-jis")
	if dataframe1.dtypes.index[5]!='会社名':
		new_col=[]
		for i in range(len(dataframe1.values[:,0])):
			new_col.append(0)
		dataframe1.insert(loc=5,column='会社名',value=new_col)
	for i in range(len(dataframe1.values[:,6])):
		if type(dataframe1.values[i,6])==str:
			dataframe1.at[i,'スポンサー表示回数']=int(dataframe1.values[i,6].replace(',',''))
	for i in range(len(dataframe1.values[:,11])):
		if type(dataframe1.values[i,11])==str:
			dataframe1.at[i,'オーガニック表示回数']=int(dataframe1.values[i,11].replace(',',''))
	for i in range(len(dataframe1.values[:,14])):
		if type(dataframe1.values[i,14])==str:
			dataframe1.at[i,'PCのスポンサー表示回数']=int(dataframe1.values[i,14].replace(',',''))
	for i in range(len(dataframe1.values[:,17])):
		if type(dataframe1.values[i,17])==str:
			dataframe1.at[i,'モバイルのスポンサー表示回数']=int(dataframe1.values[i,17].replace(',',''))
	dataframe2=pandas.read_excel('alldata_daily.xlsx',sheet_name=x,encoding="shift-jis")
	cate=len(dataframe2.values[:,0])
	try:
		start=datetime.datetime.strptime(dataframe2.values[0,0], "%Y-%m-%d").date()
		end=datetime.datetime.strptime(dataframe2.values[cate-1,0], "%Y-%m-%d").date()
	except:
		start=datetime.date.today()
		end=datetime.date.today()
	for i in range(len(dataframe2.values[:,6])):
		if type(dataframe2.values[i,5])==str:
			dataframe2.at[i,'Cost (JPY)']=int(dataframe2.values[i,5].replace(',',''))
	print(start)
	if start!=datetime.date.today():

		print(x)
		writer = pandas.ExcelWriter(x+'様'+(datetime.datetime.today()-relativedelta(months=1)).strftime("%Y%m")+'月次.xlsx', engine='xlsxwriter')
		workbook=writer.book
		worksheet = workbook.add_worksheet('職種+勤務地+詳細')
		
		

		dataframe1.to_excel(writer,sheet_name='data_アナリティクス',index=False)
		
		dataframe2.to_excel(writer,sheet_name='data_日別',index=False)
		dataframe3=pandas.read_excel('alldata_analytics_history.xlsx',sheet_name=x,encoding='shift-jis')
		dataframe3.to_excel(writer,sheet_name='data_過去',index=False)
		worksheet.set_paper(9)
		worksheet.set_margins(0.4,0.4,0.4,0.4)
		worksheet.set_header('', {'margin': 0})
		worksheet.set_footer('', {'margin': 0})
		worksheet.center_horizontally()
		worksheet.center_vertically()
		worksheet.set_print_scale(56)
		mon=(datetime.datetime.today()-relativedelta(months=1)).strftime("%Y")+u'年'+(datetime.datetime.today()-relativedelta(months=1)).strftime("%m")+u'月'
		header = workbook.add_format({'bold': True, 'font_color': 'white','bg_color':'#2E9AFE','font_size':20,'font_name':'Meiryo UI','valign':'vcenter'})
		
#	title start
		worksheet.merge_range('A1:O1','  '+x+'様　Indeed運用レポート',header)
		worksheet.set_default_row(22.5)
		worksheet.set_row(0,49.5)

		worksheet.set_column('A:O',10)
		hheader=workbook.add_format({'bold': True, 'font_color': 'white','bg_color':'#A9D0F5','font_size':14,'font_name':'Meiryo UI','align':'right'})
		worksheet.merge_range('A2:O2','=TEXT(TODAY(), "yyyy.mm.dd")&"　株式会社リクルーティング・デザイン"',hheader)
		Afour=workbook.add_format({'bold': True, 'font_color': '#0431B4','font_size':18,'font_name':'Meiryo UI'})
		worksheet.write('A4','■'+mon+'　月次レポート',Afour)
		Ofour=workbook.add_format({'font_size':11,'font_name':'Meiryo UI','valign':'vcenter','align':'right'})
		worksheet.write('O4','期間：'+(datetime.datetime.today()-relativedelta(months=1)).strftime("%Y/%m/")+'01'+'~'+(datetime.datetime.today()-relativedelta(months=1)).strftime("%Y/%m/")+str(monthrange((datetime.datetime.today()-relativedelta(months=1)).year,((datetime.datetime.today()-relativedelta(months=1)).month))[1]),Ofour)
		five=workbook.add_format({'font_size':12,'font_name':'Meiryo UI','align':'center','valign':'bottom'})
		worksheet.merge_range('A5:C5','表示回数',five)
		worksheet.merge_range('D5:F5','スポンサークリック率',five)
		worksheet.merge_range('G5:I5','スポンサークリック数',five)
		worksheet.merge_range('J5:L5','応募数',five)
		worksheet.merge_range('M5:O5','合計費用',five)
		bignum=workbook.add_format({'bold':True,'font_name':'Meiryo UI','font_size':22,'valign':'bottom','align':'center','num_format':0x03})
		bignum2=workbook.add_format({'bold':True,'font_name':'Meiryo UI','font_size':22,'valign':'bottom','align':'center','num_format':0x0a})
		
#	title end
#	center cells start
		twentyone=workbook.add_format({'bold': True,'font_name':'Meiryo UI','font_size':14,'font_name':'Meiryo UI'})
		worksheet.write('A21','■広告別実績詳細',twentyone)
		twentytwo1=workbook.add_format({'bold': True,'font_name':'Meiryo UI','font_color': 'white','bg_color':'#2E9AFE','font_size':11,'font_name':'Meiryo UI','border':1,'bottom':6})
		worksheet.merge_range('A22:C22','職種名',twentytwo1)
		worksheet.merge_range('D22:E22','勤務地',twentytwo1)
		worksheet.merge_range('F22:G22','詳細',twentytwo1)
		twentytwo2=workbook.add_format({'bold': True,'font_name':'Meiryo UI', 'font_color': 'white','bg_color':'#2E9AFE','font_size':11,'font_name':'Meiryo UI','align':'center','border':1,'bottom':6,'valign':'vcenter'})
		worksheet.write('H22','表示回数',twentytwo2)
		worksheet.write('I22','クリック率',twentytwo2)
		worksheet.write('J22','クリック数',twentytwo2)
		worksheet.write('K22','応募率',twentytwo2)
		worksheet.write('L22','応募数',twentytwo2)
		worksheet.write('M22','合計費用',twentytwo2)
		worksheet.write('N22','クリック単価',twentytwo2)
		worksheet.write('O22','応募単価',twentytwo2)
		chart=workbook.add_format({'right':1,'font_name':'Meiryo UI','left':1,'top':3,'bottom':3,'font_size':11,'num_format':0x03})
		chart2=workbook.add_format({'right':1,'font_name':'Meiryo UI','left':1,'top':3,'bottom':3,'font_size':11,'num_format':0x0a})
		chartbgyellow=workbook.add_format({'right':1,'font_name':'Meiryo UI','left':1,'top':3,'bottom':3,'font_size':11,'num_format':0x03,'bg_color':'#F3F781'})
		chart2bgyellow=workbook.add_format({'right':1,'font_name':'Meiryo UI','left':1,'top':3,'bottom':3,'font_size':11,'num_format':0x0a,'bg_color':'#F3F781'})
		if len(dataframe1.values[:,5])>20:
			a=len(dataframe1.values[:,5])+2
		else:
			a=22



		for i in range(2,a):
			worksheet.merge_range('A'+str(21+i)+':C'+str(21+i),'=IF(data_アナリティクス!$A'+str(i)+'="","",data_アナリティクス!$A'+str(i)+')',chart)
			worksheet.merge_range('D'+str(21+i)+':E'+str(21+i),'=IF(data_アナリティクス!$B'+str(i)+'="","",data_アナリティクス!$B'+str(i)+')',chart)
			worksheet.merge_range('F'+str(21+i)+':G'+str(21+i),'=IF(A23="","",IF(data_アナリティクス!$F'+str(i)+'=0,"-",data_アナリティクス!$F'+str(i)+'))',chart)
			worksheet.write_formula('H'+str(21+i),'=IF(data_アナリティクス!$G'+str(i)+'="","",VALUE(data_アナリティクス!$G'+str(i)+'))',chart)
			worksheet.write_formula('I'+str(21+i),'=IFERROR(J'+str(21+i)+'/H'+str(21+i)+',"")',chart2)
			worksheet.write_formula('J'+str(21+i),'=IF(data_アナリティクス!$I'+str(i)+'="","",VALUE(data_アナリティクス!$I'+str(i)+'))',chart)
			worksheet.write_formula('K'+str(21+i),'=IFERROR(L'+str(21+i)+'/J'+str(21+i)+',"")',chart2)
			worksheet.write_formula('L'+str(21+i),'=IF(data_アナリティクス!$K'+str(i)+'="","",VALUE(data_アナリティクス!$K'+str(i)+'))',chart)
			worksheet.write_formula('M'+str(21+i),'=IF(data_アナリティクス!$AE'+str(i)+'="","",VALUE(data_アナリティクス!$AE'+str(i)+'))',chart)
			worksheet.write_formula('N'+str(21+i),'=IFERROR(M'+str(21+i)+'/J'+str(21+i)+',"")',chart)
			worksheet.write_formula('O'+str(21+i),'=IFERROR(M'+str(21+i)+'/L'+str(21+i)+',"")',chart)
		worksheet.conditional_format('A23:H'+str(a+20), {'type':     'formula',
	                                    'criteria': '=AND($L23>=1,$L23<=500)',
	                                    'format':    chartbgyellow})
		worksheet.conditional_format('J23:J'+str(a+20), {'type':     'formula',
	                                    'criteria': '=AND($L23>=1,$L23<=500)',
	                                    'format':    chartbgyellow})
		worksheet.conditional_format('L23:O'+str(a+20), {'type':     'formula',
	                                    'criteria': '=AND($L23>=1,$L23<=500)',
	                                    'format':    chartbgyellow})
		worksheet.conditional_format('I23:I'+str(a+20), {'type':     'formula',
	                                    'criteria': '=AND($L23>=1,$L23<=500)',
	                                    'format':    chart2bgyellow})
		worksheet.conditional_format('K23:K'+str(a+20), {'type':     'formula',
	                                    'criteria': '=AND($L23>=1,$L23<=500)',
	                                    'format':    chart2bgyellow})
		total=workbook.add_format({'align':'right','font_name':'Meiryo UI','font_size':11,'border':1,'top':6,'num_format':0x03})
		total4=workbook.add_format({'align':'right','font_name':'Meiryo UI','font_size':11,'border':1,'top':6,'num_format':0x0a})
		worksheet.merge_range('A'+str(a+21)+':G'+str(a+21),'計',total)
		worksheet.write_formula('H'+str(a+21),'=SUM(H23:H'+str(a+20)+')',total)
		worksheet.write_formula('I'+str(a+21),'=IFERROR(J'+str(a+21)+'/H'+str(a+21)+',"")',total4)
		worksheet.write_formula('J'+str(a+21)+'','=SUM(J23:J'+str(a+20)+')',total)
		worksheet.write_formula('K'+str(a+21)+'','=IFERROR(L'+str(a+21)+'/J'+str(a+21)+',"")',total4)
		worksheet.write_formula('L'+str(a+21)+'','=SUM(L23:L'+str(a+20)+')',total)
		worksheet.write_formula('M'+str(a+21)+'','=SUM(M23:M'+str(a+20)+')',total)
		worksheet.write_formula('N'+str(a+21)+'','=IFERROR(M'+str(a+21)+'/J'+str(a+21)+',"")',total)
		worksheet.write_formula('O'+str(a+21)+'','=IFERROR(M'+str(a+21)+'/L'+str(a+21)+',"")',total)
		worksheet.merge_range('A6:C7','=H'+str(a+21),bignum)
		worksheet.merge_range('D6:F7','=I'+str(a+21),bignum2)
		worksheet.merge_range('G6:I7','=J'+str(a+21),bignum)
		worksheet.merge_range('J6:L7','=L'+str(a+21),bignum)
		worksheet.merge_range('M6:O7','=M'+str(a+21),bignum)
#	center cells end

#	device cells start
		worksheet.merge_range('A'+str(a+24)+':B'+str(a+24),'デバイス',twentytwo1)
		worksheet.write('C'+str(a+24),'表示回数',twentytwo2)
		worksheet.write('D'+str(a+24),'クリック率',twentytwo2)
		worksheet.write('E'+str(a+24),'クリック数',twentytwo2)
		worksheet.write('F'+str(a+24),'応募率',twentytwo2)
		worksheet.write('G'+str(a+24),'応募数',twentytwo2)
		worksheet.merge_range('I'+str(a+24)+':J'+str(a+24),'種別',twentytwo1)
		worksheet.write('K'+str(a+24),'表示回数',twentytwo2)
		worksheet.write('L'+str(a+24),'クリック率',twentytwo2)
		worksheet.write('M'+str(a+24),'クリック数',twentytwo2)
		worksheet.write('A'+str(a+23),'■デバイス別実績詳細',twentyone)
		worksheet.write('I'+str(a+23),'■オーガニック検索対スポンサー広告　実績詳細',twentyone)
		device=workbook.add_format({'font_size':11,'font_name':'Meiryo UI','border':1,'top':7,'num_format':0x03})
		device2=workbook.add_format({'font_size':11,'font_name':'Meiryo UI','border':1,'top':7,'num_format':0x0a})
		worksheet.merge_range('A'+str(a+25)+':B'+str(a+25),'PC',device)
		worksheet.merge_range('A'+str(a+26)+':B'+str(a+26),'スマートフォン',device)
		worksheet.write_formula('C'+str(a+25),'=SUM(data_アナリティクス!$O$2:$O$100)',device)
		worksheet.write_formula('C'+str(a+26),'=SUM(data_アナリティクス!$R$2:$R$100)',device)
		worksheet.write_formula('D'+str(a+25),'=E'+str(a+25)+'/C'+str(a+25),device2)
		worksheet.write_formula('D'+str(a+26),'=E'+str(a+25)+'/C'+str(a+26),device2)
		worksheet.write_formula('E'+str(a+25),'=SUM(data_アナリティクス!$P$2:$P$100)',device)
		worksheet.write_formula('E'+str(a+26),'=SUM(data_アナリティクス!$T$2:$T$100)',device)
		worksheet.write_formula('F'+str(a+25),'=G'+str(a+25)+'/E'+str(a+25),device2)
		worksheet.write_formula('F'+str(a+26),'=G'+str(a+26)+'/E'+str(a+26),device2)
		worksheet.write_formula('G'+str(a+25),'=SUM(data_アナリティクス!$Q$2:$Q$100)',device)
		worksheet.write_formula('G'+str(a+26),'=SUM(data_アナリティクス!$T$2:$T$100)',device)
		worksheet.merge_range('I'+str(a+25)+':J'+str(a+25),'スポンサー',device2)
		worksheet.merge_range('I'+str(a+26)+':J'+str(a+26),'オーガニック',device2)
		worksheet.write_formula('K'+str(a+25),'=SUM(data_アナリティクス!$G$2:$G$100)',device)
		worksheet.write_formula('K'+str(a+26),'=SUM(data_アナリティクス!$L$2:$L$100)',device)
		worksheet.write_formula('L'+str(a+25),'=M'+str(a+25)+'/K'+str(a+25),device2)
		worksheet.write_formula('L'+str(a+26),'=M'+str(a+26)+'/K'+str(a+26),device2)
		worksheet.write_formula('M'+str(a+25),'=SUM(data_アナリティクス!$I$2:$I$100)',device)
		worksheet.write_formula('M'+str(a+26),'=SUM(data_アナリティクス!$M$2:$M$100)',device)
#	device cells end
#	bottom cells start
		worksheet.write('A'+str(a+28),'■過去実績比較',twentyone)
		
		worksheet.write('A'+str(a+29),'月',twentytwo1)
		worksheet.write('B'+str(a+29),'表示回数',twentytwo2)
		worksheet.write('C'+str(a+29),'クリック率',twentytwo2)
		worksheet.write('D'+str(a+29),'クリック数',twentytwo2)
		worksheet.write('E'+str(a+29),'応募率',twentytwo2)
		worksheet.write('F'+str(a+29),'応募数',twentytwo2)
		worksheet.write('I'+str(a+29),'クリック単価',twentytwo2)
		worksheet.merge_range('G'+str(a+29)+':H'+str(a+29),'合計費用',twentytwo2)
		worksheet.merge_range('J'+str(a+29)+':K'+str(a+29),'応募単価',twentytwo2)
		sixtyfour=workbook.add_format({'bg_color':'#CEE3F6','font_name':'Meiryo UI','border':1,'font_name':'Meiryo UI'})
		
		total2=workbook.add_format({'font_size':11,'font_name':'Meiryo UI','border':1,'top':6,'num_format':0x03})
		total3=workbook.add_format({'font_size':11,'font_name':'Meiryo UI','border':1,'top':6,'num_format':0x0a})
		
		if len(dataframe3.values[:,0])>8:
			b=len(dataframe3.values[:,0])+2
		else:
			b=10
		for i in range(2,b):
			worksheet.write_formula('A'+str(i+a+28),'=IF(data_過去!H'+str(i)+'="","",data_過去!H'+str(i)+')',chart)
			worksheet.write_formula('B'+str(i+a+28),'=IF(data_過去!I'+str(i)+'="","",VALUE(data_過去!I'+str(i)+'))',chart)
			worksheet.write_formula('C'+str(i+a+28),'=IF(data_過去!C'+str(i)+'="","",VALUE(data_過去!C'+str(i)+'))',chart2)
			worksheet.write_formula('D'+str(i+a+28),'=IF(data_過去!B'+str(i)+'="","",VALUE(data_過去!B'+str(i)+'))',chart)
			worksheet.write_formula('E'+str(i+a+28),'=IF(data_過去!G'+str(i)+'="","",VALUE(data_過去!G'+str(i)+'))',chart2)
			worksheet.write_formula('F'+str(i+a+28),'=IF(data_過去!F'+str(i)+'="","",VALUE(data_過去!F'+str(i)+'))',chart)
			worksheet.merge_range('G'+str(i+a+28)+':H'+str(i+a+28),'=IF(data_過去!D'+str(i)+'="","",VALUE(data_過去!D'+str(i)+'))',chart)
			worksheet.write_formula('I'+str(i+a+28),'=IF(data_過去!A'+str(i)+'="","",VALUE(data_過去!A'+str(i)+'))',chart)
			worksheet.merge_range('J'+str(i+a+28)+':K'+str(i+a+28),'=IF(data_過去!E'+str(i)+'="","",VALUE(data_過去!E'+str(i)+'))',chart)
		worksheet.write('A'+str(a+b+28),'集計',total2)
		worksheet.write_formula('B'+str(a+b+28),'=SUM(B'+str(a+30)+':B'+str(a+b+27)+')',total2)
		worksheet.write_formula('C'+str(a+b+28),'=IFERROR(D'+str(a+b+28)+'/B'+str(a+b+28)+',"")',total3)
		worksheet.write_formula('D'+str(a+b+28),'=SUM(D'+str(a+30)+':D'+str(a+b+27)+')',total2)
		worksheet.write_formula('E'+str(a+b+28),'=IFERROR(F'+str(a+b+28)+'/D'+str(a+b+28)+',"")',total3)
		worksheet.write_formula('F'+str(a+b+28),'=SUM(F'+str(a+30)+':F'+str(a+b+27)+')',total2)
		worksheet.merge_range('G'+str(a+b+28)+':H'+str(a+b+28),'=SUM(G'+str(a+30)+':G'+str(a+b+27)+')',total2)
		worksheet.write_formula('I'+str(a+b+28),'=IFERROR(G'+str(a+b+28)+'/D'+str(a+b+28)+',"")',total2)
		worksheet.merge_range('J'+str(a+b+28)+':K'+str(a+b+28),'=IFERROR(G'+str(a+b+28)+'/F'+str(a+b+28)+',"")',total2)
		worksheet.write('A'+str(a+b+31),'■今後の運用について',twentyone)
		worksheet.merge_range('A'+str(a+b+32)+':O'+str(a+b+34),'',sixtyfour)
#	bottom cells end
#	chart cells start
		chart = workbook.add_chart({'type': 'line'})
		chart.set_size({'width':1050,'height':370})
		chart.add_series({
			'values':'=data_日別!$F$2:$F$'+str(cate+1),
			'categories':'=data_日別!$A$2:$A$'+str(cate+1),
			'marker': {'type': 'circle', 'size': 5,'fill':{'color':'white'}},
			'line':{'color':'#5858FA','width':1},
			'name':'費用',
			'smooth':True,
			
			})
		if (cate-1)>10:
			chart.set_x_axis({'interval_unit': 3})
		else:
			pass
		chart.set_y_axis({'num_font':{'color':'#5858FA'}})

		chart.add_series({
			'values':'=data_日別!$B$2:$B$'+str(cate+1),
			'categories':'=data_日別!$A$2:$A$'+str(cate+1),
			'marker': {'type': 'circle', 'size': 5,'fill':{'color':'white'}},
			'line':{'color':'#FE9A2E','width':1.5},
			'name':'スポンサークリック数',
			'y2_axis':True,
			'smooth':True,
			})
		chart.set_y2_axis({'num_font':{'color':'#FE9A2E'}})

		chart.set_legend({'position': 'bottom','font':{'font_name':'Meiryo UI'}})

		worksheet.insert_chart('A8',chart, {'x_offset':40,'y_offset':5})
#	chart end
		worksheet.print_area('A1:O'+str(a+b+34))
		writer.save()
		workbook.close()
