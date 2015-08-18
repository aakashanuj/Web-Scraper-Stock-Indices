####################################################
#############     Author: AAKASH ANUJ    ###########
#############     CII Automation tool    ###########
####################################################

import httplib 
import urllib 
import urllib2 
from BeautifulSoup import BeautifulSoup
import mechanize
from datetime import datetime, timedelta
from time import gmtime,strftime
import csv
import sys
import cookielib


# Browser
br = mechanize.Browser()

list=["BSE30  ","BSE500 ","AUTO   ","BANKEX ","BSECG  ","BSECD  ","BSEFMCG", "BSEHC  ","MIDCAP ","SMLCAP ","TECK   ","METAL  ","OILGAS "]
myfile = open('./Reports/BSE_Report.csv', 'w')
wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)

#Time
date_today=sys.argv[2]
date_then=sys.argv[1]

# Cookie Jar
cj = cookielib.LWPCookieJar()
br.set_cookiejar(cj)

# Browser options
br.set_handle_equiv(True)
br.set_handle_redirect(True)
br.set_handle_referer(True)
br.set_handle_robots(False)
br = mechanize.Browser()
br.set_handle_robots(False)
br.addheaders = [('User-agent', 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.6) Gecko/20070725 Firefox/2.0.0.6')]

print "Obtaining data for BSE..."
for item in list:
	print "Fetching data for "+item
	url = 'http://www.bseindia.com/indices/IndexArchiveData.aspx?expandable=3'
	br.open(url)
	response = br.response().read()
	br.select_form(nr=0)
	br.set_all_readonly(False)
	br.form['ctl00$ContentPlaceHolder1$txtFromDate']=date_then
	br.form['ctl00$ContentPlaceHolder1$txtToDate']=date_today
	br.form.set_value([item],name='ctl00$ContentPlaceHolder1$ddlIndex')
	br['ctl00$ContentPlaceHolder1$hidInd']=item
	response = br.submit().read()
	soup=BeautifulSoup(response)
	findall=soup.findAll('span')[26]
	close_last=''
	close_current=''
	for i in range(1,len(findall.findAll('tr'))):
		from_date=findall.findAll('tr')[i].findAll('td')[0].getText()
		temp_date_from=date_then[0:6]+date_then.split('/')[2][2:4]
		temp_date_current=date_today[0:6]+date_today.split('/')[2][2:4]

		if(from_date==temp_date_from):
			close_last=findall.findAll('tr')[i].findAll('td')[4].getText().replace(',','')
		if(from_date==temp_date_current):
			close_current=findall.findAll('tr')[i].findAll('td')[4].getText().replace(',','')
	
	#print close_last,close_current
	diff=float(close_current)-float(close_last)
	diff_file="%.2f" %((diff*100.0)/float(close_last))

	wr.writerow([item])
	wr.writerow([temp_date_current,close_current])
	wr.writerow([temp_date_from,close_last])
	wr.writerow(["CHANGE(%)",diff_file])
	wr.writerow([" "])
	
print "Finished successfully !"	
myfile.close()


