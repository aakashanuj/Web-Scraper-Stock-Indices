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

#List of indices
indices=[
["NIFTY",""],
["500",""],
["NIFTY","%20JUNIOR"],
["MIDCAP",""],
["BANK",""],
["IT",""],
["REALTY",""],
["INFRA",""]
]

# Browser
br = mechanize.Browser()

#File
myfile = open('./Reports/NSE_Report.csv', 'w')
wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)

#Time
start_day=sys.argv[1]
end_day=sys.argv[2]
sd=start_day.split('/')[0]
sm=start_day.split('/')[1]
sy=start_day.split('/')[2]
ed=end_day.split('/')[0]
em=end_day.split('/')[1]
ey=end_day.split('/')[2]
start_day=sd+'-'+sm+'-'+sy
end_day=ed+'-'+em+'-'+ey
date_today=end_day
date_then=start_day

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
br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1'),('Accept', '*/*')]

print "Obtaining data for NSE..."

for item in indices:
	if(item[1]!="%20JUNIOR"):
		print "Fetching data for "+item[0]
	else:
		print "Fetching data for "+item[0]+" "+item[1].lstrip("%20")
	str1=item[0]
	str2=item[1]
	url = "http://www.nseindia.com/products/dynaContent/equities/indices/historicalindices.jsp?indexType=CNX%20"+str1+str2+"&fromDate="+date_then+"&toDate="+date_today 
	r=br.open(url)
	data=r.read()
	soup=BeautifulSoup(data)
	table_data=soup.findAll("td",{"class":"date"})
	day_first= table_data[0].text
	day_last= table_data[len(table_data)-1].text
	table_data=soup.findAll("td",{"class":"number"})
	close_last= table_data[3].text
	close_current=table_data[len(table_data)-3].text

	diff=float(close_current)-float(close_last)
	diff_file="%.2f" %((diff*100.0)/float(close_last))

	if(str2==''):
		wr.writerow(["CNX "+str1])
	else:
		wr.writerow(["CNX "+str1+" "+str2.strip('%20')])
	wr.writerow([day_last,close_current])
	wr.writerow([day_first,close_last])
	wr.writerow(["CHANGE(%)",diff_file])
	wr.writerow([" "])

print "Finished successfully !"
myfile.close()	
