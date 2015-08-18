####################################################
#############     Author: AAKASH ANUJ    ###########
#############     CII Automation tool    ###########
####################################################

from bs4 import BeautifulSoup
import csv
import time
import mechanize
import cookielib
import sys 

# Browser
br = mechanize.Browser()

# Cookie Jar
cj = cookielib.LWPCookieJar()
br.set_cookiejar(cj)

# Browser options
br.set_handle_equiv(True)
br.set_handle_redirect(True)
br.set_handle_referer(True)
br.set_handle_robots(False)

# Follows refresh 0 but not hangs on refresh > 0
br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)

br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
map_month_code={
'01':	'Jan',
'02': 	'Feb',
'03':	'Mar',
'04':	'Apr',
'05':	'May',
'06': 	'Jun',
'07':	'Jul',
'08': 	'Aug',
'09':	'Sep',
'10': 	'Oct',
'11':	'Nov',
'12': 	'Dec'	
}

iter_url_map={
1:'http://in.finance.yahoo.com/q/hp?s=%5EDJI',
2:'http://in.finance.yahoo.com/q/hp?s=%5EN225',
3:'http://in.finance.yahoo.com/q/hp?s=%5ESTI',
4:'http://in.finance.yahoo.com/q/hp?s=%5EKS11',
5:'http://in.finance.yahoo.com/q/hp?s=%5EFTSE'
}

iter_message_map={
1:'SOURCE : Dow Jones Industrial Average (^DJI)-DJI',
2:'SOURCE : NIKKEI 225 (^N225) -Osaka',
3:'SOURCE : STI Index (^STI) -SES',
4:'SOURCE : KOSPI Composite Index (^KS11) -KSE',
5:'SOURCE : FTSE 100 (^FTSE) -FTSE'
}

start_day=sys.argv[1]
end_day=sys.argv[2]
sd=start_day.split('/')[0].lstrip('0')
sm=map_month_code[start_day.split('/')[1]]
sy=start_day.split('/')[2]
ed=end_day.split('/')[0].lstrip('0')
em=map_month_code[end_day.split('/')[1]]
ey=end_day.split('/')[2]
start_day=sd+' '+sm+', '+sy
end_day=ed+' '+em+', '+ey

messages=[]
messages.append("Data obtained from YAHOO finance")
print "Obtaining data from Yahoo finance..."


for iterator in range(1,6):
	print "Fetching data for "+iter_message_map[iterator].lstrip("SOURCE : ")
	messages=[]
	r = br.open(iter_url_map[iterator])
	data = r.read()
	soup=BeautifulSoup(data)
	list=soup.find_all(align="right")
	day_one=[]
	day_last=[]
	fortnightly_change=[]
	fortnightly_change.append("FORTNIGHTLY_CHANGE")
	day_one.append("UNTIL DAY")
	day_last.append("FROM DAY")
	
	iter=11
	incr=7
	while(list[iter].get_text()!=start_day):
		iter=iter+incr
	day_last.append(list[iter].get_text())
	close_day_last=list[iter+4].get_text()
	
	iter=11
	incr=7
	while(list[iter].get_text()!=end_day):
		iter=iter+incr
	day_one.append(list[iter].get_text())	
	close_day_one=list[iter+4].get_text()

	close_day_one = close_day_one.replace(",", "")
	close_day_last = close_day_last.replace(",", "")

	day_one.append(close_day_one)
	day_last.append(close_day_last)
	change=float(close_day_one)-float(close_day_last)
	ratio=change/float(close_day_last)
	final_change=ratio*100;
	fortnightly_change.append(str(final_change)+"%")

	myfile = open('./Reports/Yahoo_Finance_Report.csv', 'a')
	wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
	wr.writerow(messages)
	messages=[]
	messages.append(iter_message_map[iterator])
	wr.writerow(messages)
	wr.writerow(day_one)
	wr.writerow(day_last)
	wr.writerow(fortnightly_change)
	
print "Finished successfully !"
myfile.close()	
