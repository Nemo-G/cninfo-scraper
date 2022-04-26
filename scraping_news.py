import xlsxwriter
import requests
from datetime import date, timedelta, datetime
today = date.today()
yesterday = today - timedelta(days=1)
timePattern = "%Y-%m-%d"
# definition for plate
plates = {
	'深圳主板': 'szmb',
	'创业板':  'szcy',
}

########################## 查询条件  #################################
# 当前板块
activePlateKey = '创业板'

# Passed as query params
sortName, sortType = 'code', 'desc'

# d0 means start date, d1 means end date
d0, d1 = yesterday.strftime(timePattern), yesterday.strftime(timePattern)

####################################################################




url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
headers = {
    "accept": "*/*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
	"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36"
    # "cookie": "JSESSIONID=364788304203B95A8BE574A41A4BBD09; _sp_ses.2141=*; routeId=.uc2; insert_cookie=45380249; _sp_id.2141=29faee1d-3cfa-446c-8ff4-dd647e8e3199.1650973806.1.1650973880.1650973806.28f5bd22-273d-4048-a92e-cc2a6bd5aa19",
    # "Referer": "http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search&startDate=2022-04-26&endDate=2022-04-27",
    # "Referrer-Policy": "strict-origin-when-cross-origin"
  }
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("[%s~%s]_%s%s-%s_%d.xlsx" % (d0, d1, activePlateKey, sortName, sortType, datetime.now().timestamp()))
worksheet = workbook.add_worksheet()
worksheetRow = 0
pageNum = 1

# Write page header
res = worksheet.write(0, 0, "代码")
worksheet.write(0, 1, "简称")
worksheet.write(0, 2, "公告标题")
worksheet.write(0, 3, "PDF link")
worksheet.write(0, 4, "公告时间")
while(1):
	payload = "pageNum=%d&pageSize=30&column=szse&tabName=fulltext&plate=%s&stock=&searchkey=&secid=&category=&trade" \
			  "=&seDate=%s~%s&sortName=%s&sortType=%s&isHLtitle=true" % (pageNum, plates[activePlateKey], d0, d1, sortName, sortType)
	
	r = requests.post(url, data=bytes(payload, 'UTF-8'), headers=headers)
	if r.status_code != requests.codes.ok:
		print("Error requesting with code %d, current page is %d", r.status_code, pageNum)
		break
	
	data = r.json()
	announcements = data['announcements']


	for info in announcements:
		# Write some numbers, with row/column notation.
		worksheetRow += 1
		worksheet.write(worksheetRow, 0, info['secCode'])
		worksheet.write(worksheetRow, 1, info['secName'])
		worksheet.write(worksheetRow, 2, info['announcementTitle'])
		worksheet.write_url(worksheetRow, 3, "http://static.cninfo.com.cn/"+info['adjunctUrl'], string=info['announcementTitle'])
		infoDate = datetime.fromtimestamp(info['announcementTime']/1000.0)
		worksheet.write(worksheetRow, 4, infoDate.isoformat())
	
	print("Writing page %d, %d records added to sheet" % (pageNum, worksheetRow))
	# in case website blocks us for too frequent query
	pageNum += 1

	# It turns out that data will repeat the first page after page 100(records 3000).
	# So we take at most 3000
	if not data['hasMore'] or worksheetRow >= 3000:
		break
	#if worksheetRow >= data['totalRecordNum']: #not data['hasMore'] or pageNum > data['totalpages']
	#	break


workbook.close()
