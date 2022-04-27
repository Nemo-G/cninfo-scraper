import xlsxwriter
import requests
from bs4 import BeautifulSoup as BS
from datetime import date, timedelta, datetime
today = date.today()
yesterday = today - timedelta(days=1)
timePattern = "%Y%m%d"
d0 = today.strftime(timePattern)

####################################################################


def pageStr(page):
	pageString = str(page)
	while(len(pageString)<3):
		pageString = '0' + pageString
	return pageString


urlPattern = "https://www.release.tdnet.info/inbs/I_list_%s_%s.html"

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("[%s]_tdnet_%d.xlsx" % (d0, datetime.now().timestamp()))
worksheet = workbook.add_worksheet()
worksheetRow = 0
pageNum = 1

# Write page header
res = worksheet.write(0, 0, "Time")
worksheet.write(0, 1, "Code")
worksheet.write(0, 2, "Company")
worksheet.write(0, 3, "Title")
worksheet.write(0, 4, "Exchange")
while(1):
	url = urlPattern % (pageStr(pageNum), d0)
	
	r = requests.get(url)
	if r.status_code != requests.codes.ok:
		if r.status_code == requests.codes.not_found:
			print("Done.")
		else:
			print("Error requesting with code %d, current page is %d, break", r.status_code, pageNum)
		break
	
	data = BS(r.content, "html.parser")
	table = data.find(id="main-list-table").find_all("tr")



	for info in table:
		# Write some numbers, with row/column notation.
		worksheetRow += 1
		worksheet.write(worksheetRow, 0, info.find(class_="kjTime").get_text())
		worksheet.write(worksheetRow, 1, info.find(class_="kjCode").get_text())
		worksheet.write(worksheetRow, 2, info.find(class_="kjName").get_text())
		href = info.find(class_="kjTitle").find('a')
		worksheet.write_url(worksheetRow, 3, href['href'], string="https://www.release.tdnet.info/inbs/" + href.contents[0])

		# worksheet.write(worksheetRow, 3, info.find(class="kjTitle").get_text())
		worksheet.write(worksheetRow, 4, info.find(class_="kjPlace").get_text())

	print("Writing page %d, %d records added to sheet" % (pageNum, worksheetRow))
	# in case website blocks us for too frequent query
	pageNum += 1

	if pageNum > 100:
		break


workbook.close()
