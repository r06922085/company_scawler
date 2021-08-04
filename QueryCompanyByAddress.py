# -*- coding: UTF-8 -*-

from bs4 import BeautifulSoup
import openpyxl
import os.path
import requests
import sys
import time

request_headers = {
    'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36',
    'referer':'http://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do',
}

form_data = {
	'errorMsg': '',
	'validatorOpen': 'N',
	'rlPermit': '0',
	'userResp': '', 
	'curPage': '0',
	'fhl': 'zh_TW',
	'qryCond': '',
	'infoType': 'A',
	'qryType': 'cmpyType',
	'cmpyType': 'true',
	'brCmpyType': '',
	'busmType': '',
	'factType': '',
	'lmtdType': '',
	'isAlive': 'true',
	'busiItemMain': '', 
	'busiItemSub': '',
}

host = "https://findbiz.nat.gov.tw"

def main(argv):
	queryAddress = argv[0]
	recordFileName = argv[1]

	results = QueryCompanyDetail(queryAddress)
	ExportResult(results, recordFileName+'.xlsx')

def QueryCompanyDetail(queryAddress):
	form_data['qryCond'] = queryAddress

	totalPage = 1
	isSetTotalPage = False
	currentPage = 0
	results = list()
	count = 0
	businessAccountingNoSet = set()

	while currentPage < totalPage:
		form_data['curPage'] = str(currentPage)
		res = requests.post(
	            host + "/fts/query/QueryList/queryList.do", headers=request_headers, data=form_data)
		res.encoding = 'utf8'
		soup = BeautifulSoup(res.text, "html.parser")
		time.sleep(2)

		if isSetTotalPage == False:
			if soup.find("input", id="totalPage") is not None:
				totalPage = int(soup.find("input", id="totalPage").get('value'))
			isSetTotalPage = True
		
		contentBlocks = soup.find_all("div", {"class", "panel panel-default"})
		for contentBlock in contentBlocks:
			try:
				content = contentBlock.find_all("div")[1].text
			except IndexError:
				print("No record.")
				break
			index = content.find("統一編號")
			businessAccountingNo = content[index+5:index+13]

			if businessAccountingNo not in businessAccountingNoSet:
				# Invoke API to obtain company detail by business account number
				try:
					results.append(requests.get("https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq " + businessAccountingNo + "&$skip=0&$top=50").json()[0])
					businessAccountingNoSet.add(businessAccountingNo)
				except ValueError:
					print("Current business Account Number is incorrect. Continue to parse next record.")
					continue

		currentPage += 1
		print(str(currentPage)+"/"+str(totalPage))

	return results

def ExportResult(results, recordFileName):

	if os.path.isfile(recordFileName):
		workbook = openpyxl.load_workbook(recordFileName)
	else:
		workbook = openpyxl.Workbook()

	sheet = workbook.active
	sheet['A1'] = '統一編號'
	sheet['B1'] = '公司名稱'
	sheet['C1'] = '資本總額'
	sheet['D1'] = '代表人姓名'
	sheet['E1'] = '公司所在地'
	sheet['F1'] = '核准設立日期'

	for result in results:
		if result['Capital_Stock_Amount'] < 1000000:
			continue
		result['Capital_Stock_Amount'] = int(result['Capital_Stock_Amount']/1000)
		result['Capital_Stock_Amount'] = format(result['Capital_Stock_Amount'], ',')
		sheet.append([  result['Business_Accounting_NO'],
						result['Company_Name'],
						result['Capital_Stock_Amount'],
						result['Responsible_Name'],
						result['Company_Location'],
						result['Company_Setup_Date']])

	workbook.save(recordFileName)

if __name__ == "__main__":
   main(sys.argv[1:])