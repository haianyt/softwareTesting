import asyncio
import aiohttp
import xlwt
import xlrd
import json


async def getDataFromApi(url, payload):
	with aiohttp.ClientSession() as session:
		async with session.post(url,data=payload) as resp:
			return await resp.text()

async def testApi(sheet1,sheet2,url,i):
	row = dict()
	row['mins'] = sheet1.cell(i,0).value
	row['times'] = sheet1.cell(i,1).value
	row['remains'] = sheet1.cell(i,2).value
	result = await getDataFromApi(url,row)
	result = json.loads(result)['totalNum']
	for j in range(0,5):
		sheet2.write(i,j,sheet1.cell(i,j).value)
	if float(result) == float(sheet1.cell(i,4).value):
		sheet2.write(i,5,result)
	else:
		sheet2.write(i,5,result,style)



payload = {'mins': '100', 'times': '2','remains':'30'}
url = 'http://localhost:7777/charging'


data = xlrd.open_workbook('话费用例测试1.xls')
sheet1 = data.sheets()[0]

f = xlwt.Workbook()
sheet2 = f.add_sheet('sheet2',cell_overwrite_ok=True)

pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = 2
style = xlwt.XFStyle()
style.pattern = pattern

nrows = sheet1.nrows
for i in range(0,6):
	sheet2.write(0,i,sheet1.cell(0,i).value)

tasks = [testApi(sheet1,sheet2,url,i) for i in range(1,nrows-1)]


loop = asyncio.get_event_loop()
loop.run_until_complete(asyncio.wait(tasks))

f.save('话费用例测试3.xls')


