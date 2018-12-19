## Logs bot that reads paypal csv files and outwrites to an excel document 

import csv, xlsxwriter, itertools, datetime 

workbook = xlsxwriter.Workbook('')
worksheet = workbook.add_worksheet()

list1 = []
list2 = []
list3 = []


row = 0
col = 0


def time():
  print datetime.datetime.now()

time()

text = open('')


data = csv.reader(text)


for info in data :
	try :
		num = info[3]
		trk = num[4:22]
		if 'UPS' in num:
			
			list1.append(info[0])
			list2.append(trk)
			list3.append(info[9])
	except :
		print "Error"
					

for a,b,c in itertools.izip(list1,list2,list3):
	worksheet.write(row, col,   a)
	worksheet.write(row, col+1, b)
	worksheet.write(row, col+2, c)

	row += 1

text.close()
workbook.close()



time()
