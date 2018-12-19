import csv, xlsxwriter, itertools, os 



def taxTheFuckUp(directory, inAndOut, csvFile, xlxsName):
	
	os.chdir('/Applications/Python 2.7/Personal scripts/Pasypal scripts')
	
	row = 2
	
	filteredData = []

	Date = []
	Currency = []
	Time = []
	Name = []
	Type = []
	Status = []
	Gross = []
	Net = []
	Email = []
	Transaction = []
	Title = []
	Quanity = []

	text = open(csvFile)
	data = csv.reader(text)

	for i in data:
		if i[7] == 'Gross':
			pass
		else :
			if inAndOut == 'in':
				if float(i[7].replace(',', '')) > 0:
					if i[4] != 'General Currency Conversion':
						filteredData.append(i)
			else:
				if float(i[7].replace(',', '')) < 0:
					filteredData.append(i)

	for i in filteredData:
		Currency.append(i[6])
		Date.append(i[0])
		Time.append(i[1])
		Name.append(i[3])
		Type.append(i[4])
		Status.append(i[5])
		Gross.append(i[7])
		Net.append(i[9])
		Email.append(i[10])
		Transaction.append(i[12])
		Title.append(i[15])
		Quanity.append(i[27])

	os.chdir(directory)

	workbook = xlsxwriter.Workbook(xlxsName)
	worksheet = workbook.add_worksheet()
	
	worksheet.write(0, 0, 'Date')
	worksheet.write(0, 1, 'Time')
	worksheet.write(0, 2, 'Status')
	worksheet.write(0, 3, 'Type')
	worksheet.write(0, 4, 'Name')
	worksheet.write(0, 5, 'Email')
	worksheet.write(0, 6, 'Product title')
	worksheet.write(0, 7, 'Currency')
	worksheet.write(0, 8, 'Net income after fees')
	worksheet.write(0, 9, 'Product quanity')
	worksheet.write(0, 10, 'Transaction ID')



	for date, time, name, typ, status, gross, net, email, transaction, title, quanity, cur in \
	itertools.izip(Date, Time, Name, Type, Status, Gross, Net, Email, Transaction, Title, Quanity, Currency):
		
		if status != 'Pending':
			if typ != 'Account Hold for Open Authorization':
				if typ != 'Reversal of General Account Hold':
					if typ != 'Void of Authorization':
						
						if typ == 'Hidden Virtual PayPal Debit Card Transaction':
							name = 'UPS'
							title = 'Shipping lable'
						else :
							pass 

						worksheet.write(row, 0, date.decode('utf-8'))	
						worksheet.write(row, 1, time.decode('utf-8'))				
						worksheet.write(row, 2, status.decode('utf-8'))
						worksheet.write(row, 3, typ.decode('utf-8'))
						worksheet.write(row, 4, name.decode('utf-8'))
						worksheet.write(row, 5, email.decode('utf-8'))
						worksheet.write(row, 6, title.decode('utf-8'))
						worksheet.write(row, 9, quanity.decode('utf-8'))
						worksheet.write(row, 7, cur)
						worksheet.write(row, 8, net.decode('utf-8'))
						worksheet.write(row, 10, transaction.decode('utf-8'))
						row += 1


	text.close()
	workbook.close()

#taxTheFuckUp('/Users/lylelondraville/Desktop/2016/January', 'in', 'January.CSV', 'January_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/February', 'in', 'February.CSV', 'February_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/March', 'in', 'March.CSV', 'March_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/April', 'in', 'April.CSV', 'April_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/May', 'in', 'May.CSV', 'May_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/June', 'in', 'June.CSV', 'June_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/July', 'in', 'July.CSV', 'July_in_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/August', 'in', 'August.CSV', 'August_in_2016.xlsx')


#taxTheFuckUp('/Users/lylelondraville/Desktop/2016/January', 'out', 'January.CSV', 'January_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/February', 'out', 'February.CSV', 'February_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/March', 'out', 'March.CSV', 'March_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/April', 'out', 'April.CSV', 'April_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/May', 'out', 'May.CSV', 'May_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/June', 'out', 'June.CSV', 'June_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/July', 'out', 'July.CSV', 'July_out_2016.xlsx')
taxTheFuckUp('/Users/lylelondraville/Desktop/2016/August', 'out', 'August.CSV', 'August_out_2016.xlsx')

