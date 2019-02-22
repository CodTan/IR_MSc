import xlrd # for reading .xls file
import MySQLdb # for connection to MySQL
from fnmatch import fnmatch # string matching

book = xlrd.open_workbook("C:\Users\Tulip\Desktop\IR.xlsx") # open the .xls file
sheet = book.sheet_by_name("tanya") # name of the sheet in .xls file

database = MySQLdb.connect (host="127.0.0.1", port=3306, user = "root", passwd = "tulip", db = "MyPython") # connect to MySQL server using these credentials

cursor = database.cursor() # begin reading the database, place the cursor in the beginning

query = """CREATE TABLE data (ID INT NOT NULL AUTO_INCREMENT, Identifier VARCHAR(10), Title MEDIUMTEXT, Authors MEDIUMTEXT, Address MEDIUMTEXT, Abstract MEDIUMTEXT, Citations VARCHAR(10), Publication TINYTEXT, Category TEXT, Keywords TEXT, Publication_Yr VARCHAR(10), PRIMARY KEY(ID));""" # store the complete tuples

cursor.execute(query)

query = """CREATE TABLE garbage LIKE data;""" # to store the incomplete (garbage) tuples

cursor.execute(query)

query_data = """INSERT INTO data (Identifier, Title, Authors, Address, Abstract, Citations, Publication, Category, Keywords, Publication_Yr) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

query_garbage = """INSERT garbage (Identifier, Title, Authors, Address, Abstract, Citations, Publication, Category, Keywords, Publication_Yr) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

g = 0 # garbage flag
author_g =  0 # author garbage flag
unequal_g = 0

for r in range(1, sheet.nrows):
	if fnmatch(str(sheet.cell(r,0).value), '*NA*') or sheet.cell(r,0).value == "": # if cell is empty or no value exists
		g = 1 # set garbage flag
	Identifier = sheet.cell(r,0).value
	if fnmatch(sheet.cell(r,1).value, 'NA') or sheet.cell(r,1).value == "":
		g = 1
	Title = sheet.cell(r,1).value
	if fnmatch(sheet.cell(r,2).value, 'NA') or sheet.cell(r,2).value == "":
		author_g = 1 # author garbage set for empty or missing value in cells
	val = sheet.cell(r,2).value # when none of the above cells are empty, begin reading author info
	temp = val.split(';') # split author names delimited by ';' ... temp has author name
	Authors = [] # contains all authors
	for item in temp:
		Authors.append(item.strip())
	Author_count = len(Authors)

	Authors_org = val

	if fnmatch(sheet.cell(r,3).value, 'NA') or sheet.cell(r,3).value == "": # if university is empty
		g = 1

	val = sheet.cell(r,3).value
	temp = val.split(',') # university information delimited by ','

	A = []
	for item in temp:
		A.append(item.strip()) # strip of leading and trailing whitespaces if any

	A_inter = []
	for item in A:
		if fnmatch(str(item), '*Univ*') or fnmatch(str(item), '*Coll*') or fnmatch(str(item), '*Inst*'):
			A_inter.append(item) # read and store info of organization in which tuple contains any of the above strings

	A_inter_2 = []
	for item in A_inter:
		A_inter_2.append(str(item).split(';'))

	A_inter_3 = []
	for item in A_inter_2:
		for item_in in item:
			A_inter_3.append(item_in.strip())

	University = []
	for item in A_inter_3:
		if fnmatch(str(item), '*Univ*') or fnmatch(str(item), '*Coll*') or fnmatch(str(item), '*Inst*'):
			University.append(item)

	Final_univ = []
	for item in University:
		Final_univ.append(str(item).split(']'))

	Finaler_univ = []
	for item in Final_univ:
		for item_in in item:
			Finaler_univ.append(item_in.strip())

	Finalest_univ = []
	for item in Finaler_univ:
		if fnmatch(str(item), '*Univ*') or fnmatch(str(item), '*Coll*') or fnmatch(str(item), '*Inst*'):
			Finalest_univ.append(item) # store final university info

	University_count = len(Finalest_univ)

	Address = val

	if University_count != Author_count:
		unequal_g = 1 # university information is not available for every author
		g = 1 # set garbage count when not enough university or affiliation information is available for each author

	if fnmatch(sheet.cell(r,4).value, 'NA') or sheet.cell(r,4).value == "":
		g = 1
	Abstract = sheet.cell(r,4).value
	if fnmatch(str(sheet.cell(r,5).value), '*NA*') or sheet.cell(r,5).value == "":
		g = 1
	Citations = sheet.cell(r,5).value
	if fnmatch(sheet.cell(r,6).value, 'NA') or sheet.cell(r,6).value == "":
		g = 1
	Publication = sheet.cell(r,6).value
	if fnmatch(sheet.cell(r,7).value, 'NA') or sheet.cell(r,7).value == "":
		g = 1
	Category = sheet.cell(r,7).value
	if fnmatch(sheet.cell(r,8).value, 'NA') or sheet.cell(r,8).value == "":
		g = 1
	Keywords = sheet.cell(r,8).value
	if fnmatch(str(sheet.cell(r,9).value), '*NA*') or sheet.cell(r,9).value == "":
		g = 1
	Publication_Yr = sheet.cell(r,9).value

	if g == 1:
		if unequal_g == 1:
			values = (Identifier, Title, Authors_org, Address, Abstract, Citations, Publication, Category, Keywords, Publication_Yr)
			cursor.execute(query_garbage, values) # store info in garbage table if info is missing
		g = 0
		unequal_g = 0
	else:
		for i in range(0,Author_count):
			values = (Identifier, Title, Authors[i], Finalest_univ[i], Abstract, Citations, Publication, Category, Keywords, Publication_Yr)
			cursor.execute(query_data, values) # store info in data table

cursor.close()
database.commit()
database.close()

print 'Data imported to MySQL successfully!'
