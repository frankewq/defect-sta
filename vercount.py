#!/usr/bin/python
#coding=utf-8

import xlrd
import xlwt
import sqlite3

# 打开数据库文件
buglist_db = sqlite3.connect('buglist.db')
cursor = buglist_db.cursor()

# 建表
cursor.execute('DROP TABLE IF EXISTS buglist')
cursor.execute('CREATE TABLE buglist (id int PRIMARY KEY, module char(32), level int,category char(16), state char(16), ctime DATETIME)')
 
 
# 打开Excel文件
openf = "20150819-android.xls"
savef = "20150819-android-count.xls"
device_workbook = xlrd.open_workbook(openf)
device_sheet = device_workbook.sheet_by_index(0)

# 读取单元格，并将相应数据插入数据库
for row in range(1, device_sheet.nrows):
	id = device_sheet.cell(row,0).value
	module = device_sheet.cell(row,2).value
	level = device_sheet.cell(row,8).value
	category = device_sheet.cell(row,10).value
	state = device_sheet.cell(row,14).value
	ctime = xlrd.xldate.xldate_as_datetime(device_sheet.cell(row,19).value,0)
	#ctime = xlrd.xldate.xldate_as_tuple(device_sheet.cell(row,19).value,0)
	
	#状态处理，方便sql拼接
	sta = '激活'
	if state.encode("utf-8") == sta:
		state = 1
	else:
		state = 2
#	print '%s, %s, %s, %s' % (id, module, level, state)
	# 避免插入重复记录
	cursor.execute('SELECT * FROM buglist WHERE id=?', (id,))
	res = cursor.fetchone()
	if res == None:
		cursor.execute('INSERT INTO buglist (id, module, level,category, state, ctime) VALUES (?, ?, ?, ?, ?, ?)', (id, module, level, category, state, ctime))
	else:
		print "Something wrong with db！"

buglist_db.commit()


# 写入Excel文件
wb = xlwt.Workbook(encoding = 'ascii')
ws = wb.add_sheet('Count graphic')


#标题
ws.write(0, 0, "MODULE(ALL)")
ws.write(0, 1, "COUNT")
ws.write(0, 2, "MODULE(SER)")
ws.write(0, 3, "COUNT")
ws.write(0, 4, "STATE")
ws.write(0, 5, "COUNT")
ws.write(0, 6, "LEVEL")
ws.write(0, 7, "COUNT")
ws.write(0, 8, "CTIME")
ws.write(0, 9, "COUNT")
ws.write(0, 10, "CATEGORY")
ws.write(0, 11, "COUNT")

#全部问题
print "Step 1"

cursor.execute('select module,count(0) from buglist where state = 1 group by module')
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
#		print '%s	%s' %(d[0].encode('gb2312'),d[1])
		ws.write(n, 0, d[0])
		ws.write(n, 1, d[1])
		n = n-1


#严重问题统计（1、2级）
print "Step 2"

cursor.execute('select module,count(0) from buglist where level <=2 and  state =1  group by module' )
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
#		print '%s	%s' %(d[0].encode('gb2312'),d[1])
		ws.write(n, 2, d[0])
		ws.write(n, 3, d[1])
		n = n-1


#激活/已解决
print "Step 3"
cursor.execute('select state,count(0) from buglist  group by state' )
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
#		print '%s	%s' %(d[0].encode('gb2312'),d[1])
		ws.write(n, 4, d[0])
		ws.write(n, 5, d[1])
		n = n-1

#级别分布
print "Step 4"
cursor.execute('select level,count(0) from buglist  group by level' )
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
		ws.write(n, 6, d[0])
		ws.write(n, 7, d[1])
		n = n-1

#问题创建时间
print "Step 5"
cursor.execute('select ctime,count(0) from buglist  group by ctime order by ctime desc' )
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
		ws.write(n, 8, d[0])
		ws.write(n, 9, d[1])
		n = n-1

#问题分类
print "Step 6 ...Filish!"
cursor.execute('select CATEGORY,count(0) from buglist  group by CATEGORY' )
res = cursor.fetchall()

n = len(res)
for d in res:
	if n > 0:
		ws.write(n, 10, d[0])
		ws.write(n, 11, d[1])
		n = n-1


buglist_db.commit()

wb.save(savef)

