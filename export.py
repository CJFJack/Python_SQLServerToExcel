#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pymssql
import xlwt
import functools


def init_db(host, user, password, database):
	global conn
	# 初始化数据库连接
	conn = pymssql.connect(host=host, user=user, password=password, database=database, charset='utf8')

def db(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
		global cur
		global conn
		# 打开游标
		cur = conn.cursor()
		res = func(*args, **kw)
		#关闭游标
		cur.close()
		return res
    return wrapper

# 打开excel对象
workbook = xlwt.Workbook()
	
def openSt(sheetname):
	global sheet
	# 打开sheet对象
	sheet = workbook.add_sheet(sheetname,cell_overwrite_ok=True)
	
def PrintSQl(sql):
	# 打印sql
	print sql

@db
def __export__(sql):
	global cur
	global conn
	global sheet
	# 执行sql
	cur.execute(sql)
	# 获取所有结果
	data = cur.fetchall()
	# 获取字段
	fields = cur.description
	# 字段信息写入excel
	for field in range(0,len(fields)):
		sheet.write(0,field,fields[field][0])
	# 获取并写入数据段信息
	row = 1
	col = 0
	for row in range(1,len(data)+1):
		for col in range(0,len(fields)):
			sheet.write(row,col,u'%s'%data[row-1][col])
			
def save_file(file):
	# 保存文件
	workbook.save(file)