#!/usr/bin/env python
# -*- coding: utf-8 -*-

import export

host = 'rds3qejaynvbvqj.sqlserver.rds.aliyuncs.com:3433'
user = 'shinetour_dba2'
password = 'shinetour'
database = 'st_insurance'
sheetname = 'example'
file = r'./example.xlsx'
sql = '''

'''

export.init_db(host, user, password, database)
export.openSt(sheetname)
export.__export__(sql)
export.save_file(file)