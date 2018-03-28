#!/usr/bin/env python
# -*- coding: utf-8 -*-

import export

host = ''
user = ''
password = ''
database = ''
sheetname = 'example'
file = r'./example.xlsx'
sql = '''

'''

export.init_db(host, user, password, database)
export.openSt(sheetname)
export.__export__(sql)
export.save_file(file)
