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
declare  @OpDateStart datetime;
declare  @OpDateEnd datetime;

set @OpDateStart=convert(varchar(7),getdate(),120)+'-01'--开始日期
set @OpDateEnd=convert(varchar(10),getdate(),120)--结束日期

SELECT COUNT(od.OrderNo) AS 采购渠道为BSP的变更单总数
FROM dbo.tp_OrderNo od
left join wf_ProcessHeader ph on od.orderno = ph.warecode
LEFT JOIN dbo.tp_TicketProduct tp ON od.OrderNo = tp.OrderNo
WHERE od.OrderType = 2  AND od.OpDate >=@OpDateStart  and od.OpDate<=@OpDateEnd AND ph.Status='FINISH' AND tp.ProductType = 0
'''
sql2 = '''
declare  @OpDateStart datetime;
declare  @OpDateEnd datetime;

set @OpDateStart=convert(varchar(7),getdate(),120)+'-01'--开始日期
set @OpDateEnd=convert(varchar(10),getdate(),120)--结束日期

SELECT COUNT(od.OrderNo) AS 变更方式为中台在线的变更单总数
FROM dbo.tp_OrderNo od
left join wf_ProcessHeader ph on od.orderno = ph.warecode
LEFT JOIN dbo.tp_TicketProduct tp ON od.OrderNo = tp.OrderNo
WHERE od.OrderType = 2  AND od.OpDate >=@OpDateStart  and od.OpDate<=@OpDateEnd AND ph.Status='FINISH' AND tp.ProductType = 0
AND tp.OnlineType = 1 AND od.OpType <> 0
'''
sql3 = '''
declare  @OpDateStart datetime;
declare  @OpDateEnd datetime;

set @OpDateStart=convert(varchar(7),getdate(),120)+'-01'--开始日期
set @OpDateEnd=convert(varchar(10),getdate(),120)--结束日期

SELECT COUNT(od.OrderNo) AS 自动改签成功总数1
FROM dbo.tp_OrderNo od
left join wf_ProcessHeader ph on od.orderno = ph.warecode
LEFT JOIN dbo.tp_TicketProduct tp ON od.OrderNo = tp.OrderNo
LEFT JOIN dbo.tp_OrderRemark remark ON od.OrderNo = remark.OrderNo
WHERE od.OrderType = 2  AND od.OpDate >=@OpDateStart  and od.OpDate<=@OpDateEnd AND ph.Status='FINISH' AND tp.ProductType = 0
AND tp.OnlineType = 1 AND od.OpType <> 0 AND remark.IssuteWay = 3 
'''
sql4 = '''
declare  @OpDateStart datetime;
declare  @OpDateEnd datetime;

set @OpDateStart=convert(varchar(7),getdate(),120)+'-01'--开始日期
set @OpDateEnd=convert(varchar(10),getdate(),120)--结束日期

SELECT COUNT(DISTINCT od.OrderNo) AS 自动改签成功总数2
FROM dbo.tp_OrderNo od
left join wf_ProcessHeader ph on od.orderno = ph.warecode
LEFT JOIN dbo.tp_TicketProduct tp ON od.OrderNo = tp.OrderNo
LEFT JOIN dbo.tp_OrderRemark remark ON od.OrderNo = remark.OrderNo
LEFT JOIN dbo.tp_WareRemark wr ON od.OrderNo = wr.WareCode
WHERE od.OrderType = 2  AND od.OpDate >=@OpDateStart  and od.OpDate<=@OpDateEnd AND ph.Status='FINISH' AND tp.ProductType = 0
AND tp.OnlineType = 1 AND od.OpType <> 0 AND (wr.Remark LIKE '%该航班已起飞，不允许进行删除操作！%' OR wr.Remark LIKE '%出票失败,航段状态发生变化!%' OR wr.Remark IS NULL)
'''
sql5 = '''
declare  @OpDateStart datetime;
declare  @OpDateEnd datetime;

set @OpDateStart=convert(varchar(7),getdate(),120)+'-01'--开始日期
set @OpDateEnd=convert(varchar(10),getdate(),120)--结束日期

SELECT od.OrderNo AS 变更单号,
(SELECT '<br/>' + rtrim (wr.Remark) FROM dbo.tp_WareRemark wr WHERE wr.WareCode = od.OrderNo AND wr.Remark LIKE '%自动出票失败%' FOR XML PATH('')
) AS 内部备注,
remark.IssuteWay
FROM dbo.tp_OrderNo od
left join wf_ProcessHeader ph on od.orderno = ph.warecode
LEFT JOIN dbo.tp_TicketProduct tp ON od.OrderNo = tp.OrderNo
LEFT JOIN dbo.tp_OrderRemark remark ON od.OrderNo = remark.OrderNo
WHERE od.OrderType = 2  AND od.OpDate >=@OpDateStart  and od.OpDate<=@OpDateEnd AND ph.Status='FINISH' AND tp.ProductType = 0
AND tp.OnlineType = 1 AND remark.IssuteWay is null 
'''

export.init_db(host, user, password, database)
export.openSt(sheetname)
export.__export__(sql)
export.save_file(file)