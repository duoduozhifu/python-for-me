import xlrd
def zhoubao():
    path = "weekport.xlsx"
    wb = xlrd.open_workbook(f"{path}")
    sheet0 = wb.sheet_by_index(0)
    sheet1 = wb.sheet_by_index(1)
    rows0 = sheet0.nrows
    rows1 = sheet1.nrows
    dict0 = {}
    dict1 = {}
    for i in range(1, rows0):
        name0 = sheet0.cell_value(i, 0)
        count0 = int(sheet0.cell_value(i, 1))
        non_compliance = int(sheet0.cell_value(i, 2))
        compliance = int(sheet0.cell_value(i, 3))
        print(f'完成{name0} {count0}台主机基线扫描，共发现合规项数量{compliance}个，其中不合规项数量{non_compliance}个，报告已发送给厂商，等待整改后复测。')
        #dict0={name0:[count0,non_compliance,compliance]}
        #a = dict0.keys()
        #print(dict0)
    print('**************************************************************************************************************************')
    for o in range(1, rows1):
        name1 = sheet1.cell_value(o, 0)
        count1 = int(sheet1.cell_value(o, 1))
        highriskvul = int(sheet1.cell_value(o, 2))
        medriskvul = int(sheet1.cell_value(o, 3))
        print(f'完成{name1} {count1}台主机漏洞扫描，其中高危漏洞{highriskvul}个，中危漏洞{medriskvul}个；报告已发送给厂商，等待整改后复测。')
        dict1={name1:[count1,non_compliance,compliance]}
        #b = dict1.keys()
        #print(dict1)
'''
    for v in dict0.keys():
        if v in dict1.keys():
            print(f'完成{dict0.keys()} {list(dict0.values())[0]}台主机基线扫描，共发现不合规项数量{list(dict0.values())[1]}个，合规项shul{list(dict0.values())[2]}个；完成漏洞扫描，其中高危{list(dict1.values())[1]}个，中危{list(dict1.values())[2]}个；报告已发送给厂商，等待整改后复测。')
        else:
            



    for k in dict0.keys():
        if not k in dict1.keys():
            print(f'完成{name0} {count0}台主机基线扫描，其中不合规项数量{non_compliance}个，合规项数量{compliance}个，报告已发送给厂商，等待整改后复测。')
    for v in dict1.keys():
        if not v in dict0.keys():
            print(f'完成{name1} {count1}台主机漏洞扫描，其中高危漏洞{highriskvul}个，中危漏洞{medriskvul}个；报告已发送给厂商，等待整改后复测。')
        else:
            print(f'完成{name0} {count0}台主机基线扫描，共发现不合规项数量{non_compliance}个，合规项shul{compliance}个；完成漏洞扫描，其中高危{highriskvul}个，中危{medriskvul}个；报告已发送给厂商，等待整改后复测。')
'''
zhoubao()