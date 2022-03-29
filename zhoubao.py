import xlrd
def zhoubao():
    path = "weekport.xlsx"
    wb = xlrd.open_workbook(f"{path}")
    sheet1 = wb.sheet_by_index(0)
    sheet2 = wb.sheet_by_index(1)
    rows1 = sheet1.nrows
    rows2 = sheet2.nrows
    for i in range(1, rows1):
        name1 = sheet1.cell_value(i, 0)
        count1 = int(sheet1.cell_value(i, 1))
        non_compliant = int(sheet1.cell_value(i, 2))
        compliance = int(sheet1.cell_value(i, 3))

        for i in range(1, rows2):
            name2 = sheet2.cell_value(i, 0)
            count2 = int(sheet2.cell_value(i, 1))
            highriskvul = int(sheet2.cell_value(i, 2))
            medriskvul = int(sheet2.cell_value(i, 3))
            if (name1!=name2):
                print(f'完成{name1} {count1}台主机基线扫描，其中不合规项数量{non_compliant}个，合规项数量{compliance}个，报告已发送给厂商，等待整改后复测。')
                print(f'完成{name2} {count2}台主机漏洞扫描，其中高危漏洞{highriskvul}个，中危漏洞{medriskvul}个；报告已发送给厂商，等待整改后复测。')

            else:
                print(f'完成{name1} {count1}台主机基线扫描，其中不合规项数量{non_compliant}个，合规项数量{compliance}个；完成漏洞扫描，其中高危漏洞{highriskvul}个，中危漏洞{medriskvul}个；报告已发送给厂商，等待整改后复测。')

zhoubao()