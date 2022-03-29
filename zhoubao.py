import xlrd
def weekreport():
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
                print(f'。')
                print(f'。')

            else:
                print(f'。')

weekreport()
