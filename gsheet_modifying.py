import gspread

gc = gspread.service_account(filename='secret_folder/service_account.json')

sh = gc.open("Test table")

list_of_cells = [5113, 5114, 5120, 5121, 5122]
for data in list_of_cells:
    row1 = sh.sheet1.find(query=f'{str(data)}', in_column=2).row
    print(sh.sheet1.row_values(row1))
    value = sh.sheet1.cell(row=row1, col='10').value
    sh.sheet1.update_cell(row=row1, col='11', value=value)
    print(sh.sheet1.row_values(row1))

sh.sheet1.update_cell(row=15, col=11, value='=SUM(K1:K13)')
