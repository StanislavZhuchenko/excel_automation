import time
import gspread
# from excel_modifying import sheet_obj
from novapay_auto import sheet_obj
from collections import namedtuple
from secret_folder.secret_data import sheet_name, sheet_name2, sheet_name3, sheet_name4


gc = gspread.service_account(filename='secret_folder/service_account.json')

dropshippers = (sheet_name, sheet_name2, sheet_name3, sheet_name4)
# dropshippers = (sheet_name4, )
for sheet in dropshippers:
    print(sheet)
    sh = gc.open(sheet)

    order = namedtuple('order', ['ttn', 'amount'])
    # temporary_ttn = ['20450797682084',
    #                  '20450797683412',
    #                  '20450797687777',
    #                  '20450797689060',
    #                  '20450797690087']
    # temporary_amount = [530,
    #                     550,
    #                     370,
    #                     380,
    #                     460]
    # temporary_data = zip(temporary_ttn, temporary_amount)

    # list_of_order = [order(ttn=i[2], amount=i[4]) for i in list(sheet_obj.values)]
    # for Novapay
    list_of_order = [order(ttn=i[8], amount=i[3]) for i in list(sheet_obj.values)]

    # list_of_order = [order(ttn=i[0], amount=i[1]) for i in temporary_data]

    # for j in list_of_order:
    #     print(j.ttn, j.amount)
    t = 0
    for order in list_of_order:
        ttn = order.ttn
        row1 = sh.sheet1.find(query=f'{ttn}')
        if row1:
            g_sheet_amount = int(sh.sheet1.cell(row=row1.row, col=10).value)
        else:
            print('TTN not found in the sheet')
            time.sleep(5)
            continue
        if g_sheet_amount == order.amount:
            sh.sheet1.update_cell(row=row1.row, col='11', value=order.amount)
            sh.sheet1.format(ranges=sh.sheet1.cell(row=row1.row, col='11').address,
                             format={
                                 'backgroundColor':
                                     {
                                         'red': 0.6,
                                         'blue': 0.8,
                                         'green': 0.8,
                                         'alpha': 0.5
                                     },
                                 'textFormat':
                                     {
                                         'fontFamily': 'Times New Roman',
                                         'fontSize': 12
                                     },
                             }
                             )
            # print(f"{sh.sheet1.cell(row=row1.row, col='2').value} - is added by TTN {ttn}")
            print(f"{ttn} is added")
        else:
            print(f"Amount of money from Google Sheets don't equal from Novaposhta, TTN: {ttn}")

        time.sleep(5)
        t += 1
        if t == 9:
            time.sleep(30)
            t = 0
            print("I'm sleeping now for 30 sec")

print("Everything is done")
