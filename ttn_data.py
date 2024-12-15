from collections import namedtuple
import time

import requests
import json
from secret_folder.secret_data import token_novaposhta as token
import gspread
from secret_folder.secret_data import sheet_name, sheet_name2, sheet_name3, sheet_name4


def novaposhta_get_barcode(list_of_ttn):
    """
    :param list_of_ttn: list of data with dict with keys "DocumentNumber" and "Phone"(optional)
    :return: dict with keys "DocumentNumber" and values "ClientBarcode"
    """
    url = 'https://api.novaposhta.ua/v2.0/json/'
    # list_of_ttn = []
    # for ob in list_of_ttn1:
    #     if len(ob) == 14:
    #         list_of_ttn.append(ob)
    # print(list_of_ttn)

    data = {
        "apiKey": token,
        "modelName": "TrackingDocumentGeneral",
        "calledMethod": "getStatusDocuments",
        "methodProperties": {
            "Documents": [list_of_ttn, ]

        }
    }

    json_data = json.dumps(data, indent=4, ensure_ascii=False)

    response = requests.post(url=url, data=json_data)

    parse_data = json.loads(response.content.decode('utf-8'))
    ttn_barcode = dict()

    for el in parse_data['data']:
        if el['StatusCode'] == ('102' or '103'):
            ttn_barcode[el['Number']] = el['DocumentCost']

    # StatusCode  102 - Відмова від отримання (Відправником створено замовлення на повернення)
    #             103 - Відмова одержувача (отримувач відмовився від відправлення)
    return ttn_barcode

gc = gspread.service_account(filename='secret_folder/service_account.json')
# dropshippers = (sheet_name, sheet_name3, sheet_name4)
dropshippers = (sheet_name4, )


for sheet in dropshippers:
    sh = gc.open(sheet)
    # print(sh.sheet1.col_values(17))
    d = sh.sheet1.col_values(17)
    t = 0
    for row in d:
        # print(row)
        ttn = novaposhta_get_barcode(row)
        t += 1
        if ttn:
            ttn_amount = [i for i in ttn.values()][0]
            ttn_no = [i for i in ttn.keys()][0]
            row1 = sh.sheet1.find(query=f'{ttn_no}')
            if sh.sheet1.cell(row=row1.row, col=12).value is None:
                print(ttn)
                sh.sheet1.update_cell(row=row1.row, col=12, value=float(ttn_amount))

                sh.sheet1.update_cell(row=row1.row, col='10', value='Повернення')
                sh.sheet1.format(ranges=sh.sheet1.cell(row=row1.row, col='10').address,
                                     format={
                                         'backgroundColor':
                                             {
                                                 'red': 1,
                                                 'blue': 0.0,
                                                 'green': 0.6,
                                                 'alpha': 0.0
                                             },
                                         'textFormat':
                                             {
                                                 'fontFamily': 'Times New Roman',
                                                 'fontSize': 12
                                             }
                                     }
                                 )
                sh.sheet1.format(ranges=sh.sheet1.cell(row=row1.row, col='11').address,
                                     format={
                                         'backgroundColor':
                                             {
                                                 'red': 1.0,
                                                 'blue': 1.0,
                                                 'green': 1.0,
                                                 'alpha': 1.0
                                             },
                                     }
                                 )
                time.sleep(10)
        print(t)
        if t == 59:
            print("I'm sleeping now for 30 sec")
            time.sleep(30)
            t = 0
