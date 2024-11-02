import openpyxl
from novaposhta import novaposhta_get_barcode

file = '/Users/stanislav_zhucenko/PycharmProjects/excel_automation/test_data/report_moneytransfers_27-10-2024_13-56.xlsx'
workbook_obj = openpyxl.load_workbook(filename=file)

sheet_obj = workbook_obj.active

input_date = '26.10.2024'  # Choose the date of change in the state of money transfer
ttn_obj = sheet_obj['C']  # The column with TTN
amount_obj = sheet_obj['E']  # The column with the amount of money to the TTN
state_obj = sheet_obj['O']  # The column with state of money transfer
date_obj = sheet_obj['P']  # The column with the date of change in the state of money transfer

documents = []

for data in zip(ttn_obj, amount_obj, state_obj, date_obj):
    ttn, amount, state, datetime_obj = data[0].value, data[1].value, data[2].value, data[3].value
    date = datetime_obj.split(' ')[0]  # Removing time from the date
    if state == 'Видано' and date == input_date:
        documents.append(
            {"DocumentNumber": ttn},
        )
    else:
        sheet_obj.delete_rows(data[0].row)

sheet_obj.insert_cols(4)
print(len(documents))

# Divide list of documents for N lists with len < 100
if len(documents) > 100:
    sublists = [documents[i:i+100] for i in range(0, len(documents), 100)]
    barcode = dict()
    for documents_ttn in sublists:
        barcode.update(novaposhta_get_barcode(documents_ttn))
else:
    barcode = novaposhta_get_barcode(documents)  # Get barcodes from Novaposhta API

for ttn in ttn_obj:
    if ttn.value in barcode:
        value_to_set = barcode[ttn.value]
        sheet_obj.cell(row=ttn.row, column=4, value=value_to_set)

sheet_obj.delete_cols(5)

filename = 'mod.xlsx'
workbook_obj.save(filename=filename)

print(f"Your file {filename} is already done!")
