import requests
import json
from secret_data import token_novaposhta as token


def novaposhta_get_barcode(list_of_ttn):
    """

    :param list_of_ttn: list of data with dict with keys "DocumentNumber" and "Phone"(optional)
    :return: dict with keys "DocumentNumber" and values "ClientBarcode"
    """
    url = 'https://api.novaposhta.ua/v2.0/json/'

    data = {
        "apiKey": token,
        "modelName": "TrackingDocument",
        "calledMethod": "getStatusDocuments",
        "methodProperties": {
            "Documents": list_of_ttn

        }
    }

    json_data = json.dumps(data, indent=4, ensure_ascii=False)

    response = requests.post(url=url, data=json_data)

    parse_data = json.loads(response.content.decode('utf-8'))

    ttn_barcode = dict()

    for el in parse_data['data']:
        ttn_barcode[el['Number']] = el['ClientBarcode']

    return ttn_barcode
