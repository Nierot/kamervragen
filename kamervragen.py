import tkapi
import tkapi.zaak
from datetime import datetime
import openpyxl

ODATA_BASE_URL = 'https://gegevensmagazijn.tweedekamer.nl/OData/v4/2.0/Zaak'
  
# def req():
#   filters = [
#     '?$filter=Soort eq \'Schriftelijke vragen\'',
#     '&format=application/json;odata.metadata=full',
#   ]
  
#   params = {
#     '$filter': "Soort eq 'Kamervragen' or Soort eq 'Schriftelijke vragen' and GestartOp gt 2024-10-01",
#     '$count': "true",
#     '$format': "application/json;odata.metadata=full"
#   }
  
#   req = requests.PreparedRequest()
  
#   req.prepare_url(ODATA_BASE_URL, params)
  
#   print(req.url)

def get():
  api = tkapi.TKApi()
  
  filter = tkapi.zaak.Zaak.create_filter()
  # filter.filter_date_range(datetime(2024, 10, 1), datetime(2024, 10, 31))
  filter.filter_soort(tkapi.zaak.ZaakSoort.SCHRIFTELIJKE_VRAGEN)

  zaken = api.get_zaken(filter)
  
  workbook = openpyxl.Workbook()
  
  sheet = workbook.active

  headers = []
  for prop, _ in vars(zaken[0])().items():
    headers.append(prop)
  
  sheet.append(headers)

  for zaak in zaken:
    row = []
    for prop, value in vars(zaak)().items():
      row.append(value)
    sheet.append(row)
  
  workbook.save('kamervragen.xlsx')	


if __name__ == '__main__':
  get()
  
# https://gegevensmagazijn.tweedekamer.nl/OData/v4/2.0/Zaak?&$filter=Soort%20eq%20%27Schriftelijke%20vragen%27