import tkapi
import tkapi.zaak
from datetime import datetime
import dill as pickle
import traceback
import openpyxl
from tkapi.zaak import ZaakActorRelatieSoort as rel
import json

ODATA_BASE_URL = 'https://gegevensmagazijn.tweedekamer.nl/OData/v4/2.0/Zaak'

def pickle_row(idx, actors):
  with open(f'actors/actor-{idx}.pickle', 'wb') as f:
    pickle.dump(actors, f)

def check_row_pickle(idx):
  try:
    with open(f'actors/actor-{idx}.pickle', 'rb') as f:
      return pickle.load(f)
  except:
    return []


def get():
  zaken_json = []
  
  try:
    with open('kamervragen.pickle', 'rb') as f:
      zaken_json = pickle.load(f)
      print('pickle', f)
  except Exception:
    print(traceback.format_exc())
    pass
  
  if (len(zaken_json) == 0):
    print('Geen zaken gevonden in pickle, ophalen van de API')
    api = tkapi.TKApi()
    
    filter = tkapi.zaak.Zaak.create_filter()
    # filter.filter_date_range(datetime(2024, 10, 1), datetime(2024, 10, 31))
    filter.filter_soort(tkapi.zaak.ZaakSoort.SCHRIFTELIJKE_VRAGEN)

    zaken = api.get_zaken(filter)
    
    zaken_json = []
    for zaak in zaken:
      zaken_json.append(zaak.json)
      
    with open('kamervragen.pickle', 'wb') as f:
      pickle.dump(zaken_json, f)
  else:
    print('Zaken gevonden in pickle, aantal: ' + str(len(zaken_json)))	
  
  print('Aantal zaken: ' + str(len(zaken_json)))
  print('Nu verwerken en schrijven naar xlsx')
  
  workbook = openpyxl.Workbook()
  
  sheet = workbook.active
  
  headers = []
  for prop, _ in zaken_json[0].items():
    headers.append(prop)
    
  headers.append('Gericht aan')
  headers.append('Indiener')
  headers.append('Medeindiener')
  headers.append('Rapporteur')
  headers.append('Volgcommissie')
  headers.append('Voortouwcommissie')
  headers.append('ACTOR ZONDER RELATIE')
  
  sheet.append(headers)
  
  for idx in range(0, len(zaken_json), 100):
    last = idx + 100 if idx + 100 < len(zaken_json) else len(zaken_json)
    print('Verwerken van zaken ' + str(idx) + ' tot ' + str(last))
    
    rows = check_row_pickle(idx)
    
    if len(rows) > 0:
      print('Rij al verwerkt')
      # pretty_json = json.dumps(rows, indent=4, sort_keys=True, ensure_ascii=False)
      # print(pretty_json)

      # sla de rij op
      pass
    else:
      # Verwerk handmatig
      for zaak_json in zaken_json[idx:last]:
        # Voor elke zaak slaan we de data op per rij
        # Verder moeten we voor elke zaak de indieners ophalen
        cols = []

        actors = {
          rel.GERICHT_AAN: [],      
          rel.INDIENER: [],
          rel.MEDEINDIENER: [],
          rel.RAPPORTEUR: [],
          rel.VOLGCOMMISSIE: [],
          rel.VOORTOUWCOMMISSIE: [],
          'ACTOR ZONDER RELATIE': [],
        }
        
        zaak = tkapi.zaak.Zaak(zaak_json)
        
        for actor in zaak.actors:
          if actor == None:
            continue

          name = str(actor.persoon)
          
          if (name == None or name == '' or name == ' ' or name == 'None'):
            name = actor.naam
            
          if not actor.relatie:
            actors['ACTOR ZONDER RELATIE'].append(name)
          else:
            actors[actor.relatie].append(name)


        # Now create a row with the data
        for _, item in zaak_json.items():
          cols.append(item)
          
        for _, actor in actors.items():
          cols.append(', '.join(actor))

        # Row is done, append
        rows.append(cols)
        
      # pickle de rows voor later gebruik
      pickle_row(idx, rows)
        
    # Sla de data op in de worksheet
    for row in rows:
      sheet.append(row)
      
  workbook.save('kamervragen-volledig.xlsx')	

if __name__ == '__main__':
  get()
  