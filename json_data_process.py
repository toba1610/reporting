import json
import shutil

import time
from datetime import date

def read_JSON_Data(file, topic, parameter=None):

    with open(file) as config:                  #Übergebene JSON File öffnen
        config_data = json.load(config)         #Daten auslesen

    data_type = type(config_data)

    if data_type == dict:

        data_type = type(config_data[topic])

        if data_type == dict:

            data = config_data[topic]

            return data[parameter]                #gesuchten parameter zurückgeben
        
        elif data_type == list:

            return config_data[topic]

def read_JSON_Topic(file, topic):

    with open(file) as config:                  #Übergebene JSON File öffnen
        config_data = json.load(config)         #Daten auslesen

    return config_data[topic]                   

def save_JSON_Data(file, topic, parameter, value):

    with open(file) as target:                      #Ziel Datei öffnen
        target_data = json.load(target)             #Daten auslesen

    target_data[topic][parameter] = value

    # for option in target_data[topic]:               #Ziel Topic suchen
    #     option[parameter] = value                   #Neuen Wert in Ziel Parameter schreiben

    with open(file, 'w') as config:                 #Datei muss neu mit Schreibrechten geöffnet werden, nur einmal öffnen zum lesen und schreiben hat nicht funktioniert
        json.dump(target_data, config, indent=2)    #Daten speichern indent sorgt für die formatierung

def create_new_json_entry(file, Result_Type, result):
    
    with open(file) as target:                      #Ziel Datei öffnen
        target_data = json.load(target)             #Daten auslesen
    
    target_data[Result_Type].append(result)         #Ergebnisse in das Array schreiben
   
    with open(file, 'w') as config:                 #Datei muss neu mit Schreibrechten geöffnet werden, nur einmal öffnen zum lesen und schreiben hat nicht funktioniert
        json.dump(target_data, config, indent=2)

def add_data_to_json_entry(file, Result_Type, Result_Name, result, id):

    with open(file) as target:                      #Ziel Datei öffnen
        target_data = json.load(target)             #Daten auslesen

    if target_data[Result_Type] == []:
        target_data[Result_Type].append({'Result_ID':id})
        for entry in target_data[Result_Type]:
            if entry['Result_ID'] == id or id == 0:                #Prüfen ob der Eintrag für die ID des DUT bestimmt ist
                entry.update({Result_Name:result})
    elif Result_Type == "Result_Verteilung":
        for entry in target_data[Result_Type]:
            entry.update({Result_Name:result})
    else:
        for entry in target_data[Result_Type]:
            if entry['Result_ID'] == id or id == 0:                #Prüfen ob der Eintrag für die ID des DUT bestimmt ist
                entry.update({Result_Name:result})
                tested = True
                break
            else:
                tested = False
            
        if tested == False:
            target_data[Result_Type].append({'Result_ID':id})
            for entry in target_data[Result_Type]:
                if entry['Result_ID'] == id or id == 0:                #Prüfen ob der Eintrag für die ID des DUT bestimmt ist
                    entry.update({Result_Name:result})


    with open(file, 'w') as config:                 #Datei muss neu mit Schreibrechten geöffnet werden, nur einmal öffnen zum lesen und schreiben hat nicht funktioniert
        json.dump(target_data, config, indent=2)

def check_for_nio(file):
    print('OK')

    with open(file) as config:                  #Übergebene JSON File öffnen
        config_data = json.load(config)         #Daten auslesen

    for entry in config_data:
        # print(entry)
        for data in config_data[entry]:
            for value in data:
                print(data[value])
                if data[value] == 'n.i.O':
                    return False
    
    return True

def add_dictonary(destination_file, source_file, dictonary):

    source_data = {}

    source_data = read_JSON_Topic(source_file,dictonary)

    with open(destination_file, 'r') as dest:
        old_data = json.load(dest)

    old_data.update({dictonary:source_data})

    with open(destination_file, 'w') as dest:
        json.dump(old_data, dest, indent=2)

    return

def new_Report(Artikelnummer, Seriennummer):

    Datum = date.today()
    name = 'report' + '_' + Artikelnummer + '_' + str(Seriennummer) + ".json"        
    destination = './Reports/' + name
                                                   
    dest = shutil.copyfile('./Reports/report.json', destination)                                #neuen Report erstellen
    
    save_JSON_Data(destination, "header", "Artikelnummer", Artikelnummer)                       #Artikelnummer im Report eintragen
    save_JSON_Data(destination, "header", "Seriennummer", Seriennummer)                         #dito Seriennummer
    save_JSON_Data(destination, "header", "Datum", str(Datum))                                  #und Datum

    # program = f"./Programms/{Artikelnummer}.json"

    # add_dictonary(destination, program, "Typ_Data")

    return destination #, program

def load_testing_routine(Artikelnummer):

    #Laden des Programmablaufs aus der JSON Datei

    file = 'Programms\\' + Artikelnummer + '.json'

    with open(file, encoding='UTF-8' ) as routine:
        routine_data = json.load(routine)      

    return routine_data['Test_Routine']
 