import win32com
import win32com.client
import json
import os
import time
import logging
import threading
from docx2pdf import convert

def init_logger():
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        handlers=[logging.FileHandler(f".\\log\\app.log"), logging.StreamHandler()])

    global logger
    logger = logging.getLogger(__name__)

class create_report():

    def __init__(self) -> None:
        
        self.app = None
        self.doc = None

    def check_filename(self, filename:str) -> str:

        '''
        Checks the given filename for the following parameters:
            - length
            - file ending
            - file existing

        Outputs a 'error' if anything is wrong
            
        Input:
            - filename: str = The name of the template to use

        Output:
            - str = 'error' if anything goes wrong, or the name of the file
        '''

        if '.' in filename:
            temp = []
            temp = filename.split('.')

        if 'docx' in temp or 'doc' in temp:
            pass
        else:
            logger.warning(f'The filename contains no *.docx or *.doc file, check config.json for possible errors')
            return 'error'

        if len(temp) >2 :
            logger.warning(f'The filename contains multiple "." Characters, check config.json for possible errors')
            return 'error'

        if os.path.exists(f'.\\templates\\{filename}'):
            pass
        else:
            logger.warning(f"The file doesn't exists, check config.json for possible errors")
            return 'error'


        return filename

    def open_word_document(self, filename: str) -> None:

        '''
        Function to connet to word and open the application in the background

        Input:
            - filename: str = The name of the template to use

        Output:
            - None 
        '''

        logger.info('Started generating new protocol')

        try:

            path = os.path.abspath(f".\\templates\\{filename}")

            self.app = win32com.client.Dispatch("Word.Application") 
            self.app.Visible = False

            self.doc = self.app.Documents.Open(path)

        except Exception as e:

            logger.error(f'Could not open Word-Application with the given error: "{e}"')

    def get_bookmarks(self) -> list:
        
        bookmarks = []

        # Loop through the bookmarks collection
        for bookmark in self.doc.Bookmarks:
            bookmarks.append(bookmark.Name)
        
        return bookmarks

    def close_application(self, reason: str):

        self.doc.Close()
        del self.doc

        self.app.Quit()
        del self.app   

        logger.info(f'Closed Application with following reason: {reason}')

class json_data_process():

    def __init__(self) -> None:

        self.last_check_time = time.time()
        self.new_files = list

    def run_periodically(self, interval, func, *args):
        def wrapper():
            while not stop_event.is_set():
                func(*args)
                time.sleep(interval)
        
        stop_event = threading.Event()
        thread = threading.Thread(target=wrapper)
        thread.start()
        return stop_event

    def get_latest_files(self, folder_path) -> list:

        files = os.listdir(folder_path)
        current_time = time.time()
        new_files = []

        for file_name in files:
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                # Get the last modification time of the file
                file_mod_time = os.path.getmtime(file_path)
                # Check if the file was modified after the last check time
                if file_mod_time > self.last_check_time:
                    new_files.append(file_name)

        self.last_check_time = current_time
        self.new_files = new_files

        logger.info('Checked for new Data')
        print(new_files)
        
        return 

    def read_JSON_Data(self, file, topic, parameter=None):

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

# def fill_Word_report(report_json):

#     if jdp.check_for_nio(report_json) == True:
#         jdp.save_JSON_Data(report_json, 'header', 'Ergebnis', 'Bestanden')
#     else:
#         jdp.save_JSON_Data(report_json, 'header', 'Ergebnis', 'Nicht Bestanden')

#     with open(report_json) as config:                  #Übergebene JSON File öffnen
#         config_data = json.load(config)         #Daten auslesen

#     for entry in config_data["header"]: 
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             print(key_list[point])
#             print(value_list[point])
#             doc.Bookmarks(key_list[point]).Range.InsertAfter(value_list[point])
#             if key_list[point] == "Datum" or key_list[point] == "Pruefer":
#                 doc.Bookmarks(key_list[point]+ '_2').Range.InsertAfter(value_list[point])

#     for entry in config_data["Result_Verteilung"]: 
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             if key_list[point] == 'Result_ID':
#                 next
#             else:
#                 print(key_list[point])
#                 print(value_list[point])
#                 doc.Bookmarks(key_list[point]).Range.InsertAfter(value_list[point])

#     for entry in config_data["Result_Multi"]: 
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             print(key_list[point])
#             print(value_list[point])
#             result_name = key_list[point].replace('Result_', 'Multi_')
#             if result_name == "Multi_Firmware":
#                 next
#             else:
#                 doc.Bookmarks(result_name).Range.InsertAfter(value_list[point])

#     for entry in config_data["Result_Leiste"]: 
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             print(key_list[point])
#             print(value_list[point])

#             if point == 0:
#                 ID = str(value_list[point])

#             result_name = key_list[point].replace('Result_', 'Leiste_'+ID+'_')
#             if "Firmware" in result_name:
#                 next
#             else:
#                 doc.Bookmarks(result_name).Range.InsertAfter(value_list[point])

#     config.close()

#     with open('config.json') as config:                  #Übergebene JSON File öffnen
#         config_data = json.load(config)         #Daten auslesen

#     for entry in config_data['Validate_Multi']:
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             print(key_list[point])
#             print(value_list[point])

#             if point == 0:
#                 ID = str(value_list[point])

#             result_name = 'Multi_' + key_list[point]
            
#             doc.Bookmarks(result_name).Range.InsertAfter(value_list[point])

#     for entry in config_data['Validate_Leiste']:
#         key_list = list(entry.keys())                                                                               #Namen der Register auslesen
#         value_list = list(entry.values())   

#         for point in range(0, len(key_list)):
#             print(key_list[point])
#             print(value_list[point])

#             if point == 0:
#                 ID = str(value_list[point])

#             result_name = 'Leiste_' + key_list[point]
#             doc.Bookmarks(result_name).Range.InsertAfter(value_list[point])
#             result_name = 'Leiste_' + key_list[point] + '_2'
#             doc.Bookmarks(result_name).Range.InsertAfter(value_list[point])

#     # Neuen Dateinamen generieren und unter diesem Dateinamen das geänderte
#     # Dokument abspeichern
#     base, ext = os.path.splitext(filename)
#     Artikelnummer = jdp.read_JSON_Data(report_json, "header", "Artikelnummer")
#     Seriennummer = str(jdp.read_JSON_Data(report_json, "header", "Seriennummer"))
#     new_filename = base + "_" + Artikelnummer + '_' + Seriennummer + ext
#     doc.Fields.Update()
#     doc.SaveAs(new_filename)
    

#     # Das Dokument und Word schließen und die Referenzen vernichten. WICHTIG!
#     doc.Close()
#     del doc

#     app.Quit()
#     del app

#     pdf_filename = new_filename.replace('.docx', '.pdf')
#     convert(new_filename, pdf_filename)

#     # Fertig
#     print("Fertig")


if __name__ == '__main__':

    word_process = create_report()
    json_process = json_data_process()
    init_logger()
    # print(word_process.check_filename('Report_JM.docx')) 
    # file = word_process.check_filename('Test.docx')
    # word_process.open_word_document(file)
    # print(word_process.get_bookmarks())
    # word_process.close_application('Test finished')

    path_to_test_data = json_process.read_JSON_Data('config.json', 'Config_Program', 'Path_to_data')
    interval = json_process.read_JSON_Data('config.json', 'Config_Program', 'Refresh_rate')
    repeated_event = json_process.run_periodically(interval, json_process.get_latest_files, path_to_test_data)

    while True:
        
        pass
