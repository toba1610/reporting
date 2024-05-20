import win32com
import win32com.client
import json
import json_data_process as jdp
import os
import logging
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
    init_logger()
    # print(word_process.check_filename('Report_JM.docx')) 
    file = word_process.check_filename('Test.docx')
    word_process.open_word_document(file)
    print(word_process.get_bookmarks())