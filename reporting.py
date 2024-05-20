import subprocess
# import psutil
import time
# import socket
import uno

import json_data_process as jdp

def open_libreoffice():

    libreoffice_path = jdp.read_JSON_Data("config.json", "Config_Program", "Open_Office_Path")

    subprocess.Popen([libreoffice_path, "--headless", "--accept=\"socket,port=2002;urp;\""])

# def close_libreoffice():

#     # Look for 'soffice' process
#     for process in psutil.process_iter(['name']):
#         if process.info['name'] == 'soffice.bin':
#             process.terminate()
#             break

def connect_to_office():
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    try:
        office_context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        return office_context
    except:
        print("Failed to connect to LibreOffice/OpenOffice. Make sure the office suite is started with listening mode enabled.")
        exit(1)

def open_document(office_context, file_url):
    desktop = office_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", office_context)
    try:
        document = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
        return document
    except Exception as e:
        print("Error opening document:", e)
        exit(1)

def insert_text_to_textmark(doc, mark_name, text):
    bookmarks = doc.getBookmarks()
    if mark_name in bookmarks.getElementNames():
        bookmark = bookmarks.getByName(mark_name)
        text_range = bookmark.getAnchor()
        text_cursor = text_range.getText().createTextCursorByRange(text_range)
        text_cursor.setString(text)
    else:
        print("Bookmark not found.")

def save_and_close(document, file_url):
    document.storeAsURL(file_url, ())
    document.dispose()  # Dispose of the document

def terminate_libreoffice(office_context):
    desktop = office_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", office_context)
    desktop.terminate()  # Terminate the LibreOffice application


if __name__ == "__main__":

    open_libreoffice()

    time.sleep(10)

    # file_path = "Report_JM.docx"  # Change this to your file's path
    office_context = connect_to_office()
    # document = open_document(office_context, file_path)
    # bookmark_name = "Seriennummer"
    # text_to_insert = "123123123"
    # insert_text_to_textmark(document, bookmark_name, text_to_insert)
    # save_and_close(document, file_path)

    # time.sleep(10)

    # close_libreoffice()
    terminate_libreoffice(office_context)


