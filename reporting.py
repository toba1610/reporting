import subprocess
import psutil
import time


def open_libreoffice():
    libreoffice_path = "D:\LibreOffice\program\swriter.exe"

    subprocess.Popen([libreoffice_path])


def close_libreoffice():
    # Look for 'soffice' process
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'soffice.bin':
            process.terminate()  # Terminate the process
            break


time.sleep(10)

close_libreoffice()


