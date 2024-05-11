import subprocess
import psutil
import time
import json_data_process as jdp


def open_libreoffice():
    libreoffice_path = jdp.read_JSON_Data("config.json", "Config_Program", "Open_Office_Path")

    # subprocess.Popen([libreoffice_path])
    subprocess.Popen([libreoffice_path, "--headless", "--accept=\"socket,port=2002;urp;\""])


def close_libreoffice():
    # Look for 'soffice' process
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'soffice.bin':
            process.terminate()  # Terminate the process
            break

open_libreoffice()

time.sleep(10)

close_libreoffice()


