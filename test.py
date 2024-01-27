import os
import sys
import ctypes
import logging
import win32com.client
import subprocess
import urllib.request
import psutil

# Set up logging
logging.basicConfig(filename='script.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Check if the script is running with administrative privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Request administrator privileges if not already elevated
def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

# Check and prompt for admin rights
if not is_admin():
    print("This script requires administrative privileges. Restarting with elevated permissions...")
    logging.warning('The user is not running the program with administrative privileges.')
    run_as_admin()
    sys.exit()

def general_optimization():
    command = 'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
    subprocess.call(command, shell=True)
    logging.info('Ran "powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"')

def clean_temp_files():
    logging.info('Cleaning temporary files...')
    os.system("cleanmgr /sagerun:1")
    logging.info('Ran "cleanmgr /sagerun:1"')
    logging.info('Temporary files cleaned.')

def update_defender():
    defender_update_cmd = r'"C:\Program Files\Windows Defender\MpCmdRun.exe" -SignatureUpdate'
    subprocess.call(defender_update_cmd)
    logging.info('Ran windows defender update')

def update_windows():
    update_session = win32com.client.Dispatch("Microsoft.Update.Session")
    update_searcher = update_session.CreateUpdateSearcher()
    update_searcher.Online = True

    logging.info("Searching for available updates...")
    search_result = update_searcher.Search("IsInstalled=0")

    if search_result.Updates.Count == 0:
        logging.info("No updates are available.")
    else:
        logging.info(f"{search_result.Updates.Count} update(s) are available.")
        logging.info("Downloading and installing updates...")

        update_collection = win32com.client.Dispatch("Microsoft.Update.UpdateColl")
        for update in search_result.Updates:
            update_collection.Add(update)

        downloader = update_session.CreateUpdateDownloader()
        downloader.Updates = update_collection
        downloader.Download()

        installer = update_session.CreateUpdateInstaller()
        installer.Updates = update_collection
        installation_result = installer.Install()

        if installation_result.ResultCode == 2:
            logging.info("All updates were successfully installed.")
        else:
            logging.error(f'Failed to install updates. Error code: {installation_result.ResultCode}')

def optimize_drive():
    logging.info('Optimizing drives...')

    drives = [drive.device for drive in psutil.disk_partitions()]

    for drive in drives:
        logging.info(f'Optimizing {drive}...')
        os.system(f'defrag {drive} /O')
        logging.info(f'{drive} optimized.')

    logging.info('All drives optimized.')

def check_errors():
    logging.info('Checking for errors...')
    logging.info("Running SFC")
    subprocess.run(["sfc", "/scannow"])
    logging.info("Running chkdsk")
    subprocess.run(["chkdsk", "/f"])
    logging.info("Running dism")
    subprocess.run(["dism", "/online", "/cleanup-image", "/restorehealth"])
    logging.info('Errors checked.')

def scan_for_viruses():
    logging.info('Quick scanning for viruses...')
    os.system("powershell.exe Start-MpScan -ScanType QuickScan")
    logging.info('Quick scanning for viruses completed.')
    logging.info('Virus scan completed.')

def install_opera_gx():
    logging.info('Downloading Opera GX...')
    url = "https://net.geo.opera.com/opera_gx/stable/windows?utm_tryagain=yes&utm_source=google&utm_medium=ose&utm_campaign=(none)&http_referrer=https%3A%2F%2Fwww.google.com%2F&utm_site=opera_com&&utm_lastpage=opera.com/"
    file_name = "Opera_GX_Setup.exe"
    urllib.request.urlretrieve(url, file_name)
    subprocess.call([file_name, "/silent", "/install"])
    os.remove(file_name)
    logging.info('Opera GX has been installed successfully!')

def restart():
    input("Press enter to restart")
    logging.info('Restarting computer')
    os.system("shutdown /r /t 20")

# Execute the script
logging.info('Starting script...')
check_errors()
update_defender()
update_windows()
scan_for_viruses()
clean_temp_files()
general_optimization()
optimize_drive()
check_errors()
answer = input("Do you want to install Opera GX (it takes fewer resources)? Y/N: ")
if answer.lower() in ["y", "yes"]:
    install_opera_gx()
else:
    print("Okay")
restart()
logging.info('Script completed.')
