import os
import sys
import ctypes
import logging
import win32com.client
import subprocess
import urllib.request
import psutil
logging.basicConfig(filename='script.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')
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

# Set up logging



def general_optimization():
    command = 'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
    subprocess.call(command, shell=True)
    logging.info('Ran "powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"')


def clean_temp_files():
    print('Cleaning temporary files...')
    os.system("cleanmgr /sagerun:1")
    logging.info('Ran "cleanmgr /sagerun:1"')
    print('Temporary files cleaned.')


def update_defender():
    defender_update_cmd = r'"C:\Program Files\Windows Defender\MpCmdRun.exe" -SignatureUpdate'
    subprocess.call(defender_update_cmd)
    logging.info('Ran windows defender update')


def update_windows():
    
    update_session = win32com.client.Dispatch("Microsoft.Update.Session")
    update_searcher = update_session.CreateUpdateSearcher()
    update_searcher.Online = True

    print("Searching for available updates...")
    search_result = update_searcher.Search("IsInstalled=0")

    if search_result.Updates.Count == 0:
        print("No updates are available.")
    else:
        print(f"{search_result.Updates.Count} update(s) are available.")
        print("Downloading and installing updates...")

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
            print("All updates were successfully installed.")
            logging.info('Ran windows update successfully')
        else:
            print(f"Failed to install updates. Error code: {installation_result.ResultCode}")
            logging.error(f'Failed to install updates. Error code: {installation_result.ResultCode}')


def optimize_drive():
    print('Optimizing drives...')

    drives = [drive.device for drive in psutil.disk_partitions()]

    for drive in drives:
        print(f'Optimizing {drive}...')
        os.system(f'defrag {drive} /O')
        print(f'{drive} optimized.')

    print('All drives optimized.')
    logging.info('All drives optimized.')


def check_errors():
    print('Checking for errors...')
    print("Running SFC")
    logging.info('Running SFC')
    subprocess.run(["sfc", "/scannow"])
    print("\n")
    print("Running chkdsk")
    logging.info('Running chkdsk')
    subprocess.run(["chkdsk /f"])
    print("\n")
    print("Running dism")
    logging.info('Running dism')
    subprocess.run(["dism", "/online", "/cleanup-image", "/restorehealth"])
    print('Errors checked.')
    logging.info('Errors checked.')
    print("\n")


def scan_for_viruses():
    print('Quick scanning for viruses...')
    os.system("powershell.exe Start-MpScan -ScanType QuickScan")
    logging.info('Quick scanning for viruses completed.')
    print('Virus scan completed.')


def install_opera_gx():
    url = "https://net.geo.opera.com/opera_gx/stable/windows?utm_tryagain=yes&utm_source=google&utm_medium=ose&utm_campaign=(none)&http_referrer=https%3A%2F%2Fwww.google.com%2F&utm_site=opera_com&&utm_lastpage=opera.com/"
    file_name = "Opera_GX_Setup.exe"
    urllib.request.urlretrieve(url, file_name)
    subprocess.call([file_name, "/silent", "/install"])
    os.remove(file_name)
    print("Opera GX has been installed successfully!")
    logging.info('Opera GX has been installed successfully!')


def restart():
    input("Press enter to restart")
    logging.info('Restarting computer')
    os.system("shutdown /r /t 20")


# Execute the script
print('Starting script...')
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
print('Script completed.')
