import urllib.request
import os
import subprocess

# Download URL for the latest version of Opera GX for Windows
url = "https://net.geo.opera.com/opera_gx/stable/windows?utm_tryagain=yes&utm_source=google&utm_medium=ose&utm_campaign=(none)&http_referrer=https%3A%2F%2Fwww.google.com%2F&utm_site=opera_com&&utm_lastpage=opera.com/"

# File name for the downloaded installer
file_name = "Opera_GX_Setup.exe"

# Download the installer
urllib.request.urlretrieve(url, file_name)

# Install Opera GX
subprocess.call([file_name, "/silent", "/install"])

# Clean up the installer file
os.remove(file_name)

print("Opera GX has been installed successfully!")
