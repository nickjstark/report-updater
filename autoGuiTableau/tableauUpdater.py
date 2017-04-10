'''
Purpose: Update Tableau workbooks by issuing commands to:
         1. Save data extract as .tde file in it's proper location
         2. opening Tableau workbook using it's file location .exe
         3. Clicking refresh within Tableau in order to use the already created .tde file
         4. Saving newly refreshed workbook as a packaged workbook (.twbx)
         5. Emailing newly packaged workbook/link location to desired recipients
         
Author: Nick Stark
        Written in Python 3.6.1
'''
import os
import pyautogui
import subprocess
import sys

# CHANGE DIRECTORY AND LAUNCH TABLEAU WORKBOOK

os.chdir(r'C:\Users\NStark\Desktop\Nick Stark\Tableau\Scrap')

# subprocess.Popen(r'C:\Users\NStark\Desktop\Nick Stark\Tableau\Scrap\Scrap.twb', shell=True)

output = subprocess.check_output("dir", shell=True)
print(output)