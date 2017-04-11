'''
Purpose: Update Tableau workbooks by issuing commands to:
         1. Save data extract as .tde file in it's proper location
         2. opening Tableau workbook using it's file location .exe
         3. Clicking refresh within Tableau in order to use the already created .tde file
         4. Saving newly refreshed workbook as a packaged workbook (.twbx)
         5. Emailing newly packaged workbook/link location to desired recipients
         
Author: Nick Stark
        Written in Python 3.6
'''
import os
import pyautogui
import subprocess
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s.%(message)03d: %(message)s', datefmt='%H:%M:%S')
# logging.disable(logging.info) # UNCOMMENT TO BLOCK DEBUG LOG MESSAGES

screenRegion = pyautogui.size()
directory = "C:\\Users\\NStark\Documents\Python Scripts\ReportUpdater\\report-updater\\autoGuiTableau"
reportName = 'Scrap.twb'
# tableauReport = pyautogui.prompt('Please type in the full path of the Tableau report file (with .twb).')

# os.chdir(r'C:\Users\NStark\Documents\Python Scripts\ReportUpdater\report-updater\autoGuiTableau')


def main():
    '''
    Runs the entire program. 
    :return: End result of program: Saved .twbx and an email sent with updated .twbx
    '''
    logging.info('Program started. Press Ctrl-C to abort at any time.')
    logging.info('To pause mouse movement, move cursor to the very top left of the screen.')
    logging.info('Changing current working directory...')
    os.chdir(directory)     # CHANGES TO THE CORRECT DIRECTORY
    logging.info('Opening Tableau...')
    openTableau()
    logging.info('Finding Tableau...')
    getScreenRegion()
    logging.info('Logging in to database used in report...')
    navigateLogin()
    logging.info('Refreshing data and saving workbook as packaged workbook (.twbx)...')
    navigateDataMenu()
    saveAsPackagedWorkbook()


def openTableau():
    '''
    Open Tableau report within the current working directory (change directory variable to change CWD)
    :return: None, opens Tableau when called
    '''
    subprocess.Popen(reportName, shell=True)


def imPath(filename):
    '''
    A shortcut for joining the 'images'/'' file path, since it is used so often
    :param filename: Name of the image file (.png).
    :return: Filename with 'images/' prepended.
    '''
    return os.path.join('images', filename)


def getScreenRegion():
    '''
    Gets the region for Tableau
    :return: Region in which Tableau is currently open on the screen
    '''
    global screenRegion

    # FIND TABLEAU ON USERS SCREEN
    logging.info('Finding Tableau on the screen...')
    region = pyautogui.locateOnScreen(imPath('data_source_tab.png'))
    if region is None:
        raise Exception('Could not find Tableau. Is it visible?')
    else:
        pyautogui.click(region)


def navigateLogin():
    '''
    Click "Data Source" tab, login prompt comes up, hit tab 5 times, enter password, hit enter
    :return: Nothing, it logs in to the database to refresh the extract
    '''
    pyautogui.press(['tab', 'tab', 'tab', 'tab', 'tab'])
    pyautogui.typewrite('Mayzooper11')
    pyautogui.press('enter')


def navigateDataMenu():
    # TODO: Finish function


def saveAsPackagedWorkbook():
    # TODO: Finish function

if __name__ == '__main__':
    main()

# TODO: Email with attached .twbx file