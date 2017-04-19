'''
Purpose: Update Tableau workbooks by issuing commands to:
         1. Save data extract as .tde file in it's proper location
         2. opening Tableau workbook using it's file location .exe
         3. Clicking refresh within Tableau in order to use the already created .tde file
         4. Saving newly refreshed workbook as a packaged workbook (.twbx)
         5. Emailing newly packaged workbook/link location to desired recipients (address_book)
         
Author: Nick Stark
        Written in Python 3.6
'''
import os
import pyautogui
import subprocess
import logging
import sys
import smtplib
from email.mime import multipart
from email.mime import text
from email.mime import base
from email import encoders

logging.basicConfig(level=logging.INFO, format='%(asctime)s.%(message)03d: %(message)s', datefmt='%H:%M:%S')
# logging.disable(logging.info) # UNCOMMENT TO BLOCK DEBUG LOG MESSAGES

screenRegion = pyautogui.size()
directory = "C:\\Users\\NStark\Documents\Python Scripts\ReportUpdater\\report-updater\\autoGuiTableau"
reportName = 'Scrap.twb'
# tableauReport = pyautogui.prompt('Please type in the full path of the Tableau report file (with .twb).')


def main():
    '''
    Runs the entire program. 
    :return: End result of program: Saved .twbx and an email sent with updated .twbx file attached
    '''
    logging.info('Program started. Press Ctrl-C to abort at any time.')
    logging.info('To pause mouse movement, move cursor to the very top left of the screen.')
    logging.info('Changing current working directory...')
    os.chdir(directory)     # CHANGES TO THE CORRECT DIRECTORY
    openTableau()
    getScreenRegion()
    navigateLogin()
    navigateDataMenu()
    saveAsPackagedWorkbook()
    sendEmail()


def openTableau():
    '''
    Open Tableau report within the current working directory (change directory variable to change CWD)
    :return: None, opens Tableau when called
    '''
    logging.info('Opening Tableau...')
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
    region = pyautogui.locateCenterOnScreen(imPath('data_source_tab.png'))
    if region is None:
        raise Exception('Could not find Tableau. Is it visible?')
    else:
        pyautogui.click(region)


def navigateLogin():
    '''
    Click "Data Source" tab, login prompt comes up, hit tab 5 times, enter password, hit enter
    :return: Nothing, it logs in to the database to refresh the extract
    '''
    logging.info('Logging in to database used in report...')
    pyautogui.press(['tab', 'tab', 'tab', 'tab', 'tab'], interval=0.5)
    pyautogui.typewrite('Mayzooper11')
    pyautogui.press('enter')


def navigateDataMenu():
    '''
    Click "Automatically Update" to refresh the data extract (.tde file).
    Finds coordinates for "Automatically Update" button, then clicks on that button.
    '''
    logging.info('Refreshing data...')
    updateButton = pyautogui.locateCenterOnScreen(imPath('automatically_update.png'))
    if updateButton is None:
        raise Exception('Can not find Auto Update button, make sure Tableau is in focus.')
    pyautogui.click(updateButton)


def saveAsPackagedWorkbook():
    '''
    Click file - Save As - Change file type to .twbx, type in Scrap.twbx, click okay
    '''
    # TODO: Finish function
    logging.info('Saving file as a packaged workbook (.twbx), readable by Tableau Reader...')
    pyautogui.hotkey('alt', 'f')
    pyautogui.hotkey('a')

# TODO: Email with attached .twbx file
def sendEmail():
    '''
    Sends an email with the updated .twbx file, able to be opened in Tableau Reader
    '''

    logging.info('Building email message...')

    # CREDENTIALS
    fromaddr = 'nstark@bradfordwhite.com'               # YOUR EMAIL
    toaddr = 'nstark@bradfordwhite.com'                 # EMAIL ADDRESS YOU SEND TO
    # outlook_pwd = ''                                  # CAN HARDCODE THIS OR USE sys.argv[] AND YOU'LL BE PROMPTED

    # MESSAGE INFORMATION
    msg = multipart()

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = 'TEST'                            # SUBJECT OF EMAIL

    body = "TESTING THE BODY OF THE EMAIL"             # TEXT YOU WANT TO SEND

    msg.attach(text(body, 'plain'))

    # EMAIL ATTACHMENT INFO
    filename = 'Scrap.twbx'
    attachment = open(r'C:\Users\NStark\Desktop\Nick Stark\Tableau\Scrap\Scrap.twbx', 'rb')

    # ADD ATTACHMENT TO HEADER OF EMAIL
    part = base('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    logging.info('Sending email to: ' + toaddr + '.\nFrom: ' + fromaddr)

     # CONNECT AND SEND MAIL
    smtpserver = smtplib.SMTP('w2k8e-exch.grr.bradfordwhite.com', 587)          # PORT MAY BE WRONG
    smtpserver.ehlo()
    smtpserver.starttls()
    # smtpserver.login(fromaddr, outlook_pwd)                                   # UNCOMMENT FOR HARDCODED PWD STORED
    smtpserver.login(fromaddr, sys.argv[1])
    sendMailStatus = smtpserver.sendmail(from_addr=fromaddr, to_addrs=toaddr, msg=msg)


    # NOTIFY IF ERROR
    if sendMailStatus != {}:
        print('There was a problem sending an email to %s: %s' % (toaddr, sendMailStatus))


# CALL MAIN
if __name__ == '__main__':
    main()

# EXIT PROGRAM UPON COMPLETION OF MAIN FUNCTION
sys.exit("Exiting program.")