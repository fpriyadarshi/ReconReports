from selenium import webdriver
from chromedriver_py import binary_path
from selenium.webdriver.support.ui import WebDriverWait
import os
import sqlite3
from simple_salesforce import Salesforce, SFType, SalesforceLogin
from getgauge.python import Messages, Screenshots, after_step
from datetime import datetime
import chromedriver_autoinstaller
import geckodriver_autoinstaller
from pathlib import Path
import logging
from datetime import date
from pathlib import Path
import pdb
from selenium.common.exceptions import ElementNotVisibleException, ElementNotSelectableException

driver = None
driverWait = None
sf = None
driver4Adobe = None
driver4AdobeWait = None
dbConn = None
dbCursor = None
logger = None

def Initialize():
    # executable_path = {'executable_path': binary_path}
    cwd = os.path.dirname(os.path.realpath(__file__))
    print(cwd)
    # CHORME_PATH = cwd + "\\" + "chromedriver.exe"
    chromedriver_autoinstaller.install()
    global driver
    global driverWait
    chrome_options = webdriver.ChromeOptions()
    
    chrome_options.add_argument('--disable-application-cache')
    prefs = {"download.default_directory": cwd,
             'download.directory_upgrade': True}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path=binary_path,chrome_options=chrome_options)
    # driverWait = WebDriverWait(driver, 120)
    driverWait = WebDriverWait(driver, 10, poll_frequency=1, ignored_exceptions=[
                               ElementNotVisibleException, ElementNotSelectableException])


    return driver, driverWait


def Initialize_Window_For_Adobe():
    chromedriver_autoinstaller.install()
    global driver4Adobe
    global driver4AdobeWait
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option(
        "excludeSwitches", ["enable-logging"])
    driver4Adobe = webdriver.Chrome(chrome_options=chrome_options)
    driver4AdobeWait = WebDriverWait(driver4Adobe, 720)
    return driver4Adobe, driver4AdobeWait

def Initialize_Logger():
    global logger
    try:
        today = date.today()
        oppDateToday = today.strftime("%Y%m%d")
        logger = None
        root = Path(__file__).parents[1]
        loggerFileName = str(root) + "\\logs\\" + f"{os.getenv('ORG_TYPE')}_logs_{oppDateToday}.log"

        logger = logging.getLogger(os.getenv('ORG_TYPE'))
        
        if not logger.handlers:    
            logger.setLevel(logging.INFO)
            format = logging.Formatter(
                '%(asctime)s >> %(name)s >> %(levelname)s >> %(message)s')

            fileHandler = logging.FileHandler(loggerFileName)
            fileHandler.setLevel(logging.INFO)
            fileHandler.setFormatter(format)

            consoleHandler = logging.StreamHandler()
            consoleHandler.setLevel(logging.INFO)
            consoleHandler.setFormatter(format)

            logger.addHandler(fileHandler)
            logger.addHandler(consoleHandler)
            # pdb.set_trace()
        return logger

    except Exception:
            logger.exception("exception occurred", exc_info=True)


def Initialize_SalesForce_Instance():
    global sf
    # session_id, instance = SalesforceLogin(
    #     username=os.getenv("USER_ID"), password=os.getenv("USER_PASSWORD"), security_token=os.getenv("USER_SECURITY_TOKEN"), domain='test')

    session_id, instance = SalesforceLogin(
        username=os.getenv("USER_ID"), password=os.getenv("USER_PASSWORD"), security_token=os.getenv("USER_SECURITY_TOKEN"), domain='test')

    print(session_id, "\n", instance)
    Messages.write_message(str(session_id) + " : " + str(instance))
    sf = Salesforce(instance=instance, session_id=session_id)
    return sf


def Initialize_Database_Instance():
    global dbConn
    global dbCursor
    root = Path(__file__).parents[1]
    dbFilePath = str(root) + "\\Data\\" + os.getenv('DB_NAME')
    print(dbFilePath)
    dbConn = sqlite3.connect(dbFilePath)
    print("Opened database successfully", dbConn)
    Messages.write_message(f"Opened DB {dbConn} Successfully")
    dbCursor = dbConn.cursor()
    print("Cursor Object: ", dbCursor)
    Messages.write_message(f"Cursor Object {dbCursor}")
    # return SQLite3Connection.cur

    return dbConn, dbCursor


def CloseDriver():
    driver.quit()


@after_step
def after_step_hook(context):
    if context.step.is_failing == True:
        Messages.write_message(context.step.text)
        # Messages.write_message(context.step.message)
        Screenshots.capture_screenshot()
