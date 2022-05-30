import os
import re
import time
import logging
import zipfile
import datetime
import requests
import traceback
import pandas as pd
from pathlib import Path
from time import strftime
from selenium import webdriver
from argparse import ArgumentParser
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from supabase import create_client, Client
from googleapiclient.discovery import build
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

RESUME_LINKS = {}
TAG_FILE = datetime.datetime.now().isoformat()
TAG_DIR = datetime.datetime.now().date().isoformat()
LOG_FILE_NAME = "{}{}LOG_{}.txt"
LOG_PATH = os.getcwd()+os.sep+TAG_DIR
if not os.path.exists(LOG_PATH):
    os.makedirs(LOG_PATH)
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(filename=LOG_FILE_NAME.format(LOG_PATH, os.sep, TAG_FILE),
                    format='%(asctime)s - %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)

def parse_arguments():
    parser = ArgumentParser(description='\n',
    usage='''
        ***************************************
            This script is an automated solution for saving candidates for review in Instahyre
            python3 instaTest.py -c <Due category> --help
        **************************************\n''')
    parser.add_argument('-id', '--jobId',\
        help='Opening URL to check for candidates. ', required=True)
    parser.add_argument('-un', '--username',\
        help='User name to login. ', required=True)
    parser.add_argument('-pwd', '--password',\
        help='Password to login. ', required=True)
    parser.add_argument('-nc', '--noOfCandidates', type=int, const=1, nargs='?',\
        help='Maximum no of candidates to save. ', required=False)
    return parser.parse_args()

def googleAuth():
    gauth = GoogleAuth()
        # Try to load saved client credentials
    gauth.LoadCredentialsFile("mycreds.txt")
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.GetFlow()
        gauth.flow.params.update({'access_type': 'offline'})
        gauth.flow.params.update({'approval_prompt': 'force'})
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    return gauth

def login(username,password):
    """
    This code block will help in logging in to Instahyre.
    """
    try:
        driver = webdriver.Chrome(ChromeDriverManager().install())
        url = "https://www.instahyre.com/employer/candidates/"
        driver.get(url)
        driver.implicitly_wait(15)
        loginBox = driver.find_element(by=By.XPATH, value='//*[@id ="email"]')
        loginBox.send_keys(username)
        passwordBox = driver.find_element(by=By.XPATH, value='//*[@id ="password"]')
        passwordBox.send_keys(password)
        loginButton = driver.find_element(by=By.XPATH, value='//*[@id ="login-form"]')
        loginButton.click()
        return driver
    except Exception as exc:
        logger.info("Login Unsuucessful .\n Exception Raised : \n", exc)
        logger.info("Full Traceback for debugging: \n", traceback.format_exc())

def saveForReview(driver, jobId, noOfCandidates):
    """
    This code block will help in saving the candidates for review in Instahyre.
    """
    time.sleep(5)
    url = "https://www.instahyre.com/employer/candidates/" + jobId +"/0/"
    driver.get(url)
    time.sleep(5)
    companyNamesWithProductMatch = []
    companyNamesWithNoMatch = []
    companyNamesWithServiceMatch = []
    try:
        i = 1
        while i <= noOfCandidates:
            refreshPage = False
            result = driver.find_elements(By.CLASS_NAME, "employer-applications")
            time.sleep(10)
            candidates = WebDriverWait(result[0], 100).until(EC.presence_of_all_elements_located((By.XPATH, "//*[@ng-repeat='candidate in candidates']")))
            for candidate in candidates:
                companyName = candidate.find_elements(By.CLASS_NAME, "candidate-company-designation")[0].text
                if companyName:
                    if " at " in companyName:
                        companyName = companyName.split(" at ")[1].strip()
                        parms = {"company_name" : companyName, "api_key" : "9874"}
                        response = requests.post("https://app.pragti.in/api/employer/applicants/get_company_info", data = parms)
                        if response.status_code == 200:
                            companyType = response.json().get('data').get('type')
                        else:
                            companyType = " "
                            logger.info("Couldn't get information from the company database.")
                        if companyType == "product":
                            companyNamesWithProductMatch.append(companyName)
                            saveForReviewButton = candidate.find_elements(By.CLASS_NAME,"button-hide-save")[0]
                            driver.execute_script("arguments[0].click();", saveForReviewButton)
                            refreshPage = True
                            i += 1
                            break
                        elif companyType == "no match":
                            companyNamesWithNoMatch.append(companyName)
                        else:
                            companyNamesWithServiceMatch.append(companyName)
                    else:
                        continue
                else:
                    logger.info("Couldn't find company name")
            if refreshPage:
                continue
            paginationButton = driver.find_elements(By.CLASS_NAME, "pagination")
            if paginationButton:
                tags = paginationButton[0].find_elements(By.TAG_NAME, "li")
                tag = [tag for tag in tags if tag.text == "Next »"]
                if tag:
                    tag = tag[0]
                else:
                    break
                if tag.is_enabled():
                    driver.execute_script("arguments[0].click();", tag)
            else:
                logger.info("*********************Could only Found {} candidates*********************".format(i))
                break
        logger.info("Company Names with product Match : \n {}".format(companyNamesWithProductMatch))
        logger.info("Company Names no product Match : \n {}".format(companyNamesWithNoMatch))
        logger.info("Company Names no service Match : \n {}".format(companyNamesWithServiceMatch))
    except Exception as exc:
        logger.info("Excepton raised during execution ", exc)
        logger.info("Full Traceback for debugging: \n", traceback.format_exc())

def downloadResume(driver, jobId):
    """
    This code block will help in downloding the candidates resume from instahyre.
    """
    try:
        time.sleep(5)
        url = "https://www.instahyre.com/employer/candidates/" + jobId +"/3/"
        driver.get(url)
        time.sleep(5)
        selectedCandidates = 0
        result = driver.find_elements(By.CLASS_NAME, "employer-applications")
        candidates = WebDriverWait(result[0], 100).until(EC.presence_of_all_elements_located((By.XPATH, "//*[@ng-repeat='candidate in candidates']")))
        for candidate in candidates:
            checkboxSelected = candidate.find_elements(By.CLASS_NAME, "fa-check-circle")
            if checkboxSelected:
                data = checkboxSelected[0].get_attribute("data-html")
                if data != "true":
                    candidate.find_elements(By.CLASS_NAME, "ng-pristine")[0].click()
                    selectedCandidates += 1
            else:
                candidate.find_elements(By.CLASS_NAME, "ng-pristine")[0].click()
                selectedCandidates += 1

        if selectedCandidates > 0:
            downloadButton = driver.find_elements(By.CLASS_NAME, "button-bulk-download-resume")[0]
            driver.execute_script("arguments[0].click();", downloadButton)
            zipFileButton = driver.find_elements(By.ID, "download-zip")[0]
            driver.execute_script("arguments[0].click();",zipFileButton)
            excelButton = driver.find_elements(By.ID, "download-excel")[0]
            driver.execute_script("arguments[0].click();",excelButton)
            downloadResume = driver.find_elements(By.CLASS_NAME, "download-resume-action")[0]
            downloadButton = downloadResume.find_elements(By.CLASS_NAME, "btn-success")[0]
            driver.execute_script("arguments[0].click();",downloadButton)
            return True
        else:
            logger.info("No canddiates to download resume ")
            return False
    except Exception as exc:
        logger.info("Method downloadResume failed . Error : \n {}".format(exc))

def uploadResumeToGoogleDrive(path):
    """
    This code block will help in uploading the candidates resume to google drive.
    """
    try:
        gauth = googleAuth()
        drive = GoogleDrive(gauth)
        gauth.LocalWebserverAuth()
        zipFiles = [os.path.join(path, x) for x in os.listdir(path) if x.endswith(".pdf")]
        for files in zipFiles:
            gfile = drive.CreateFile({'parents': [{'id': "1bxfWE5ZtScGmVPXcsQX9idWXC_Am7tti"}]})
            gfile.SetContentFile(files)
            gfile["title"] = '{}_{}'.format(files.split("/")[-1], datetime.datetime.now().strftime("%Y-%m-%d-%H:%M:%S"))
            gfile.Upload()
        uploadedFiles = drive.ListFile({"q": "'" + "1bxfWE5ZtScGmVPXcsQX9idWXC_Am7tti" + "' in parents and trashed=False and mimeType!='application/vnd.google-apps.folder'"}).GetList()
        for files in uploadedFiles:
            RESUME_LINKS[files["title"]] = files["alternateLink"]
    except Exception as exc:
        logger.info("Method uploadResumeToGoogleDrive failed . Error : \n {}".format(exc))
        logger.info("Full Traceback for debugging: \n", traceback.format_exc())

def uploadCandidateToDatabase():
    """
    This code block will help in uploading the candidates details to the database.
    """
    try:
        api_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiIsImlhdCI6MTYzNDI5NDc5NiwiZXhwIjoxOTQ5ODcwNzk2fQ.iN5gN3TRvKibAhZwlO20zBJRo8JoqSRRxk6ooeNOKtE"
        downloads_path = str(Path.home() / "Downloads")
        zipFiles = [os.path.join(downloads_path, x) for x in os.listdir(downloads_path) if x.endswith(".zip")]
        newest = max(zipFiles , key = os.path.getctime)
        today = datetime.datetime.now().isoformat()
        dirName = os.path.join(downloads_path,"resume_{}".format(today))
        if not os.path.exists(dirName):
            os.mkdir(dirName)
        with zipfile.ZipFile(newest,"r") as zip_ref:
            zip_ref.extractall(os.path.join(dirName))
        files = os.listdir(dirName)
        uploadResumeToGoogleDrive(dirName)
        filesXls = [f for f in files if f[-4:] == 'xlsx']
        fullExcelFilePath = os.path.join(dirName, filesXls[0])
        df = pd.read_excel(fullExcelFilePath, sheet_name="Sheet1")
        companyJobDetails = df.iloc[[0]].values[0][1]
        df.columns = df.iloc[2]
        df = df.iloc[3:]
        df.columns = df.columns.str.replace(' ','_')
        listOfCandidates = []
        for index, row in df.iterrows():
            candidate = {}
            # Code Block to get the resume link for the candidate
            for key in RESUME_LINKS:
                if key.startswith(row.Candidate_Name.replace(' ', '_')):
                    candidate["resume_link"] = RESUME_LINKS[key]
                    break
            # Code Block to get company Id
            url = "https://ocsuosnatmmlnkxgybgh.supabase.co"
            headers = {'Content-Type': 'application/json','apikey':api_key}
            key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiIsImlhdCI6MTYzNDI5NDc5NiwiZXhwIjoxOTQ5ODcwNzk2fQ.iN5gN3TRvKibAhZwlO20zBJRo8JoqSRRxk6ooeNOKtE"
            supabase: Client = create_client(url, api_key)
            jd_data = supabase.table("company").select('id').eq('name', row.Current_Employer).execute()
            if len(jd_data.data) > 0:
                candidate["current_company"] = jd_data.data[0].get("id")
            else:
                candidate["current_company"] = None
                logger.info("Could not find company {} in database".format(row.Current_Employer))
            candidate["skill"] = row.Key_Skills
            candidate["name"] = row.Candidate_Name
            candidate["mobile"] = row.Phone_Number
            candidate["total_exp"] = int(row.Work_Experience.split(" ")[0])
            candidate["current_ctc"] = int(float(row.Current_Salary.split(" ")[0]))
            candidate["job_prefer_location"] = row.Preferred_Locations
            candidate["email"] = row.Email_Address
            if row.Notice_Period.lower()== "immediately":
                candidate["notice_period_days"] = 0
            else:
                candidate["notice_period_days"] = int(float(re.findall(r'\d+',row.Notice_Period.split(" ")[0])[0]))
            candidate["designation"] = row.Current_Designation
            candidate["notice_period_status"] = None
            candidate["source"] = "Instahyre"
            candidate["college_name_highest_degree"] = None
            candidate["job_search_profile"] = None
            candidate["earliest_joining_date"] = None
            candidate["earliest_joining_days"] = None
            candidate["job_search_status"] = None
            candidate["holding_ctc"] = None
            candidate["verified"] = None
            candidate["candidate_status"] = None
            candidate["remark"] = None
            candidate["assigned_to"] = None
            candidate["verified_at"] = None
            candidate["holding_offer"] = None
            candidate["job_prefer_type"] = None
            candidate["min_expected_ctc"] = None
            listOfCandidates.append(candidate)
        logger.info("Adding all tthe beelow candidates to the database : \n {}".format(listOfCandidates))
        print("Adding all tthe beelow candidates to the database : \n {}".format(listOfCandidates))
        url = 'https://ocsuosnatmmlnkxgybgh.supabase.co/rest/v1/candidates'
        headers = {'Content-Type': 'application/json','apikey':api_key}
        response = requests.post(url,headers=headers,json=listOfCandidates)
        if response.status_code == 201:
            logger.info("Successfully added candidates to the database")
        else:
            logger.info("Could not add candidate details in database . Error : \n {}".format(response.json()))
    except Exception as exc:
        logger.info("Method uploadCandidateToDatabase failed . Error : \n {}".format(exc))
        logger.info("Full Traceback for debugging: \n", traceback.format_exc())

def slackNotification(filename, filepath):
    my_file={'file':open(filepath, 'rb')}
    channel = "sourcing-alerts"
    token = "xoxb-1645315758097-3573159057125-shnhyhuw7l9i1FogwrOwi3wQ"
    payload={
        "filename": filename, 
        "initial_comment": "Instahyre sourcing details.",
        "token":token, 
        "channels":[channel],
        }
    response=requests.post("https://slack.com/api/files.upload", params=payload, files=my_file)
    if response.status_code != 200:
        logger.info("Could not share log files in slack. Error: \n", respose.json())

if __name__ == '__main__':
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    # here enter the id of your google sheet
    SAMPLE_SPREADSHEET_ID_input = '1aQq5OIF6zFfi4aW2shRFUPgtwbaIJwPfqqcRu9QGXSg'
    SAMPLE_RANGE_NAME = 'A1:AA1000'
    gauth = googleAuth()
    service = build('sheets', 'v4', credentials=gauth.credentials)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])
    required_data = values_input[1:]
    username="tanmay@pragti.in"
    password="pragti@432"
    driver = ''
    driver = login(username=username,password=password)
    JobIds = []
    try:
        for data in required_data:
            jobId = data[0]
            JobIds.append(jobId)
            noOfCandidates = int(data[1])
            logger.info("*********************************************************************************************")
            logger.info("Starting to execute Job Id {}".format(jobId))
            logger.info("*********************************************************************************************")
            saveForReview(driver=driver, jobId=jobId, noOfCandidates=noOfCandidates)
            if driver:
                returnResult = downloadResume(driver=driver,jobId=jobId)
                if returnResult:
                    time.sleep(10)
                    uploadCandidateToDatabase()
                else:
                    logger.info("No candidates downloaded.")
        LOG_FILE_NAME.format(LOG_PATH, os.sep, TAG_FILE)
    except Exception as exc:
        logger.info("Execution Failed for. Error : \n {}".format(exc))
    finally:
        filename = "IntsaHyreSourcing_{}".format(JobIds)
        slackNotification(filename=filename, filepath=LOG_FILE_NAME.format(LOG_PATH, os.sep, TAG_FILE))
        if driver:
            driver.quit()