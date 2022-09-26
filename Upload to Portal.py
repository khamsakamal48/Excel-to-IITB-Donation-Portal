#!/usr/bin/env python3

import os, time, datetime, sys, ssl, smtplib, imaplib, glob, pprint
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from random import *
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from jinja2 import Environment
from datetime import datetime, date

# Set current directory
os.chdir(os.getcwd())

# Disable Webdriver messages
os.environ['WDM_LOG_LEVEL'] = '0'

# Load webdriver with options
options = webdriver.ChromeOptions()
options.headless = True
options.add_argument("--log-level=3")
s=Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=s,options=options)

# Printing the output to file for debugging
sys.stdout = open('Process.log', 'w')

from dotenv import load_dotenv

load_dotenv()

# Retrieve contents from .env file
MAIL_USERN = os.getenv("MAIL_USERN")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
IMAP_URL = os.getenv("IMAP_URL")
IMAP_PORT = os.getenv("IMAP_PORT")
SMTP_URL = os.getenv("SMTP_URL")
SMTP_PORT = os.getenv("SMTP_PORT")
ERROR_EMAILS_TO  = os.getenv("ERROR_EMAILS_TO")
URL = os.getenv("URL")
LOGIN = os.getenv("LOGIN")
PASSWORD = os.getenv("PASSWORD")
        
def send_error_emails(subject):
    print("Sending email for an error")
    
    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = MAIL_USERN
    message["To"] = ERROR_EMAILS_TO

    # Adding Reply-to header
    message.add_header('reply-to', MAIL_USERN)
        
    TEMPLATE="""
    <table style="background-color: #ffffff; border-color: #ffffff; width: auto; margin-left: auto; margin-right: auto;">
    <tbody>
    <tr style="height: 127px;">
    <td style="background-color: #363636; width: 100%; text-align: center; vertical-align: middle; height: 127px;">&nbsp;
    <h1><span style="color: #ffffff;">&nbsp;Raiser's Edge Automation: {{job_name}} Failed</span>&nbsp;</h1>
    </td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff; border-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="width: 100%; height: 18px; background-color: #ffffff; border-color: #ffffff; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #455362;">This is to notify you that execution of Auto-updating Alumni records has failed.</span>&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff; border-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 61px;">
    <td style="width: 100%; background-color: #2f2f2f; height: 61px; text-align: center; vertical-align: middle;">
    <h2><span style="color: #ffffff;">Job details:</span></h2>
    </td>
    </tr>
    <tr style="height: 52px;">
    <td style="height: 52px;">
    <table style="background-color: #2f2f2f; width: 100%; margin-left: auto; margin-right: auto; height: 42px;">
    <tbody>
    <tr>
    <td style="width: 50%; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #ffffff;">Job :</span>&nbsp;</td>
    <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{{job_name}}&nbsp;</td>
    </tr>
    <tr>
    <td style="width: 50%; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #ffffff;">Failed on :</span>&nbsp;</td>
    <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{{current_time}}&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; width: 100%; background-color: #ffffff; text-align: center; vertical-align: middle;">Below is the detailed error log,</td>
    </tr>
    <tr style="height: 217.34375px;">
    <td style="height: 217.34375px; background-color: #f8f9f9; width: 100%; text-align: left; vertical-align: middle;">{{error_log_message}}</td>
    </tr>
    </tbody>
    </table>
    """
    
    # Create a text/html message from a rendered template
    emailbody = MIMEText(
        Environment().from_string(TEMPLATE).render(
            job_name = subject,
            current_time=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            error_log_message = Argument
        ), "html"
    )
    
    # Add HTML parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(emailbody)
    attach_file_to_email(message, 'Process.log')
    emailcontent = message.as_string()
    
    # Create secure connection with server and send email
    context = ssl._create_unverified_context()
    with smtplib.SMTP_SSL(SMTP_URL, SMTP_PORT, context=context) as server:
        server.login(MAIL_USERN, MAIL_PASSWORD)
        server.sendmail(
            MAIL_USERN, ERROR_EMAILS_TO, emailcontent
        )

    # Save copy of the sent email to sent items folder
    with imaplib.IMAP4_SSL(IMAP_URL, IMAP_PORT) as imap:
        imap.login(MAIL_USERN, MAIL_PASSWORD)
        imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), emailcontent.encode('utf8'))
        imap.logout()

def attach_file_to_email(message, filename):
    # Open the attachment file for reading in binary mode, and make it a MIMEApplication class
    with open(filename, "rb") as f:
        file_attachment = MIMEApplication(f.read())
    # Add header/name to the attachments    
    file_attachment.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # Attach the file to the message
    message.attach(file_attachment)
    
def housekeeping():
    # Housekeeping
    multiple_files = glob.glob("Do*.csv")

    # Iterate over the list of filepaths & remove each file.
    print("Removing old files")
    for each_file in multiple_files:
        try:
            os.remove(each_file)
        except:
            pass
        
def exit():
    
    print('Exiting the script')
    
    sys.exit()
    
def add_to_dataframe():
    
    print('Loading CSV file to Dataframe')
    
    global dataframe
    
    dataframe = pd.read_csv('Donation.csv', sep=',')
    
    print(dataframe)
    
def login_to_portal():
    
    print('Logging in to the Donation Portal')
    
    # Open HRMS Page
    driver.get(URL)
    
    # Wait
    driver.implicitly_wait(30)
    
    # Enter Username
    driver.find_element(By.ID, 'txtusername').send_keys(LOGIN)
    
    # Enter password
    driver.find_element(By.ID, 'txtuserpassword').send_keys(PASSWORD)
    
    # Click on the Login Button
    driver.find_element(By.ID, 'btn_login').click()
    
    # Wait
    driver.implicitly_wait(15)
    
    # Close the Anoymous Donor list
    driver.find_element(By.ID, 'openModalHFRaisedDetails').click()
    
    # Wait
    driver.implicitly_wait(5)
    
def enter_donation_in_portal(first_name, last_name, email, donation_amount, project, project_name, affiliation, batch, department, degree, hostel, address_line_1, address_line_2, country, state, city, zip, pan, phone, payment_type, csr_type, gift_type, hf_grant_no, currency, currency_amount, currency_rate, remarks, transaction_type, cheque, deposited_date, ifsc, account_no, sap_reference_no, office):
    
    print('Adding donation in the Portal')
        
    # Enter First Name
    if first_name == 'nan':
        first_name = '.'

    driver.find_element(By.ID, 'MainContent_txtfirstname').send_keys(first_name)
    
    # Enter Last Name
    if last_name == 'nan':
        last_name = '.'

    driver.find_element(By.ID, 'MainContent_txtlastname').send_keys(last_name)
    
    # Enter Email
    if email == 'nan':
        email = '.'
    
    driver.find_element(By.ID, 'MainContent_txtemail').send_keys(email)
        
    # Click on Proceed
    driver.find_element(By.ID, 'MainContent_btnLogin').click()
    
    # Wait
    driver.implicitly_wait(30)
    
    # Enter Donation Amount
    if donation_amount == 'nan':
        donation_amount = '0'
    
    driver.find_element(By.ID, 'MainContent_txtDonationamount').click()
    driver.implicitly_wait(1)
    driver.find_element(By.ID, 'MainContent_txtDonationamount').send_keys(donation_amount)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Select Project (By Default others will be selected)
    driver.find_element(By.ID, 'MainContent_ddlProject').send_keys(project)
    
    if project == 'Others':
        if project_name == 'nan':
            project_name = 'Others'
            
        driver.find_element(By.ID, 'MainContent_txtProjectName').send_keys(project_name)
        
    # Wait
    driver.implicitly_wait(5)
        
    # Enter Affiliation
    if affiliation == 'nan':
        affiliation = 'Wellwisher'

    driver.find_element(By.ID, 'MainContent_ddlAffilation').send_keys(affiliation)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Batch
    if batch == 'nan':
        batch = 'NA'

    driver.find_element(By.ID, 'MainContent_ddlBatch').send_keys(batch)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Department
    if department == 'nan':
        department = 'Other'

    driver.find_element(By.ID, 'MainContent_ddlDepartment').send_keys(department)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Degree
    if degree == 'nan':
        degree = 'Other'

    driver.find_element(By.ID, 'MainContent_ddlDegree').send_keys(degree)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Hostel
    if hostel == 'nan':
        hostel = 'Other'

    driver.find_element(By.ID, 'MainContent_ddlHostel').send_keys(hostel)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Address Line 1
    if address_line_1 != 'nan':
        driver.find_element(By.ID, 'MainContent_txtAddr1').send_keys(address_line_1)
        
        # Wait
        driver.implicitly_wait(5)
     
    # Enter Address Line 2
    if address_line_2 != 'nan':
        driver.find_element(By.ID, 'MainContent_txtAddr2').send_keys(address_line_2)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Country
    if country == 'nan':
        country = 'Other'

    driver.find_element(By.ID, 'MainContent_ddlCountry').send_keys(country)
    
    # Wait
    driver.implicitly_wait(5)
    
    # # Enter State
    # if state == 'nan':
    #     state = 'NA'

    # driver.find_element(By.ID, 'MainContent_ddlState').send_keys(state)
    
    # # Wait
    # driver.implicitly_wait(5)
    
    # # Enter City
    # if city != 'nan':
    #     driver.find_element(By.ID, 'MainContent_ddlCity').send_keys(city)
        
    #     # Wait
    #     driver.implicitly_wait(5)
    
    # Enter Zip
    if zip != 'nan':
        driver.find_element(By.ID, 'MainContent_txtZipcode').send_keys(zip)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter PAN
    if country == 'India':
        driver.find_element(By.ID, 'MainContent_txtPancard').send_keys(pan)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Contact No.
    if phone != 'nan':
        driver.find_element(By.ID, 'MainContent_txtContact').send_keys(phone)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Payment Type
    if payment_type != 'nan':
        driver.find_element(By.ID, 'MainContent_rbPaymentType_0').click()
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter CSR Type
    if csr_type != 'nan':
        driver.find_element(By.ID, 'MainContent_rbCSRType_0').click()
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Gift Type
    if gift_type != 'nan':
        driver.find_element(By.ID, 'MainContent_rbGiftType_1').click()
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter HF Grant No.
    if office == 'HF' and hf_grant_no != 'nan':
        driver.find_element(By.ID, 'MainContent_txtHFGrant').send_keys(hf_grant_no)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Currency
    if currency != 'nan':
        driver.find_element(By.ID, 'MainContent_ddlCurrency').send_keys(currency)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Currency Amount
    if currency_amount != 'nan':
        driver.find_element(By.ID, 'MainContent_txtCurrencyAmount').send_keys(currency_amount)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Currency Rate
    if currency_rate != 'nan':
        driver.find_element(By.ID, 'MainContent_txtCurrencyRate').send_keys(currency_rate)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Remarks
    if remarks != 'nan':
        driver.find_element(By.ID, 'MainContent_txtRemarks').send_keys(remarks)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Type of Transaction
    if transaction_type == 'nan':
        transaction_type = 'NEFT/RTGS'
        
    driver.find_element(By.ID, 'MainContent_ddlTransaction').send_keys(transaction_type)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Cheque No.
    if cheque != 'nan':
        driver.find_element(By.ID, 'MainContent_txtChequeNo').send_keys(cheque)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Deposited Date
    if deposited_date == 'nan':
        deposited_date = datetime.now().strftime("%d/%b/%Y")
        
    driver.find_element(By.ID, 'MainContent_txtChequeDate').send_keys(deposited_date)
        
    # Wait
    driver.implicitly_wait(5)
    
    # Enter IFSC Code
    if ifsc == 'nan':
        ifsc = '.'
        
    driver.find_element(By.ID, 'MainContent_txtChequeBank').send_keys(ifsc)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter Account No.
    if account_no == 'nan':
        account_no = '.'
        
    driver.find_element(By.ID, 'MainContent_txtAccountNo').send_keys(account_no)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Enter SAP Reference No.
    if sap_reference_no != 'nan':
        driver.find_element(By.ID, 'MainContent_txt_sapreference').send_keys(sap_reference_no)
        
        # Wait
        driver.implicitly_wait(5)
    
    # Enter Office
    if office == 'nan':
        office = 'ACR Office'
        
    driver.find_element(By.ID, 'MainContent_ddlOffice').send_keys(office)
    
    # Wait
    driver.implicitly_wait(5)
    
    # Click the Donate Button
    # driver.find_element(By.ID, 'MainContent_btn_Donation').click()

    
def upload_donation_to_portal():
    
    # global first_name, last_name, email, donation_amount, project, project_name, affiliation, batch, department, degree, hostel, address_line_1, address_line_2, country, state, city, zip, pan, payment_type, csr_type, gift_type, hf_grant_no, currency, currency_amount, currency_rate, remarks, transaction_type, cheque, deposited_date, ifsc, account_no, sap_reference_no, office
    
    for index, row in dataframe.iterrows():
        
        first_name = row['First Name']
        last_name = row['Last name']
        email = row['Email']
        donation_amount = row['Donation Amount']
        project = row['Project']
        project_name = row['Project Name']
        affiliation = row['Affiliation']
        batch = row['Batch']
        department = row['Department']
        degree = row['Degree']
        hostel = row['Hostel']
        address_line_1 = row['Address Line 1']
        address_line_2 = row['Address Line 2']
        country = row['Country']
        state = row['State']
        city = row['City']
        zip = row['Zip']
        pan = row['PAN']
        phone = row['Contact No.']
        payment_type = row['Payment Type']
        csr_type = row['CSR Type']
        gift_type = row['Gift Type']
        hf_grant_no = row['HF Grant No.']
        currency = row['Currency']
        currency_amount = row['Currency Amount']
        currency_rate = row['Currency Rate']
        remarks = row['Remarks']
        transaction_type = row['Type of Transaction']
        cheque = row['Cheque No.']
        deposited_date = row['Deposited Date']
        ifsc = row['IFSC Code']
        account_no = row['Account No.']
        sap_reference_no = row['SAP Reference No.']
        office = row['Office']
        
        print(f'Working on Line No. {index + 1} with content:\n{row}')
    
        # Login to the Donation Portal
        login_to_portal()
        
        # Enter donation in the portal
        enter_donation_in_portal(first_name, last_name, email, donation_amount, project, project_name, affiliation, batch, department, degree, hostel, address_line_1, address_line_2, country, state, city, zip, pan, phone, payment_type, csr_type, gift_type, hf_grant_no, currency, currency_amount, currency_rate, remarks, transaction_type, cheque, deposited_date, ifsc, account_no, sap_reference_no, office)
        
        # Close the browser
        driver.quit()
        
        # Sleep for 2 seconds
        time.sleep(2)
    
try:
    
    # Add CSV to Dataframe
    add_to_dataframe()
    
    # Upload donation in Dataframe to Portal
    upload_donation_to_portal()

    
except Exception as Argument:
    send_error_emails(subject='Error while downloading opportunity list from Raisers Edge')
    
finally:
    
    # Housekeeping
    # housekeeping()
    
    # Proper Exit
    exit()