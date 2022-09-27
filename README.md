# Excel to IITB Donation Portal

This program is meant to upload data(donations) from a CSV file to the Offline Donation Portal of IIT Bombay.

The script will, 
- Load data from Donation.csv to Pandas Dataframe
- Loop through each row of the Dataframe to,
  - Login to the Offline Donation Portal of IIT Bombay using Selenium
  - Fill up the donation details from the Dataframe
  - Submit the donation


## Prerequisites
- Google Chrome browser
- Python3 and a few below packages
- Copy below commands in your terminal

```bash

   apt install python3-pip # This will install Python3 on Linux systems

   pip3 install pandas
   pip3 install selenium
   pip3 install webdriver-manager
   pip3 install python-dotenv

```

## Usage
- Clone the project
```bash

   git clone https://github.com/khamsakamal48/Excel-to-IITB-Donation-Portal.git

```
- Refer to the Template.CSV file and create a new file **Donation.csv** inside the folder of the project
- Create a .env file with the below parameters and fill up the details
```bash

   MAIL_USERN= # Username of the email account to be used for sending error emails
   MAIL_PASSWORD=  # Password of the email account to be used for sending error emails
   IMAP_URL=  # IMAP Address of the email server to be used for sending error emails
   IMAP_PORT=  # IMAP Port of the email server to be used for sending error emails
   SMTP_URL=  # SMTP Address of the email server to be used for sending error emails
   SMTP_PORT=  # SMTP Port of the email server to be used for sending error emails
   ERROR_EMAILS_TO= # Email address to be used for sending error emails (Recipient Address)
   URL= # URL of the IIT Bombay Offline Donation Portal
   LOGIN= # Username of the IIT Bombay Offline Donation Portal
   PASSWORD= # Password of the IIT Bombay Offline Donation Portal

```
- Run the below command in the terminal

```bash

   python3 'Upload to Portal.py'

```
- If you run into an issue related to Service not found, kindly upgrade Selenium to v4 as per the steps mentioned [here](https://www.selenium.dev/documentation/webdriver/getting_started/upgrade_to_selenium_4/)
```bash

   pip install selenium==4.4.3

```

## Notes
- The dates in the **Donation.csv** file should be in the format: **27/Sep/2022**
- The values in the CSV file should match the dropdown entries as in the IITB Donation Portal
- The Project is by default selected as *'Others'* and the values from the CSV file is captured in the Project Name