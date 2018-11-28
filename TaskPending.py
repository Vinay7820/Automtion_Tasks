import re
import time
import smtplib
import jira.client
from jira import JIRA
from jira.client import JIRA
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from config import config
import pandas as pd
from datetime import date

print("This script queries for the issues which are in PENDING STATE!")

try:
    print("Trying to Login into JIRA")
    jira = JIRA(basic_auth=(config['jira_username'], config['jira_password']), options={'server': 'https://jira.move.com'})
    print("Logged into JIRA succesfully")
except Exception as e:
    print("Not able to login into JIRA. Please check the connection and try again!")
    print(e)

if config['start_date'] != '' and config['end_date'] != '':
    print("Taking the Specified Date Range from config File and forming the query")
    query = 'project = DE AND issuetype = "Data Quality" AND status in (Blocked, "To Do", Deferred, "Waiting on Others", "In Progress", Open, "In Review", Reopened, "In Review / Validation", "Waiting for Others", "User Misunderstanding") AND created >= {0} AND updated <= {1} ORDER BY key ASC, status ASC, cf[10500] ASC, created DESC, due DESC'.format(config['start_date'], config['end_date'])
    print('Query Date Period - from {0} to {1}'.format(config['start_date'], config['end_date']))
    Date_Range = 1
else:
    Date_Range = 0
    print("Date Range is not specified in the config File - By Default taking data from 1st Apr 2018 to till date")
    query = 'project = DE AND issuetype = "Data Quality" AND status in (Blocked, "To Do", Deferred, "Waiting on Others", "In Progress", Open, "In Review", Reopened, "In Review / Validation", "Waiting for Others", "User Misunderstanding") AND created >= 2018-04-01 AND updated <= {0} ORDER BY key ASC, status ASC, cf[10500] ASC, created DESC, due DESC'.format(date.today())
    print('Query Period - From 2018-04-01 to {0}'.format(date.today()))

print("Querying JIRA for the issues")
issues_in_project = jira.search_issues(query)
time.sleep(15)
i = 0
df = pd.DataFrame(columns=['Key', 'Label', 'Summary', 'Data_Category', 'Status', 'Assignee', 'Reporter', 'Created', 'Updated'])
for issue in issues_in_project:
    issue_key = issue.key
    raw_data = issue.raw
    searchObj = re.search(r"('value': 'Page Views'| 'value': 'Leads'| 'value': 'Other'| 'value': 'UUs' | 'value': 'Visits')", str(raw_data))
    if searchObj:
        remove_quotes = (searchObj.group().replace("'", ""))
        issue_data_category = remove_quotes.split(': ')[1]
    else:
        issue_data_category = "No Data Category Associated with this issue or Data Category is different"
    issue_status = issue.fields.status.name
    issue_summary = issue.fields.summary
    issue_created_date = issue.fields.created
    issue_updated_date = issue.fields.updated
    issue_assignee = issue.fields.assignee.name
    issue_reporter = issue.fields.reporter.name
    issue_label = issue.fields.labels
    df.loc[i] = [issue_key, issue_label, issue_summary, issue_data_category,issue_status, issue_assignee,  issue_reporter, issue_created_date, issue_updated_date]
    i += 1
    #print(df)

if Date_Range:
    savefilename = 'Tickets Pending from {0} to {1}.xlsx'.format(config['start_date'], config['end_date'])
else:
    savefilename = 'Tickets_Pending from 2018-04-01 to {0}.xlsx'.format(date.today())

print('Converting the data into file - {0}'.format(savefilename))
df.to_excel(savefilename)
print('Completed Writing Issues to XLSX File - {0}'.format(savefilename))

def send_mail(send_from, send_to, subject, text, server, port, username='', password='', isTls=True):
    try:
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(savefilename, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename=savefilename')
        msg.attach(part)
        # context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
        # SSL connection only working on Python 3+
        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        print("Email Sent")

    except Exception as e:
        print("Could not send the email. Please Check!")
        print(e)


def main():
    subject_keyword_taskpending_no_date_range =  'DQ Tickets Pending since 2018-04-01 to {0}'.format(date.today())
    subject_keyword_taskpending_date_range = 'DQ Tickets Pending from {0} to {1}'.format(config['start_date'], config['end_date'])
    if Date_Range:
        send_mail(config['from'], config['to'],subject_keyword_taskpending_date_range , "Additional Information to be added", config['server'],config['port'], config['login'], config['password'])
    else:
        send_mail(config['from'], config['to'], subject_keyword_taskpending_no_date_range,"Additional Information to be added", config['server'], config['port'], config['login'],config['password'])

if __name__ == '__main__':
    main()