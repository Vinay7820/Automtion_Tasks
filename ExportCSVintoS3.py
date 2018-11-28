import re
import boto3
import time
import pandas as pd
from jira import JIRA
from config import config
from datetime import date, timedelta

print("This Script will query your issues from JIRA and export to S3 Bucket")

try:
    print("Logging into JIRA")
    jira = JIRA(basic_auth=(config['jira_username'], config['jira_password']), options={'server': 'https://jira.move.com'})
    print("Logged into JIRA succesfully")
except Exception as e:
    print("Not able to login to JIRA. Please Check the connectivity and try again!")
    print(e)

if config['start_date'] != '' and config['end_date'] != '':
    print("Taking the Specified Date Range from config File and forming the query")
    query = 'project = DE AND issuetype = "Data Quality" AND status in (Blocked, "To Do", Deferred, "Waiting on Others", "In Progress", Open, "In Review", Reopened, Resolved, Closed, Done, "In Review / Validation", "Waiting for Others", "User Misunderstanding", Fixed) AND created >= {0} AND created <= {1}  ORDER BY key ASC, status ASC, cf[10500] ASC, created DESC, due DESC'.format(config['start_date'], config['end_date'])
    print('Query Date Period - from {0} to {1}'.format(config['start_date'], config['end_date']))
    Date_Range = 1
else:
    Date_Range = 0
    print("Date Range is not specified in the config File - By Default taking previous 7 days data to form the query")
    d = date.today() - timedelta(days=7)
    query = 'project = DE AND issuetype = "Data Quality" AND status in (Blocked, "To Do", Deferred, "Waiting on Others", "In Progress", Open, "In Review", Reopened, Resolved, Closed, Done, "In Review / Validation", "Waiting for Others", "User Misunderstanding", Fixed) AND created >= {0} AND created <= {1} ORDER BY key ASC, status ASC, cf[10500] ASC, created DESC, due DESC'.format(d,date.today())
    print('Query Period - From {0} to {1}'.format(d,date.today()))


print("Querying JIRA for Issues")
issues_in_project = jira.search_issues(query)
time.sleep(15)
i = 0
df = pd.DataFrame(columns=['Key','Status','Created_Date','Updated_Date','Assignee', 'Reporter', 'Data_Category','Label'])
for issue in issues_in_project:
    issue_key = issue.key
    raw_data = issue.raw
    print(raw_data)
    searchObj = re.search(r"('value': 'Page Views'| 'value': 'Leads'| 'value': 'Other'| 'value': 'UUs' | 'value': 'Visits')", str(raw_data))
    if searchObj:
        remove_quotes = (searchObj.group().replace("'", ""))
        issue_data_category = remove_quotes.split(': ')[1]
    else:
        issue_data_category =  "No Data Category Associated with this issue or Data Category is different"
    issue_status = issue.fields.status.name
    issue_created_date = issue.fields.created
    issue_updated_date = issue.fields.updated
    issue_assignee = issue.fields.assignee.name
    issue_reporter = issue.fields.reporter.name
    issue_label = issue.fields.labels
    df.loc[i] = [issue_key,issue_status,issue_created_date,issue_updated_date,issue_assignee,issue_reporter,issue_data_category,issue_label]
    i+=1
    #print(df)

if Date_Range:
    savefilename = 'Data_Quality_Tickets_from_{0}_to_{1}.csv'.format(config['start_date'], config['end_date'])
else:
    savefilename = 'Data_Quality_Tickets_from_{0}_to_{1}.csv'.format(d, date.today())

print('Converting the data into file - {0}'.format(savefilename))
df.to_csv(savefilename)
print('Completed Writing Issues to CSV File - {0}'.format(savefilename))

print("Connecting to AWS S3 to upload the file")
s3_client = boto3.client(service_name='s3')
local_filepath = 'C:/Users/msunku/PycharmProjects/Automation/Data_Quality_tickets.csv'
bucket_name = 'move-dataeng-temp-dev'
if Date_Range:
    s3_filepath = 'Rajesh/Data_Quality_tickets_from_{0}_to_{1}.csv'.format(config['start_date'], config['end_date'])
else:
    s3_filepath = 'Rajesh/Data_Quality_tickets_from_{0}_to_{1}.csv'.format(d,date.today())
try:
    s3_client.upload_file(local_filepath, bucket_name, s3_filepath)
    print("Uploaded to " + "s3://" + bucket_name + "/" + s3_filepath)
except BaseException as e:
    print("Upload error for " + local_filepath)
    print(str(e))