from datetime import datetime

mydate = datetime.now()
Year = (datetime.now().strftime('%Y'))
Month = (datetime.now().strftime('%m'))
Date = (datetime.now().strftime('%d'))

config = {
    'from': '',
    'to': '',
    'server': 'smtp.gmail.com',
    'port': '587',
    'tls': 'yes',
    'login': '',
    'password': '',
    'jira_username':'',
    'jira_password':'',
    'start_date':'2018-10-20',   ###Specify the date in YYYY-MM-DD format
    'end_date': '2018-10-30',
    'subject_keyword_DQ_Quality_Tickets': 'DQ tickets for the month {0}-{1}'.format(mydate.strftime("%B"), Year),
    'sql_hostname' : '',
    'sql_username' :'',
    'sql_password' :'',
    'sql_main_database' : '',
    'sql_port' : 3306,
    'ssh_host' : '',
    'ssh_user' : '',
    'ssh_port' : 22
}
