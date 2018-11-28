import pymysql
import smtplib
import paramiko
from sshtunnel import SSHTunnelForwarder
from os.path import expanduser
from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from config import config

query_DP = """SELECT aud.app_id,job_id,stage,start_time FROM etl_audit.auditdata aud
inner join etl_audit.etl_mst mst
on aud.app_id = mst.app_id
WHERE end_time is NULL
AND flag='Y'
AND start_time >= hour(now())-2
AND date(start_time)=current_date() and aud.app_id is not null; """


query_batch_jobs = """SELECT app_id,job_id,start_time FROM etl_audit.batchauditdata
WHERE end_time is NULL
AND start_time >= hour(now())-2;
"""

DP_Subject = "DP Jobs that are executing for more than 2 hrs"
Batch_Subject = "Batch Jobs that are executing for more than 2 hrs"

firstrow_DP = ['App_ID', 'Job_ID', 'Stage', 'Start_Time']
firstrow_Batch = ['App_ID', 'Job_ID','Start_Time']

def execute(query):
    home = expanduser('~')
    mypkey = paramiko.RSAKey.from_private_key_file(home + "\\MySQL.pem")

    try:
        print("Trying to login to MYSQL Server")
        with SSHTunnelForwarder(
                (config['ssh_host'], config['ssh_port']),
                ssh_username=config['ssh_user'],
                ssh_pkey=mypkey,
                remote_bind_address=(config['sql_hostname'], config['sql_port'])) as tunnel:
            conn = pymysql.connect(host='127.0.0.1', user=config['sql_username'],
                                   passwd=config['sql_password'], db=config['sql_main_database'],
                                   port=tunnel.local_bind_port)
            data = pd.read_sql_query(query, conn)
            conn.close()
            return data
    except Exception as e:
        print("Not able to login to MySQL Server. Please Check the connectivity")
        print(e)


def send_mail(data, subject, firstrow):
    text = """
    Hello, Friend.

    Here is your data:

    {table}

    Regards,

    Me"""

    html = """
    <html>
    <head>
    <style> 
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 5px; }}
    </style>
    </head>
    <body><p>Hello, Friend This data is from a data frame.</p>
    <p>Here is your data:</p>
    {table}
    <p>Regards,</p>
    <p>Me</p>
    </body></html>
    """

    text = text.format(table=tabulate(data, headers=firstrow, tablefmt="grid"))
    html = html.format(table=tabulate(data, headers=firstrow, tablefmt="html"))

    message = MIMEMultipart(
        "alternative", None, [MIMEText(text), MIMEText(html, 'html')])
    try:
        print("Trying to send the mail")
        message['Subject'] = subject
        message['From'] = config['from']
        message['To'] = config['to']
        server = smtplib.SMTP("smtp.gmail.com:587")
        server.ehlo()
        server.starttls()
        server.login(config['login'], config['password'])
        server.sendmail(config['from'], config['to'], message.as_string())
        server.quit()
    except Exception as e:
        print("Not able to send the Mail due to some exception. Please Check!")
        print(e)

def main():
    print("This script fetches DP and Batch Jobs which are running beyond 2 hrs")
    print("DP Jobs")
    data = execute(query_DP)
    send_mail(data, DP_Subject, firstrow_DP)
    print("Batch Jobs")
    data = execute(query_batch_jobs)
    send_mail(data, Batch_Subject, firstrow_Batch)

if __name__ == '__main__':
    main()
