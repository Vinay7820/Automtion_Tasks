#/usr/bin/python3
##############################################################################################################################
#   Date            Author          Description:
#   05-21-2018     Brillio          Script to create Weekly KPI Report
##############################################################################################################################


import os, sys, boto3, re, xlrd, xlwt, json
from xlutils.copy import copy

#import git
#from pyathenajdbc import connect
from move_dl_common_api.athena_util import AthenaUtil

'''def get_git_repository(git_url):
    git.Git().clone(git_url)


    #get_git_repository("https://e3a92906b38e210876d50d4f833ba134a349cb3d@github.move.com/DataEngineering/Common-Queries")
    get_git_repository("https://github.move.com/DataEngineering/Metrics_Framework/tree/master/Metrics_SQL_Queries")'''

#https://github.move.com/DataEngineering/Metrics_Framework/tree/master/Metrics_SQL_Queries

#This is sample
'''WeeklyKPI = {'Total Web UUs':'uu_web_wk.sql','Paid UUs':'uu_paid_wk.sql'}'''

#This part is working
WeeklyKPI = {'Total Web UUs':'uu_web_wk.sql','Total iOS Core App UUs':'Total iOS Core App UUs.sql',
           'Total iOS Other App UUs':'Total iOS Other App UUs.sql','Total iOS App UUs':'Total iOS Apps UUs.sql',
           'Total Android Core App UUs':'Total Android Core App UUs.sql',
           'Total Android Instant App UUs':'Total Android Instant Apps UUs.sql',
           'Total Android Other App UUs':'Total Android Other Apps UUs.sql','Non-Paid UUs':'Total Non-Paid UUs.sql',
           'SEO UUs':'TOtal SEO UUS.sql','Paid UUs':'Total Paid UU.sql',
           'Total Web FS UUs':'Total Web FS UU.sql','Total iOS Core App FS UUs':'Total iOS Core Apps FS UUs.sql',
           'Total iOS Other App FS UUs':'Total iOS Other Apps FS UUs.sql',
           'Total iOS App FS UUs':'Total iOS Apps FS UUs.sql',
           'Total Android Core App FS UUs':'Total Android Core App FS UUs.sql','Total Android Instant App FS UUs':'Total Android Instant App FS UUs.sql',
           'Total Android Other App FS UUs':'Total Android Other App FS UUs.sql',
           'Total For Sale Leads':'Total For Sale Leads.sql','Total CfB Leads':'Total CfB Leads.sql','Total Advantage Leads':'Total Advantage Leads.sql',
           'Total iOS Core Apps LDP FS PVs':'Total iOS Core Apps LDP FS PVs.sql',
           'Total Android Core Apps LDP FS PVs':'Total Android Core Apps LDP FS PVs.sql','Total Mobile Apps LDP FS PVs':'Total Mobile Apps LDP FS PVs.sql',
           'Total Phone Web LDP FS PVs':'Total Phone Web LDP FS PVs.sql',
           'Total Web LDP FS PVs':'Total Web LDP FS PVs.sql',
           'Total Advantage Email Leads':'Total Advantage Email Leads.sql',
           'Total Advantage Phone+TPN Leads':'Total Advantage Phone+TPN Leads.sql',
           'Direct UUs (Web Only)':'Direct UUs (Web Only).sql',
           'Total Desktop':'Total Desktop Tab Web LDP FS PVs.sql'}

#This part is not working
#WeeklyKPI = {'Total Desktop':'Total Desktop Tab Web LDP FS PVs.sql'}
#columntype = {'0':'text_format','1':'date_format','2':'date_format','3':'text_format','4':'General','5': 'text_format', '6': 'text_format', '7': 'date_format', '8': 'text_format','9': 'date_format'}


'''def connect_athena():
    try:
        conn = connect(s3_staging_dir='s3://move-dataeng-temp-prod/pthakur/athena-results/',
                       region_name="us-west-2")
        return conn
    except Exception as err:
        print ("Unable to connect to Athena...(%s)" %(err))'''

def extract_athena_data(weekly_query):
    try:
        print ("Extract the data from Athena...!!!")
        #conn = connect_athena()
        #curr = conn.cursor()
        if os.path.exists('WeeklyKpiReport.xls'):
            wbk = xlrd.open_workbook('WeeklyKpiReport.xls', formatting_info=True)
            workbook = copy(wbk)
        else:
            workbook = xlwt.Workbook()

        worksheet = workbook.add_sheet(weekstartdate)
        '''bold = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        text_format = workbook.add_format()

        # Add a number format for cells with money.
        money = workbook.add_format({'num_format': '$#,##0'})'''
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'yyyy-mm-dd'
        # Write some data headers.
        worksheet.write(0, 0, 'metric_name')
        worksheet.write(0, 1, 'metric_id')
        worksheet.write(0, 2, 'Week Start Date')
        worksheet.write(0, 3, 'Week End Date')
        worksheet.write(0, 4, 'time_dim_key')
        worksheet.write(0, 5, 'metric_source_id')
        worksheet.write(0, 6, 'metric_value')
        worksheet.write(0, 7, 'created_by')
        worksheet.write(0, 8, 'created_date')
        worksheet.write(0, 9, 'updated_by')
        worksheet.write(0, 10, 'updated_date')
        row = 1
        col = 0

        util = AthenaUtil(s3_staging_folder='s3://move-dataeng-temp-prod/athena-results/')
        for k, v in weekly_query.items():
            print (v)
            #curr.execute(v)
            #queryoutput = curr.fetchone()
            snap_result = util.execute_query(sql_query=v, use_cache=False)
            queryoutput = snap_result["ResultSet"]["Rows"][0]['Data']
            print (queryoutput[0])
            worksheet.write(row, col, k)
            for i in range(len(queryoutput)):
                if i == 1 or i == 2 or i == 7 or i == 9:
                    worksheet.write(row, col + 1 + i, queryoutput[i]['VarCharValue'], date_format)
                else:
                    worksheet.write(row, col + 1 + i, queryoutput[i]['VarCharValue'])

            query = ("""INSERT INTO dm_rdckpi.metric_actual_fact(metric_id, timedimkey, metric_source_id, metric_value, created_by, created_date, updated_by, updated_date) VALUES(%s,'%s',%s,%s,'%s','%s','%s','%s');""" %(queryoutput[0]['VarCharValue'],queryoutput[3]['VarCharValue'],queryoutput[4]['VarCharValue'],queryoutput[5]['VarCharValue'],queryoutput[6]['VarCharValue'],queryoutput[7]['VarCharValue'],queryoutput[8]['VarCharValue'], queryoutput[9]['VarCharValue']))
            worksheet.write(row, 11, query)
            row += 1
        workbook.save('WeeklyKpiReport.xls')

    except Exception as err:
        print ("Here is the error...('%s')" %(err))

def read_query_sql(filename,weekstartdate):
    try:
        print ("Hi....Get the query")
        session = boto3.Session()
        s3 = session.client('s3')
        #bucket = s3.Bucket('move-dataeng-temp-dev')
        file_key = 'pthakur/athena-results/weekly_kpi_sql/' + filename
        obj = s3.get_object(Bucket='move-dataeng-temp-dev', Key=file_key)
        sql_data = obj['Body'].read().decode('utf-8')
        if filename == 'Direct UUs (Web Only).sql':
            split_data = sql_data.split('/')
            query = split_data[2] + '/' + split_data[3]
            query_string = query.replace("$Period_Start_Date", weekstartdate)
            print(query_string)
        else:
            query_string = sql_data.split('/')[2].replace("\n"," ").replace("$Period_Start_Date",weekstartdate)
        return query_string
    except Exception as err:
        print ("Exception is here...(%s)" %(err))


def main(weekstartdate):
    try:

        weekly_query = {}
        print ("Hi...I am going to execute!!!")
        # import pdb;
        # pdb.set_trace()
        for key,value in WeeklyKPI.items():
            query_stmt = read_query_sql(WeeklyKPI[key],weekstartdate)
            print (query_stmt)
            weekly_query[key] = str(query_stmt)

        #print (weekly_query)
        extract_athena_data(weekly_query)

    except Exception as err:
        print ("Hi...Here is the error.... (%s)" %(err))


if __name__ == "__main__":
    arguments_count = len(sys.argv)
    weekstartdate = sys.argv[1]
    # get_git_repository("https://e3a92906b38e210876d50d4f833ba134a349cb3d@github.move.com/DataEngineering/Common-Queries")
    '''try:
        get_git_repository("https://github.move.com/DataEngineering/Metrics_Framework/tree/master/Metrics_SQL_Queries")
    except Exception as err:
        print err'''
    main(weekstartdate)