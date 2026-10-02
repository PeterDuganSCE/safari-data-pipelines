# -*- coding: utf-8 -*-
import sys

import pandas as pd
import warnings
import os
import pyodbc
# from sqlalchemy import create_engine
import datetime as dt
import datetime
from datetime import datetime
import glob
from openpyxl import load_workbook
from pathlib import Path


warnings.filterwarnings('ignore')
pd.set_option('display.width', 500)
pd.set_option('display.max_column', 500)
pd.set_option('display.max_rows', 200)
pd.set_option('display.max_colwidth', 500)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.sharepoint.graph_api import authenticate_user
from shared.sharepoint.utils import upload_dataframe_to_sharepoint_list
from shared.utils import system_assignments

# iTOA SQL Server connection details
DB_DRIVER = "{SQL Server}" # Or whatever driver is installed in your machine
DB_SERVER = 'tcp:AYWCPSQL116\SQLP740,1989'  # Irvine Server
# DB_SERVER = 'tcp:AYWCPSQL651\SQLP648,1989'  # Alhambra Server
DATABASE_NAME = 'ITOA'  # development server, for testing
itoa_connection = pyodbc.connect(f"DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DATABASE_NAME};Trusted_Connection=no;")


itoa_query = '''SELECT DISTINCT APP_ID, LINE_OF_BUSINESS, APPLICATION_STATUS, EQUIPMENT, VOLTAGE_LEVEL, MAIN_ACTUAL_OUT, DISTURBANCE_DURATION, DISTURBANCE_NO, AFFECTED_DISTRICT, WORK_CENTERS, WEATHER, MISCELLANEOUS_COMMENTS, event_activity_comments, INITIATING_CAUSE_CODE, INITIATING_SUB_CAUSE, CAUSE_CATEGORY, OUTAGE_CATEGORY, PROCESS_STATUS FROM dbo.GENERIC_INT_TADS_V'''
df_itoa = pd.read_sql(itoa_query, itoa_connection)

df_itoa = df_itoa.drop_duplicates()
df_itoa = df_itoa[df_itoa['LINE_OF_BUSINESS'] == 'Transmission']
df_itoa = df_itoa[df_itoa['PROCESS_STATUS'] != 'VOID']
df_itoa['iTOA_link'] = df_itoa['APP_ID'].apply(lambda x: f'https://orls.sce.com/itoa/autooutage/view.htmlx?editedOutage.appId={x}&eventAnalysis.analysisId=&referer=interruption')
del df_itoa['PROCESS_STATUS']

# =============================================================================
# Need to run info for 2 days prior since notifications loaded at 6pm, so have to make sure all created on dates give the chance for info to get into system
# =============================================================================
# today = datetime.now().strftime("%m/%d/%Y")
today_str = datetime.today().strftime(" %d%b%Y")
today = datetime.strptime(today_str, " %d%b%Y")

#Get date of last run to pull data from days that it didnt run
list_of_files = glob.glob('pipelines/transmission/previous_runs_itoa/*')
latest_file = max(list_of_files, key=os.path.getctime)


t = latest_file.split("iTOA_Transmission_Events_")[1]
t = t.split(".")[0]
t = t.replace("_", "/")
last_date = pd.to_datetime(t)
last_date_str = last_date.strftime(" %d%b%Y")

#Load previous file to prevent duplicates when running the extra day for 6pm to midnight data
prev_run_file = pd.read_excel(latest_file)

#One delay to take notifications back another day in case events came in after 6pm that day (Loaded data at 6pm so cutoff in SAP)
full_day = last_date - dt.timedelta(days=1)
full_day_str = full_day.strftime(" %d%b%Y")

yesterday = today - dt.timedelta(days=2)
yesterday_str = yesterday.strftime(" %d%b%Y")

# =============================================================================
# %% Limit to last day
# =============================================================================
df_itoa['Out_Date'] = pd.to_datetime(df_itoa['MAIN_ACTUAL_OUT'], errors='coerce')


# today = dt.datetime.today()
# yesterday = today - dt.timedelta(days=1)
# today = today.strftime('%m/%d/%Y')
# formatted_today = dt.datetime.strptime(today, '%m/%d/%Y')

# yesterday = yesterday.strftime('%m/%d/%Y')
itoa_gen_t_2 = df_itoa[df_itoa['Out_Date'] >= full_day]
itoa_gen_t_2 = itoa_gen_t_2[itoa_gen_t_2['Out_Date'] < today]


if len(itoa_gen_t_2) > 0:
#Write out data for team to see the information
    formatted_today = today = today.strftime('%m/%d/%Y')
    formatted_today = dt.datetime.strptime(formatted_today, '%m/%d/%Y') #convert back to a datetime object
    month_number = str(formatted_today.month)
    day_number = str(formatted_today.day)
    year_number = str(formatted_today.year)

# =============================================================================
# Clean up circuit name
# =============================================================================
    itoa_gen_t_3 = itoa_gen_t_2

    if itoa_gen_t_3['EQUIPMENT'].str.contains(' kV CB').any():
        itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].apply(lambda x: x.split(" kV CB")[0])
        
    if itoa_gen_t_3['EQUIPMENT'].str.contains(' kV Line').any():
        itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].apply(lambda x: x.split(" kV Line")[0])
            
    if itoa_gen_t_3['EQUIPMENT'].str.contains(' KV LINE').any():
        itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].apply(lambda x: x.split(" KV LINE")[0])

#Remove space and voltage from Equipment(Ckt name)
    itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3.apply(
        lambda row: row['EQUIPMENT'].replace(f" {row['VOLTAGE_LEVEL']}", "").strip() 
                 if f" {row['VOLTAGE_LEVEL']}" in row['EQUIPMENT'] 
                 else row['EQUIPMENT'], 
    axis=1
    )
# =============================================================================
# Bring in if circuit is HFRA
# =============================================================================
    trans_ckt_info = pd.read_excel('data/HFRA Transmission and Subtransmission Master Circuit List.xlsx')

    trans_ckt_info_sub = trans_ckt_info[['Circuit Name', 'HFRA TIERS']]
    
    itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].fillna('')
    
    itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].str.upper()
    
    trans_ckt_info_sub['Circuit Name'] = trans_ckt_info_sub['Circuit Name'].astype(str)
    
    itoa_gen_t3 = pd.merge(itoa_gen_t_3, trans_ckt_info_sub, left_on= ['EQUIPMENT'],
                         right_on= ['Circuit Name'], how='left').drop(['Circuit Name'], axis=1)

    itoa_gen_t3['HFRA TIERS'] = itoa_gen_t3['HFRA TIERS'].fillna('Non HFRA')

    def hfra_check(tier): # decide highest level Tier
        if 'T3' in tier:
            return 'T3'
        elif 'T2' in tier:
            return 'T2'
        else:
            return 'NON-HFRA'

    itoa_gen_t3['HFRA_UDF'] = itoa_gen_t3.apply(lambda x: hfra_check(x['HFRA TIERS']), axis=1)
    
# =============================================================================
# Fix if 2 switching centers in data
# =============================================================================
    itoa_gen_t3['WORK_CENTERS'] = itoa_gen_t3['WORK_CENTERS'].fillna('')

    if itoa_gen_t3['WORK_CENTERS'].str.contains(',').any():
        itoa_gen_t3['WORK_CENTERS'] = itoa_gen_t3['WORK_CENTERS'].apply(lambda x: x.split(",")[0])

# =============================================================================
# Assign Engineers based on System Assignments in Keyword Scrape which comes from PowerAutomate flow to update excel there
# =============================================================================
#Engineer Assigned from System Assignments
    sys_assign = system_assignments()  # pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/System_Assignments.xlsx')
    # sys_assign = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/System_Assignments.xlsx')

    itoa_gen_t3['WORK_CENTERS'] = itoa_gen_t3['WORK_CENTERS'].str.upper()

    itoa_gen_t4 = pd.merge(itoa_gen_t3, sys_assign, left_on= ['WORK_CENTERS'],
                         right_on= ['Switching Center'], how='left').drop(['Switching Center'], axis=1)

    itoa_gen_t4['Engineer Assigned'] = itoa_gen_t4['Engineer Assigned'].fillna('Unassigned')

    itoa_gen_t4['Subject'] = 'Trans_itoa_' + itoa_gen_t4['APP_ID'].astype(str)
    itoa_gen_t4['Entry_Method'] = 'ILS'
    
    
    def extract_substation(name):
        if pd.isna(name) or name.strip() == "":
            return ""
        else:
        # Split by the dash and return the first part
            return name.split('-')[0].strip()

# Create a new column 'Substation' using apply
    itoa_gen_t4['Substation'] = itoa_gen_t4['EQUIPMENT'].apply(extract_substation)
        
        
    # itoa_gen_t4['Substation'] = ''
    itoa_gen_t4['Email_Date'] = formatted_today
    itoa_gen_t4['NotfCreatedBy'] = 'ILS_Automation'

    itoa_gen_t4['event_activity_comments'] = itoa_gen_t4['event_activity_comments'].fillna('')
    itoa_gen_t4['LONGTEXT1'] = itoa_gen_t4['iTOA_link'] + ' ; ' + itoa_gen_t4['event_activity_comments']

    date_str = datetime.now().strftime("%Y_%m_%d")

    itoa_gen_t4 = itoa_gen_t4.rename(columns = {"EQUIPMENT" : "NotfCircuitName",
                                                "VOLTAGE_LEVEL" : "Voltage",
                                                "Engineer Assigned": "Engineer_Assigned",
                                                "WORK_CENTERS": "Switching_Center",
                                                "HFRA_UDF": "HFRA_CIRCUIT?",
                                                "Out_Date": "Notf_Date"})


    # print(itoa_gen_t4.head())
    # itoa_gen_t4.to_excel('output.xlsx')
    # sys.exit()
# =============================================================================
# Change timestamp to account for UTC in PowerAutomate
# =============================================================================

    # itoa_gen_t4['Notf_Date_adjust'] = pd.to_datetime(itoa_gen_t4['Notf_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    # itoa_gen_t4['Notf_Date'] = pd.to_datetime(itoa_gen_t4['Notf_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    # itoa_gen_t4['Notf_Date'] = itoa_gen_t4['Notf_Date'].str.strip()

    # itoa_gen_t4['Email_Date_adjust'] = pd.to_datetime(itoa_gen_t4['Email_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    # itoa_gen_t4['Email_Date'] = pd.to_datetime(itoa_gen_t4['Email_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    # itoa_gen_t4['Email_Date'] = itoa_gen_t4['Email_Date'].str.strip()

# =============================================================================
# Upload iTOA data directly to the Repair Order Events SharePoint list
# =============================================================================
    sharepoint_site = 'https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis'
    sharepoint_list_name = 'Repair Order Events'
    sharepoint_field_mapping = {
        'Subject': 'Title',
        'NotfCircuitName': 'Circuit',
        'Entry_Method': 'Entry_Method',
        'Email_Date': 'Email_x0020_Date',
        'Notf_Date': 'Incident_x0020_Date',
        'NotfCreatedBy': 'Reported_x0020_By_x0020_Email',
        'LONGTEXT1': 'Comments',
        'HFRA_CIRCUIT?': 'HFRA_x0020_Circuit',
        'Switching_Center': 'Switching_x0020_Center',
        'Voltage': 'Voltage_x0020__x0028_kV_x0029_',
        'Substation': 'Substation',
        'Engineer_Assigned': 'Engineer_x0020_Assigned',
    }
    sharepoint_upload_columns = list(sharepoint_field_mapping)
    sharepoint_upload_data = itoa_gen_t4[sharepoint_upload_columns]
    authenticate_user()
    upload_dataframe_to_sharepoint_list(
        sharepoint_site=sharepoint_site,
        sharepoint_list_name=sharepoint_list_name,
        dataframe=sharepoint_upload_data,
        sharepoint_field_mapping=sharepoint_field_mapping,
    )

# =============================================================================
# Write to previous run to not get duplicates and base pull off of last run
# =============================================================================
    output_filepath = 'pipelines/transmission/previous_runs_itoa/iTOA_Transmission_Events_' + date_str + ".xlsx"
    itoa_gen_t4.to_excel(output_filepath, index=False)
        
else:
    print("No new iTOA Events to process")
