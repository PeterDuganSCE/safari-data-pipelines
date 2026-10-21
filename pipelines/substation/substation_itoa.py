# -*- coding: utf-8 -*-
"""
Created on Fri Jul 26 06:26:55 2024

@author: brownsjm
"""
import pandas as pd
import warnings
import os
import pyodbc
import saspy
from sqlalchemy import create_engine
import datetime as dt
import datetime
from datetime import datetime
import glob
from openpyxl import load_workbook
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.sharepoint.graph_api import authenticate_user
from shared.sharepoint.utils import upload_dataframe_to_sharepoint_list
from shared.utils import system_assignments
from shared.logging import setup_logging

logger = setup_logging('substation_itoa')


def _log_unhandled_exception(exception_type, exception_value, exception_traceback):
    logger.critical(
        'Substation iTOA pipeline failed.',
        exc_info=(exception_type, exception_value, exception_traceback),
    )


sys.excepthook = _log_unhandled_exception
logger.info('Starting substation iTOA pipeline.')

warnings.filterwarnings('ignore')
pd.set_option('display.width', 500)
pd.set_option('display.max_column', 500)
pd.set_option('display.max_rows', 200)
pd.set_option('display.max_colwidth', 500)

#pri_workcenter field in iTOA tables (GENERIC)

#%%Load SQL data for iTOA

# SQL Server connection info
DB_DRIVER = "{SQL Server}" # Or whatever driver is installed in your machine
DB_SERVER = 'tcp:AYWCPSQL116\SQLP740,1989'  # Irvine Server
# DB_SERVER_2 = 'tcp:AYWCPSQL651\SQLP648,1989'  # Alhambra Server
DATABASE_NAME = 'ITOA'  # development server, for testing
logger.info('Connecting to iTOA database %s on %s.', DATABASE_NAME, DB_SERVER)
connection_itoa = pyodbc.connect(f"DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DATABASE_NAME};Trusted_Connection=no;")
logger.info('iTOA database connection established.')


itoa_gen_q = '''SELECT DISTINCT APP_ID, LINE_OF_BUSINESS, STATION, APPLICATION_STATUS, EQUIPMENT, VOLTAGE_LEVEL, MAIN_ACTUAL_OUT, DISTURBANCE_DURATION, DISTURBANCE_NO, AFFECTED_DISTRICT, WORK_CENTERS, WEATHER, MISCELLANEOUS_COMMENTS, event_activity_comments, INITIATING_CAUSE_CODE, INITIATING_SUB_CAUSE, CAUSE_CATEGORY, OUTAGE_CATEGORY, PROCESS_STATUS FROM dbo.GENERIC_INT_TADS_V'''
logger.info('Reading iTOA outage records.')
itoa_gen = pd.read_sql(itoa_gen_q, connection_itoa)
logger.info('Loaded %s iTOA records.', len(itoa_gen))

itoa_gen = itoa_gen.drop_duplicates()
logger.info('Retained %s records after removing duplicates.', len(itoa_gen))

itoa_gen_t = itoa_gen[itoa_gen['LINE_OF_BUSINESS'] == 'Substation']
itoa_gen_t = itoa_gen_t[itoa_gen_t['PROCESS_STATUS'] != 'VOID']
logger.info('Retained %s non-VOID substation records.', len(itoa_gen_t))

del itoa_gen_t['PROCESS_STATUS']

itoa_gen_t['iTOA_link'] = itoa_gen_t['APP_ID'].apply(lambda x: f'https://orls.sce.com/itoa/autooutage/view.htmlx?editedOutage.appId={x}&eventAnalysis.analysisId=&referer=interruption')

# =============================================================================
# Need to run info for 2 days prior since notifications loaded at 6pm, so have to make sure all created on dates give the chance for info to get into system
# =============================================================================
# today = datetime.now().strftime("%m/%d/%Y")
today_str = datetime.today().strftime(" %d%b%Y")
today = datetime.strptime(today_str, " %d%b%Y")

#Get date of last run to pull data from days that it didnt run
previous_runs_dir = PROJECT_ROOT / 'pipelines' / 'substation' / 'previous_runs_itoa'
list_of_files = glob.glob(str(previous_runs_dir / '*'))
if not list_of_files:
    raise FileNotFoundError(f'No previous substation iTOA run workbooks were found in {previous_runs_dir}.')
latest_file = max(list_of_files, key=os.path.getctime)
logger.info('Loading previous run workbook: %s', latest_file)


t = latest_file.split("iTOA_Substation_Events_")[1]
t = t.split(".")[0]
t = t.replace("_", "/")
last_date = pd.to_datetime(t)
last_date_str = last_date.strftime(" %d%b%Y")

# Load previous file to prevent duplicates when running the extra day for 6pm to midnight data
prev_run_file = pd.read_excel(latest_file)
logger.info('Loaded %s records from the previous run workbook.', len(prev_run_file))

# One delay to take notifications back another day in case events came in after 6pm that day (Loaded data at 6pm so cutoff in SAP)
full_day = last_date - dt.timedelta(days=1)
full_day_str = full_day.strftime(" %d%b%Y")

yesterday = today - dt.timedelta(days=2)
yesterday_str = yesterday.strftime(" %d%b%Y")

# =============================================================================
# %% Limit to last day
# =============================================================================
itoa_gen_t['Out_Date'] = pd.to_datetime(itoa_gen_t['MAIN_ACTUAL_OUT'], errors='coerce')
invalid_date_count = int(itoa_gen_t['Out_Date'].isna().sum())
if invalid_date_count:
    logger.warning('Excluding %s records with missing or invalid outage dates.', invalid_date_count)


# today = dt.datetime.today()
# yesterday = today - dt.timedelta(days=1)
# today = today.strftime('%m/%d/%Y')
# formatted_today = dt.datetime.strptime(today, '%m/%d/%Y')

# yesterday = yesterday.strftime('%m/%d/%Y')
itoa_gen_t_2 = itoa_gen_t[itoa_gen_t['Out_Date'] >= full_day]
itoa_gen_t_2 = itoa_gen_t_2[itoa_gen_t_2['Out_Date'] < today]
logger.info(
    'Selected %s records for outage dates from %s (inclusive) to %s (exclusive).',
    len(itoa_gen_t_2), full_day.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'),
)

itoa_gen_t_2['STATION'] = itoa_gen_t_2['STATION'].str.upper()


if len(itoa_gen_t_2) > 0:
    logger.info('Enriching substation records with HFRA and engineer assignments.')
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
        
#Remove space and voltage from Equipment(Ckt name)
    itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3.apply(
        lambda row: row['EQUIPMENT'].replace(f" {row['VOLTAGE_LEVEL']}", "").strip() 
                 if f" {row['VOLTAGE_LEVEL']}" in row['EQUIPMENT'] 
                 else row['EQUIPMENT'], 
    axis=1
    )
else: itoa_gen_t_3 = []
    
# =============================================================================
# No circuit information as this is substation events but can take first Trans line as Circuit
# Get HFRA Tier for where the substation is location
# =============================================================================
if len(itoa_gen_t_3) > 0:
    sas = saspy.SASsession(
        cfgfile=str(PROJECT_ROOT / 'config' / 'sascfg_personal.py'),
        cfgname='iomcom', m5dsbug = [True])

    logger.info(f'SAS Connection established with {sas.sasver}.')
    logger.info(f'SAS server host: {sas.symget("SYSHOSTNAME")} (config iomhost: {getattr(sas.sascfg, "iomhost", "n/a")})')
    logger.info(f'Assigned librefs: {sas.assigned_librefs()}')
#print(sas.lastlog())
#Write ro_inc to SAS to use in query; Saves a lot of time of reading out larger list from SAS
    # TMP_INC = sas.df2sd(itoa_gen_t_3_for_sas, 'event_for_sas')
    # print(sas.lastlog())
#print(sas.lastlog())

    # c = sas.submitLST("""PROC SQL;
    #                       CREATE TABLE WORK.LAT_LON AS
    #                       SELECT DISTINCT *
    #                       FROM TDINTGRN.INT_EQFLOCMASTER;
    #                       QUIT;
    #                       """)
                          
    c = sas.submitLST("""PROC SQL;
                          CREATE TABLE WORK.SUB AS
                          SELECT DISTINCT *
                          FROM TDSTAGNG.SUBSTATION_HFRA;
                          QUIT;
                          """)
    logger.info('Completed SAS substation HFRA query.')

    sub_hfra = sas.sasdata2dataframe('SUB', libref='WORK', method='CSV')
    logger.info('Loaded %s substation HFRA records.', len(sub_hfra))
    
    itoa_gen_t_3['STATION_U'] = itoa_gen_t_3['STATION'].str.upper()
    
    sub_hfra['Substation'] = sub_hfra['Substation'].str.upper()

    itoa_gen_t3 = pd.merge(itoa_gen_t_3, sub_hfra, left_on= ['STATION_U'],
                          right_on= ['Substation'], how='left').drop(['Substation'], axis=1)
    
    
    def hfra_check(tier): # decide highest level Tier
        if pd.isna(tier):
            return 'Non HFRA'
        elif 'Extreme' in tier:
            return 'T3'
        elif 'Elevated' in tier:
            return 'T2'
        else:
            return 'NON-HFRA'

    itoa_gen_t3['HFRA_UDF'] = itoa_gen_t3.apply(lambda x: hfra_check(x['FLOC_Tier_Level']), axis=1)

    # trans_ckt_info = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\Jbrow\PowerAppAutomation\HFRA Transmission and Subtransmission Master Circuit List.xlsx')
    
    # trans_ckt_info_sub = trans_ckt_info[['Circuit Name', 'HFRA TIERS']]
    
    # itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].fillna('')
    
    # itoa_gen_t_3['EQUIPMENT'] = itoa_gen_t_3['EQUIPMENT'].str.upper()
    
    # trans_ckt_info_sub['Circuit Name'] = trans_ckt_info_sub['Circuit Name'].astype(str)
    
    # itoa_gen_t3 = pd.merge(itoa_gen_t_3, trans_ckt_info_sub, left_on= ['EQUIPMENT'],
    #                      right_on= ['Circuit Name'], how='left').drop(['Circuit Name'], axis=1)

    # itoa_gen_t3['HFRA TIERS'] = itoa_gen_t3['HFRA TIERS'].fillna('Non HFRA')

    # def hfra_check(tier): # decide highest level Tier
    #     if 'T3' in tier:
    #         return 'T3'
    #     elif 'T2' in tier:
    #         return 'T2'
    #     else:
    #         return 'Non HFRA'

    # itoa_gen_t3['HFRA_UDF'] = itoa_gen_t3.apply(lambda x: hfra_check(x['HFRA TIERS']), axis=1)
    
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
    logger.info('Loading SharePoint system assignments.')
    sys_assign = system_assignments()
    logger.info('Loaded %s system assignment records.', len(sys_assign))

    itoa_gen_t3['WORK_CENTERS'] = itoa_gen_t3['WORK_CENTERS'].str.upper()

    itoa_gen_t4 = pd.merge(itoa_gen_t3, sys_assign, left_on= ['WORK_CENTERS'],
                         right_on= ['Switching Center'], how='left').drop(['Switching Center'], axis=1)

    itoa_gen_t4['Engineer Assigned'] = itoa_gen_t4['Engineer Assigned'].fillna('Unassigned')
    unassigned_count = int(itoa_gen_t4['Engineer Assigned'].eq('Unassigned').sum())
    if unassigned_count:
        logger.warning('%s records have no assigned engineer.', unassigned_count)
    

    itoa_gen_t5 = (
        itoa_gen_t4
        .groupby('APP_ID', as_index=False)
        .agg('first')
)

    
#Add columns for Engineers to track comments and review times (Old before automation)
# itoa_gen_t['Engineer_Assigned'] = ''
# itoa_gen_t['SAP_Notif'] = ''
# itoa_gen_t['SAP_Notif_Num'] = ''
# itoa_gen_t['Desktop_Review_Time'] = ''
# itoa_gen_t['Days_for_Sr_Patrol_Info'] = ''
# itoa_gen_t['Engineer_Comments'] = ''
    itoa_gen_t5['Subject'] = 'Sub_itoa_' + itoa_gen_t5['APP_ID'].astype(str)
    itoa_gen_t5['Entry_Method'] = 'ILS'
    
    
    # def extract_substation(name):
    #     if pd.isna(name) or name.strip() == "":
    #         return ""
    #     else:
    #     # Split by the dash and return the first part
    #         return name.split('-')[0].strip()

# Create a new column 'Substation' using apply
    itoa_gen_t5['Substation'] = itoa_gen_t5['STATION']
        
        
    # itoa_gen_t4['Substation'] = ''
    itoa_gen_t5['Email_Date'] = formatted_today
    itoa_gen_t5['NotfCreatedBy'] = 'ILS_Automation'

    itoa_gen_t5['event_activity_comments'] = itoa_gen_t5['event_activity_comments'].fillna('')
    itoa_gen_t5['LONGTEXT1'] = itoa_gen_t5['iTOA_link'] + ' ; ' + itoa_gen_t5['event_activity_comments']

    date_str = datetime.now().strftime("%Y_%m_%d")

    itoa_gen_t5 = itoa_gen_t5.rename(
        columns = {
            "EQUIPMENT" : "NotfCircuitName",
            "VOLTAGE_LEVEL" : "Voltage",
            "Engineer Assigned": "Engineer_Assigned",
            "WORK_CENTERS": "Switching_Center",
            "HFRA_UDF": "HFRA_CIRCUIT?",
            "Out_Date": "Notf_Date"
        }
    )
    
    

# =============================================================================
# Change timestamp to account for UTC in PowerAutomate
# =============================================================================

    itoa_gen_t5['Notf_Date_adjust'] = pd.to_datetime(itoa_gen_t5['Notf_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t5['Notf_Date'] = pd.to_datetime(itoa_gen_t5['Notf_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t5['Notf_Date'] = itoa_gen_t5['Notf_Date'].str.strip()

    itoa_gen_t5['Email_Date_adjust'] = pd.to_datetime(itoa_gen_t5['Email_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t5['Email_Date'] = pd.to_datetime(itoa_gen_t5['Email_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t5['Email_Date'] = itoa_gen_t5['Email_Date'].str.strip()

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
    sharepoint_upload_data = itoa_gen_t5[sharepoint_upload_columns]
    logger.info('Authenticating to SharePoint for the substation upload.')
    authenticate_user()
    logger.info('Uploading %s records to SharePoint list %s.', len(sharepoint_upload_data), sharepoint_list_name)
    uploaded_count = upload_dataframe_to_sharepoint_list(
        sharepoint_site=sharepoint_site,
        sharepoint_list_name=sharepoint_list_name,
        dataframe=sharepoint_upload_data,
        sharepoint_field_mapping=sharepoint_field_mapping,
    )
    logger.info('Uploaded %s records to SharePoint list %s.', uploaded_count, sharepoint_list_name)
    

# =============================================================================
# Write to previous run to not get duplicates and base pull off of last run
# =============================================================================
    output_filepath = previous_runs_dir / ('iTOA_Substation_Events_' + date_str + '.xlsx')
    logger.info('Saving run-history workbook: %s', output_filepath)
    itoa_gen_t5.to_excel(output_filepath, index=False)
    logger.info('Saved %s records to %s.', len(itoa_gen_t5), output_filepath)
        
else:
    logger.info('No new iTOA events to process.')

#read in existing master list
# master = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\FIPA\Transmission_Events_test\01_Transmission_Events_Master.xlsx')

# if len(itoa_gen_t) > 0:
#     master_2 = master.append(itoa_gen_t, ignore_index = True, sort = False)

#     master_2.to_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\FIPA\Transmission_Events_test\01_Transmission_Events_Master.xlsx', index = False)

# #Older write per day, currently replacing with master list
#     write_location = '//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/Transmission_Events_test/iTOA_Transmission_Events_' + month_number + '_' + day_number + '_' + year_number +'.xlsx'

#     itoa_gen_t.to_excel(write_location, index = False)
#===============================================

logger.info('Substation iTOA pipeline completed.')
