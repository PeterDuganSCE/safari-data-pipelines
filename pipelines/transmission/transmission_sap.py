import logging
import os
from datetime import datetime
import datetime as dt
import glob
from openpyxl import load_workbook
import saspy
import pandas as pd
import sqlalchemy as sa
from urllib.parse import urlparse
from pathlib import Path
import sys
import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.utils import district_inboxes, system_assignments
from shared.sharepoint.graph_api import authenticate_user, pull_sharepoint_list_data
from shared.sharepoint.utils import upload_dataframe_to_sharepoint_list
from shared.paths import AUTH_PATH


logger = logging.Logger('TransmissionNotif')
logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(
    filename='logs/sap_trans_run.log',
    mode='w'
)

file_handler.setLevel(logging.INFO)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

formatter = logging.Formatter(
    fmt='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# =============================================================================
# Need to run info for 2 days prior since notifications loaded at 6pm, so have to make sure all created on dates give the chance for info to get into system
# =============================================================================
# today = datetime.now().strftime("%m/%d/%Y")
today_str = datetime.today().strftime(" %d%b%Y")
today = datetime.strptime(today_str, " %d%b%Y")

#Get date of last run to pull data from days that it didnt run
previous_runs_dir = Path(__file__).resolve().parent / "previous_runs_sap"
list_of_files = glob.glob(str(previous_runs_dir / "*"))
if not list_of_files:
    raise FileNotFoundError(f"No previous run files found in {previous_runs_dir}")
latest_file = max(list_of_files, key=os.path.getctime)


t = latest_file.split("Transmission_Events_")[1]
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
# Code for SAS query of Transmission events
# =============================================================================
sas_trans_e1_query = """
PROC SQL;
   CREATE TABLE WORK.TRANS_E1P1 AS 
   SELECT DISTINCT t1.NotfCircuitName, 
          t1.Floc, 
          t1.Notification, 
          t1.NotfType, 
          t1.Priority, 
          t1.Notf_Date,
          t1.Notf_Createdon_Date,
          t1.Problem_Stat, 
          t1.Notf_MAT, 
          t1.MAT_Desc, 
          t1.NotfPatrolTypeID, 
          t1.NotfPatrolType, 
          t1.Level_4, 
          t1.Level_3, 
          t1.Level_2, 
          t1.Level_1, 
          t1.Tier_Level, 
          t1.NotfCreatedBy, 
          t1.Sys_Status_All, 
          t1.Sys_Status
      FROM TDINTGRN.NOTIF_PRIORITY t1
      WHERE t1.Level_3 = 'Transmission' AND t1.Notf_Createdon_Date >= '""" + full_day_str + """'d AND t1.Notf_Createdon_Date < '""" + today_str + """'d AND t1.NotfType = 'E1' AND t1.Priority = '1' AND 
           t1.Sys_Status NOT IN 
           (
           'OSNO',
           'DLFL'
           )
      ORDER BY t1.Notf_Date,
               t1.Notf_Time;
QUIT;

PROC SQL;
   CREATE TABLE WORK.NOTIF_LONGTEXT AS 
   SELECT t1.QMNUM, t1.LONGTEXT1
      FROM TDSTAGNG.DAILYLT t1
      WHERE t1.ERDAT >= '1Jan2024'd;
QUIT;

PROC SQL;
   CREATE TABLE WORK.TRANS_NOTIF_LT AS 
   SELECT t1.NotfCircuitName, 
          t1.Floc, 
          t1.Notification, 
          t1.NotfType, 
          t1.Priority, 
          t1.Notf_Createdon_Date,
          t1.Notf_Date, 
          t1.Problem_Stat, 
          t2.LONGTEXT1, 
          t1.Notf_MAT, 
          t1.MAT_Desc, 
          t1.NotfPatrolTypeID, 
          t1.NotfPatrolType, 
          t1.Level_4, 
          t1.Level_3, 
          t1.Level_2, 
          t1.Level_1, 
          t1.Tier_Level, 
          t1.NotfCreatedBy, 
          t1.Sys_Status_All, 
          t1.Sys_Status
      FROM WORK.TRANS_E1P1 t1
           LEFT JOIN WORK.NOTIF_LONGTEXT t2 ON (t1.Notification = t2.QMNUM);
QUIT;

PROC SQL;
   CREATE TABLE WORK.TRANS_NOTIF_LT_DIS AS 
   SELECT t1.NotfCircuitName, 
          t1.Floc, 
		  t2.FLOC_DistrictID,
		  t2.District_Name,
          t1.Notification, 
          t1.NotfType, 
          t1.Priority, 
          t1.Notf_Createdon_Date,
          t1.Notf_Date, 
          t1.Problem_Stat, 
          t1.LONGTEXT1, 
          t1.Notf_MAT, 
          t1.MAT_Desc, 
          t1.NotfPatrolTypeID, 
          t1.NotfPatrolType, 
          t1.Level_4, 
          t1.Level_3, 
          t1.Level_2, 
          t1.Level_1, 
          t1.Tier_Level, 
          t1.NotfCreatedBy, 
          t1.Sys_Status_All, 
          t1.Sys_Status
      FROM WORK.TRANS_NOTIF_LT t1
           LEFT JOIN TDINTGRN.INT_EQFLOCMASTER t2 ON (t1.Floc = t2.FLOC_FunctionalLocationID);
QUIT;

PROC SQL;
CREATE TABLE WORK.SWITCH_CENTER AS
SELECT
t1.FLOC_FunctionalLocation,
t1.FLOC_SwitchingCenter
FROM TDSTAGNG.STG_FLOCMASTER t1
WHERE t1.FLOC_TechObjectType LIKE '%CKTS'
AND t1.FLOC_StructureID LIKE 'ET-%' ;
QUIT;

PROC SQL;
   CREATE TABLE WORK.TRANS_NOTIF_LT_DIS_2 AS 
   SELECT t1.NotfCircuitName, 
          t1.Floc, 
		  t1.FLOC_DistrictID,
		  t1.District_Name,
          t2.FLOC_SwitchingCenter,
          t1.Notification, 
          t1.NotfType, 
          t1.Priority, 
          t1.Notf_Createdon_Date,
          t1.Notf_Date, 
          t1.Problem_Stat, 
          t1.LONGTEXT1, 
          t1.Notf_MAT, 
          t1.MAT_Desc, 
          t1.NotfPatrolTypeID, 
          t1.NotfPatrolType, 
          t1.Level_4, 
          t1.Level_3, 
          t1.Level_2, 
          t1.Level_1, 
          t1.Tier_Level, 
          t1.NotfCreatedBy, 
          t1.Sys_Status_All, 
          t1.Sys_Status
      FROM WORK.TRANS_NOTIF_LT_DIS t1
           LEFT JOIN WORK.SWITCH_CENTER t2 ON (t1.NotfCircuitName = t2.FLOC_FunctionalLocation);
QUIT;
"""


try:

    sas = saspy.SASsession(
        cfgfile='config/sascfg_personal.py',
        cfgname='iomcom',
    )

    logger.info(f'SAS Connection established with {sas.sasver}.')
    logger.info(f'SAS server host: {sas.symget("SYSHOSTNAME")} (config iomhost: {getattr(sas.sascfg, "iomhost", "n/a")})')
    logger.info(f'Assigned librefs: {sas.assigned_librefs()}')

    c = sas.submitLST(sas_trans_e1_query) # uncomment this line to run the SAS query
    print(sas.lastlog())
    
    df = sas.sasdata2dataframe(
        table='TRANS_NOTIF_LT_DIS_2',
        libref='WORK',
        method='CSV'
    )
    print(sas.lastlog())
    sas.endsas()
    
    date_str = datetime.now().strftime("%Y_%m_%d")

    df2 = df.drop_duplicates()
    
    if len(df2) > 0:
        df2 = df2[~df2['Notification'].isin(prev_run_file['Notification'])]
    
        df2['Notification'] = df2['Notification'].astype(str)
    
        df2['Subject'] = 'Trans_SAP_' + df2['Notification']
        df2['Entry_Method'] = 'SAP'

#Substation Name, Voltage, HFRA_Circuit, Switching Center
        trans_ckt_info = pd.read_excel(r'data\HFRA Transmission and Subtransmission Master Circuit List.xlsx')
    
        trans_ckt_info_sub = trans_ckt_info[['Circuit Name', 'Voltage (kV)', 'HFRA CIRCUIT?', 'Switching Center', 'FLOC']]
    
        df2['NotfCircuitName'] = df2['NotfCircuitName'].fillna('')
    
        df2['NotfCircuitName'] = df2['NotfCircuitName'].str.upper()
    
        trans_ckt_info_sub['Circuit Name'] = trans_ckt_info_sub['Circuit Name'].astype(str)
    
        df3 = pd.merge(df2, trans_ckt_info_sub, left_on= ['NotfCircuitName'],
                         right_on= ['Circuit Name'], how='left').drop(['Circuit Name'], axis=1)
    
        df3['District'] = df3['FLOC_DistrictID'].str.replace('ED', '', regex=True)
    
# ============================================================================
#Fill in Switching Center where it can be filled in for notifications without Circuits as well as assignment to Engineers
# =============================================================================
    #Engineer Assignment
        authenticate_user()
        sys_assign = system_assignments()
        # sys_assign = pull_sharepoint_list_data('https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis', 'System Assignments')
        sys_dist_mapping = pd.read_excel('data/SWCNTR_DISTRICT_mapping.xlsx')
        sys_dist_mapping = sys_dist_mapping[['SWITCHING_CENTER', 'DISTRICT_NUM']]
        sys_dist_mapping['DISTRICT_NUM'] = sys_dist_mapping['DISTRICT_NUM'].astype(str)
    
#District Num to District Name, Engineer Assigned (Old mapping before going by system)
        dist_inbox = district_inboxes()
        # dist_inbox = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/District Inboxes.xlsx')
        # dist_inbox = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/District Inboxes.xlsx')
    
        dist_inbox_sub = dist_inbox[['District Number', 'District Name']]
    
        dist_inbox_sub['District Number'] = dist_inbox_sub['District Number'].astype(str)
    
        df4 = pd.merge(df3, sys_dist_mapping, left_on= ['District'],
                         right_on= ['DISTRICT_NUM'], how='left').drop(['DISTRICT_NUM'], axis=1)
    
    #Take Notification swithching center if there but take district system mapping if nan
    
        df4['FLOC_SwitchingCenter'] = df4['FLOC_SwitchingCenter'].fillna('')
    
        def switch_center(not_switch, dist_switch):
            if not_switch == '' or not_switch == 'Missing' or not_switch == pd.isna(not_switch):
                return dist_switch
            else:
                return not_switch

        df4['Switching_Center_UDF'] = df4.apply(lambda x: switch_center(x['FLOC_SwitchingCenter'], x['SWITCHING_CENTER']), axis=1)
    
        del df4['Switching Center']
    
    
        df5 = pd.merge(df4, sys_assign, left_on= ['Switching_Center_UDF'],
                         right_on= ['Switching Center'], how='left').drop(['Switching Center'], axis=1)
    
        df5['Structure'] = df5['Floc'].str.replace('OH-', '', regex=True)
        df5['Structure'] = df5['Structure'].str.replace('UG-', '', regex=True)

        df5['Structure Type'] = df5['Structure'].str.extract(r'^(UG|OH)').fillna('')
        
        def extract_substation(name):
            if pd.isna(name) or name.strip() == "":
                return ""
            else:
        # Split by the dash and return the first part
                return name.split('-')[0].strip()

# Create a new column 'Substation' using apply
        df5['Substation'] = df5['NotfCircuitName'].apply(extract_substation)
        # df5['Substation'] = ''
    
        def hftd(Tier):
            if Tier == 'Elevated':
                return 'T2'
            elif Tier == 'Extreme':
                return 'T3'
            else:
                return 'NON-HFRA'

        df5['HFRA_Ckt'] = df5.apply(lambda x: hftd(x['Tier_Level']), axis=1)
    
        df6 = pd.merge(df5, dist_inbox_sub, left_on= ['District'],
                         right_on= ['District Number'], how='left').drop(['District Number'], axis=1)
    
        df6['Voltage (kV)'] = df6['Voltage (kV)'].fillna(0).astype(int).astype(str).replace('0', '')
    
# =============================================================================
#Get date from Long text and use for Outage Datetime and Format date time for PowerApps
# =============================================================================
    
        df6['OutageDT'] = df6['LONGTEXT1'].apply(lambda x: x.split(" PST")[0])
        df6['OutageDT'] = df6['OutageDT'].apply(lambda x: x.split("* ")[1])
    
        df6['OutageDT_adjust'] = pd.to_datetime(df6['OutageDT']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df6['OutageDT'] = pd.to_datetime(df6['OutageDT_adjust']).dt.strftime('%m/%d/%Y %H:%M')
        df6['OutageDT'] = df6['OutageDT'].str.strip()

        df6['Notf_Createdon_Date_adjust'] = pd.to_datetime(df6['Notf_Createdon_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df6['Notf_Createdon_Date'] = pd.to_datetime(df6['Notf_Createdon_Date_adjust']).dt.strftime('%m/%d/%Y')
        df6['Notf_Createdon_Date'] = df6['Notf_Createdon_Date'].str.strip()
    
        df6['Notf_Date_adjust'] = pd.to_datetime(df6['Notf_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df6['Notf_Date'] = pd.to_datetime(df6['Notf_Date_adjust']).dt.strftime('%m/%d/%Y')
        df6['Notf_Date'] = df6['Notf_Date'].str.strip()

        df6 = df6.drop_duplicates()
        
        df6['LONGTEXT2'] = 'Notif: ' + df6['Notification'] + df6['LONGTEXT1']
    
    
# =============================================================================
# Upload SAP data directly to the Repair Order Events SharePoint list
# =============================================================================
        sharepoint_site = 'https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis'
        sharepoint_list_name = 'Repair Order Events'
        sharepoint_field_mapping = {
            'Subject': 'Title',
            'NotfCircuitName': 'Circuit',
            'Entry_Method': 'Entry_Method',
            'Notf_Createdon_Date': 'Email_x0020_Date',
            'OutageDT': 'Incident_x0020_Date',
            'NotfCreatedBy': 'Reported_x0020_By_x0020_Email',
            'LONGTEXT2': 'Comments',
            'HFRA_Ckt': 'HFRA_x0020_Circuit',
            'Switching_Center_UDF': 'Switching_x0020_Center',
            'Voltage (kV)': 'Voltage_x0020__x0028_kV_x0029_',
            'Substation': 'Substation',
            'Engineer Assigned': 'Engineer_x0020_Assigned',
            'District': 'District_x0020_Number',
            'District Name': 'District_x0020_Name',
            'Structure': 'Structure',
            'Structure Type': 'Structure_x0020_Type',
        }
        sharepoint_upload_columns = list(sharepoint_field_mapping)
        sharepoint_upload_data = df6[sharepoint_upload_columns].copy()
        for column in ['Voltage (kV)', 'District']:
            sharepoint_upload_data[column] = pd.to_numeric(
                sharepoint_upload_data[column].replace('', None)
            )
        sharepoint_upload_data['Structure Type'] = sharepoint_upload_data['Structure Type'].replace('', None)
        sharepoint_upload_data['HFRA_Ckt'] = sharepoint_upload_data['HFRA_Ckt'].replace('NON-HFRA', 'Non-HFRA')
        sharepoint_upload_data['Notf_Createdon_Date'] = pd.to_datetime(
            df6['Notf_Createdon_Date_adjust'], format='%m/%d/%Y %H:%M', utc=True
        )
        sharepoint_upload_data['OutageDT'] = pd.to_datetime(
            sharepoint_upload_data['OutageDT'], format='%m/%d/%Y %H:%M', utc=True
        )
        authenticate_user()
        upload_dataframe_to_sharepoint_list(
            sharepoint_site=sharepoint_site,
            sharepoint_list_name=sharepoint_list_name,
            dataframe=sharepoint_upload_data,
            sharepoint_field_mapping=sharepoint_field_mapping,
        )

        logger.info(f"Uploaded {len(sharepoint_upload_data)} records to SharePoint completed.")

        output_filepath = 'pipelines/transmission/previous_runs_sap/Transmission_Events_' + date_str + ".xlsx"
        df6.to_excel(output_filepath, index=False)

        logger.info(f"Created file -> {output_filepath}")

    else:
        print("No new SAP notifications")

except Exception as e:
    logger.error(e)
finally:
    logger.info('SAS Connection terminated.')
