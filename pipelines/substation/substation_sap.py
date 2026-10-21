import logging
import os
import sys
from datetime import datetime
import datetime as dt
import glob
import tempfile
from pathlib import Path

from openpyxl import load_workbook
import saspy
import pandas as pd
import numpy as np
import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.paths import AUTH_PATH
from shared.utils import system_assignments, district_inboxes
from shared.sharepoint.graph_api import authenticate_user
from shared.sharepoint.utils import download_file, upload_file, upload_dataframe_to_sharepoint_list


logger = logging.Logger('SubstationNotif')
logger.setLevel(logging.INFO)

file_handler = logging.FileHandler(
    filename='logs/substation_sap.log',
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
logger.info('Starting substation SAP pipeline.')

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
logger.info('Loading previous run workbook: %s', latest_file)


t = latest_file.split("Substation_Events_")[1]
t = t.split(".")[0]
t = t.replace("_", "/")
last_date = pd.to_datetime(t)
last_date_str = last_date.strftime(" %d%b%Y")

#Load previous file to prevent duplicates when running the extra day for 6pm to midnight data
prev_run_file = pd.read_excel(latest_file)
logger.info('Loaded %s records from the previous run workbook.', len(prev_run_file))

#One delay to take notifications back another day in case events came in after 6pm that day (Loaded data at 6pm so cutoff in SAP)
full_day = last_date - dt.timedelta(days=1)
full_day_str = full_day.strftime(" %d%b%Y")

yesterday = today - dt.timedelta(days=2)
yesterday_str = yesterday.strftime(" %d%b%Y")
# =============================================================================
# Code for SAS query of Transmission events
# =============================================================================
sas_sub_e1_query = """
PROC SQL;
   CREATE TABLE WORK.SUB_E1P1 AS 
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
      WHERE t1.Level_3 = 'Substation' AND t1.Notf_Createdon_Date >= '""" + full_day_str + """'d AND t1.Notf_Createdon_Date < '""" + today_str + """'d AND t1.NotfType = 'E1' AND t1.Priority = '1' AND 
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
   CREATE TABLE WORK.SUB_NOTIF_LT AS 
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
      FROM WORK.SUB_E1P1 t1
           LEFT JOIN WORK.NOTIF_LONGTEXT t2 ON (t1.Notification = t2.QMNUM);
QUIT;

PROC SQL;
   CREATE TABLE WORK.SUB_NOTIF_LT_DIS AS 
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
      FROM WORK.SUB_NOTIF_LT t1
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
   CREATE TABLE WORK.SUB_NOTIF_LT_DIS_2 AS 
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
      FROM WORK.SUB_NOTIF_LT_DIS t1
           LEFT JOIN WORK.SWITCH_CENTER t2 ON (t1.NotfCircuitName = t2.FLOC_FunctionalLocation);
QUIT;
"""
sas_ckt_sub_query = """
PROC SQL;
   CREATE TABLE WORK.SUB_NUM AS 
   SELECT DISTINCT
   t1.SUB_NO,
   t1.SUBSTATION,
   t1.SWTCH_CNTR_DESC,
   t1.DISTRICT_NO, t1.DISTRICT_NAME
   FROM TDSTAGNG.CIRCUITSUBSTAT t1;
   QUIT;
"""

# =============================================================================
# Run code
# =============================================================================
with open(AUTH_PATH, "r", encoding="utf-8") as auth_file:
    sas_credentials = (yaml.safe_load(auth_file) or {})["sas"]

try:
    sas = saspy.SASsession(
        cfgfile=str(PROJECT_ROOT / 'config' / 'sascfg_personal.py'),
        cfgname='iomcom',
    )

    logger.info(f'SAS Connection established with {sas.sasver}.')
    logger.info(f'SAS server host: {sas.symget("SYSHOSTNAME")} (config iomhost: {getattr(sas.sascfg, "iomhost", "n/a")})')
    # logger.info(f'Assigned librefs: {sas.assigned_librefs()}')

    logger.info('Running SAS substation notification and long-text query.')
    c = sas.submitLST(sas_sub_e1_query)
    # print(sas.lastlog())
    
    logger.info('Loading substation notification records from SAS.')
    df = sas.sasdata2dataframe(
        table='SUB_NOTIF_LT_DIS_2',
        libref='WORK',
        method='CSV'
    )
    # print(sas.lastlog())
    logger.info('Loaded %s substation notification records from SAS.', len(df))

    
    logger.info('Running SAS substation reference query.')
    c = sas.submitLST(sas_ckt_sub_query)
    # print(sas.lastlog())
    
    logger.info('Loading substation reference records from SAS.')
    df_sub = sas.sasdata2dataframe(
        table='SUB_NUM',
        libref='WORK',
        method='CSV'
    )
    # print(sas.lastlog())
    logger.info('Loaded %s substation reference records from SAS.', len(df_sub))
    sas.endsas()
    logger.info('SAS session closed.')
    
    date_str = datetime.now().strftime("%Y_%m_%d")
    

    
# =============================================================================
# Add in information that doesnt come from SAP
# =============================================================================

    #Testing
    # df.to_excel(r'C:\Users\brownsjm\Desktop\test_Tran_event.xlsx')
    # df = pd.read_excel(r'C:\Users\brownsjm\Desktop\test_Tran_event.xlsx')

    df2 = df.drop_duplicates()
    logger.info(
        'Retained %s notification records after removing %s duplicates.',
        len(df2), len(df) - len(df2),
    )
    
    if len(df2) > 0:
        records_before_history_filter = len(df2)
        df2 = df2[~df2['Notification'].isin(prev_run_file['Notification'])]
        logger.info(
            'Excluded %s notifications already in run history; %s new records remain.',
            records_before_history_filter - len(df2), len(df2),
        )
    
    if len(df2) > 0:
    
        df2['Notification'] = df2['Notification'].astype(str)
    
        df2['Subject'] = 'Sub_SAP_' + df2['Notification']
        df2['Entry_Method'] = 'SAP'
        
        df2['ES_KEY'] = df2['Floc'].str[:7]

# =============================================================================
# Get df_substation information on both files ready to merge and get a district matching
# =============================================================================
        df_sub_1 = df_sub
        df_sub_1['SUB_FLOC'] = ('ES-' + df_sub_1['SUB_NO'].astype('Int64').astype(str).str.zfill(4))
        
        df_sub_1 = df_sub[~df_sub['SWTCH_CNTR_DESC'].isna()]
        del df_sub_1['SUB_NO']
        df_sub_1 = df_sub_1.drop_duplicates()
        logger.info('Prepared %s unique substation reference records for enrichment.', len(df_sub_1))
        
# ============================================================================
#Fill in Switching Center where it can be filled in for notifications without Circuits as well as assignment to Engineers
# =============================================================================
    #Engineer Assignment
        sys_assign = system_assignments()
        logger.info('Loaded %s system assignment records.', len(sys_assign))
        # sys_assign = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/System_Assignments.xlsx')
        # sys_assign = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/System_Assignments.xlsx')
    #District to System mapping file
        sys_dist_mapping = pd.read_excel('data/SWCNTR_DISTRICT_mapping.xlsx')
        logger.info('Loaded %s switching-center district mappings.', len(sys_dist_mapping))
    
        sys_dist_mapping = sys_dist_mapping[['SWITCHING_CENTER', 'DISTRICT_NUM']]
        sys_dist_mapping['DISTRICT_NUM'] = sys_dist_mapping['DISTRICT_NUM'].astype(str)
    
#District Num to District Name, Engineer Assigned (Old mapping before going by system)
        dist_inbox = district_inboxes()
        logger.info('Loaded %s district inbox records.', len(dist_inbox))
        # dist_inbox = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/District Inboxes.xlsx')
        # dist_inbox = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/District Inboxes.xlsx')
    
        dist_inbox_sub = dist_inbox[['District Number', 'District Name']]
    
        dist_inbox_sub['District Number'] = dist_inbox_sub['District Number'].astype(str)
    
        df3 = pd.merge(df2, df_sub_1, left_on= ['ES_KEY'],
                         right_on= ['SUB_FLOC'], how='left').drop(['SUB_FLOC'], axis=1)
        logger.info('Enriched %s notification records with substation reference data.', len(df3))
        
        # df4 = pd.merge(df3, sys_dist_mapping, left_on= ['District'],
        #                  right_on= ['DISTRICT_NUM'], how='left').drop(['DISTRICT_NUM'], axis=1)
    
    #Take Notification swithching center if there but take district system mapping if nan
    
        df3['SWTCH_CNTR_DESC'] = df3['SWTCH_CNTR_DESC'].fillna('')
        df3['SUBSTATION'] = df3['SUBSTATION'].fillna('')
    
        # def switch_center(not_switch, dist_switch):
        #     if not_switch == '' or not_switch == 'Missing' or not_switch == pd.isna(not_switch):
        #         return dist_switch
        #     else:
        #         return not_switch

        # df3['Switching_Center_UDF'] = df3.apply(lambda x: switch_center(x['FLOC_SwitchingCenter'], x['SWITCHING_CENTER']), axis=1)
    
        # del df4['Switching Center']
    
    
        df4 = pd.merge(df3, sys_assign, left_on= ['SWTCH_CNTR_DESC'],
                         right_on= ['Switching Center'], how='left').drop(['Switching Center'], axis=1)
        
        df4['Engineer Assigned'] = df4['Engineer Assigned'].fillna('Unassigned')
        unassigned_count = int(df4['Engineer Assigned'].eq('Unassigned').sum())
        if unassigned_count:
            logger.warning('%s records have no assigned engineer.', unassigned_count)
    
        # df5['Structure'] = df5['Floc'].str.replace('OH-', '', regex=True)
        # df5['Structure'] = df5['Structure'].str.replace('UG-', '', regex=True)
        
#         def extract_substation(name):
#             if pd.isna(name) or name.strip() == "":
#                 return ""
#             else:
#         # Split by the dash and return the first part
#                 return name.split('-')[0].strip()

# # Create a new column 'Substation' using apply
#         df5['Substation'] = df5['NotfCircuitName'].apply(extract_substation)
#         # df5['Substation'] = ''
    
        def hftd(Tier):
            if Tier == 'Elevated':
                return 'T2'
            elif Tier == 'Extreme':
                return 'T3'
            else:
                return 'NON-HFRA'

        df4['HFRA_Ckt'] = df4.apply(lambda x: hftd(x['Tier_Level']), axis=1)
        
        df4['DISTRICT_NO'] = df4['DISTRICT_NO'].astype("Int64").astype(str).fillna('')
        df4['DISTRICT_NAME'] = df4['DISTRICT_NAME'].fillna('')
        
        df4 = df4.rename(columns = {"DISTRICT_NO" : "District",
                                    "DISTRICT_NAME" : "District Name"})
    
        # df5 = pd.merge(df4, dist_inbox_sub, left_on= ['DISTRICT_NO'],
        #                  right_on= ['District Number'], how='left').drop(['District Number'], axis=1)
    
        # df6['Voltage (kV)'] = df6['Voltage (kV)'].fillna(0).astype(int).astype(str).replace('0', '')
        df5 = df4
# =============================================================================
#Get date from Long text and use for Outage Datetime and Format date time for PowerApps
# =============================================================================
    
        df5['OutageDT'] = df5['LONGTEXT1'].apply(lambda x: x.split(" PST")[0])
        df5['OutageDT'] = df5['OutageDT'].apply(lambda x: x.split("* ")[1])
    
        df5['OutageDT_adjust'] = pd.to_datetime(df5['OutageDT']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df5['OutageDT'] = pd.to_datetime(df5['OutageDT_adjust']).dt.strftime('%m/%d/%Y %H:%M')
        df5['OutageDT'] = df5['OutageDT'].str.strip()

        df5['Notf_Createdon_Date_adjust'] = pd.to_datetime(df5['Notf_Createdon_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df5['Notf_Createdon_Date'] = pd.to_datetime(df5['Notf_Createdon_Date_adjust']).dt.strftime('%m/%d/%Y')
        df5['Notf_Createdon_Date'] = df5['Notf_Createdon_Date'].str.strip()
    
        df5['Notf_Date_adjust'] = pd.to_datetime(df5['Notf_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
        df5['Notf_Date'] = pd.to_datetime(df5['Notf_Date_adjust']).dt.strftime('%m/%d/%Y')
        df5['Notf_Date'] = df5['Notf_Date'].str.strip()

        df5 = df5.drop_duplicates()
        logger.info('Prepared %s unique records for run history and upload.', len(df5))
        
        df5['LONGTEXT2'] = 'Notif: ' + df5['Notification'] + df5['LONGTEXT1']
        df5['Voltage (kV)'] = ''

# =============================================================================
#     Create work order info to match notification info
# =============================================================================
    # df_wo_1 = df_wo
    
    # df_wo_1['WorkOrderID'] = df_wo_1['WorkOrderID'].astype(str)
    # df_wo_1['Subject'] = 'Trans_SAP_WO_' + df_wo_1['WorkOrderID']
    # df_wo_1['EntryMethod'] = 'SAP'
    
    # df_wo_1['WO_FlocID'] = df_wo_1['WO_FlocID'].str.slice(0,8)
#WOFlocID may not be the one to link everything, shows up as structure for some and ET Circuit ID for others

        output_filepath = 'brownstein/substation/previous_runs_sap/Substation_Events_' + date_str + ".xlsx"
        logger.info('Saving run-history workbook: %s', output_filepath)
        df5.to_excel(output_filepath, index=False)
        logger.info('Saved %s records to %s.', len(df5), output_filepath)

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
            'SWTCH_CNTR_DESC': 'Switching_x0020_Center',
            'Voltage (kV)': 'Voltage_x0020__x0028_kV_x0029_',
            'SUBSTATION': 'Substation',
            'Engineer Assigned': 'Engineer_x0020_Assigned',
            'District': 'District_x0020_Number',
            'District Name': 'District_x0020_Name',
        }
        sharepoint_upload_columns = list(sharepoint_field_mapping)
        sharepoint_upload_data = df5[sharepoint_upload_columns]
        logger.info('Authenticating to SharePoint for the substation upload.')
        authenticate_user()
        logger.info(
            'Uploading %s records to SharePoint list %s.',
            len(sharepoint_upload_data), sharepoint_list_name,
        )
        uploaded_count = upload_dataframe_to_sharepoint_list(
            sharepoint_site=sharepoint_site,
            sharepoint_list_name=sharepoint_list_name,
            dataframe=sharepoint_upload_data,
            sharepoint_field_mapping=sharepoint_field_mapping,
        )

        logger.info('Uploaded %s records to SharePoint list %s.', uploaded_count, sharepoint_list_name)
    else:
        logger.info('No new substation notifications to process.')

except Exception as e:
    logger.exception('Substation SAP pipeline failed: %s', e)
finally:
    logger.info('SAS Connection terminated.')
