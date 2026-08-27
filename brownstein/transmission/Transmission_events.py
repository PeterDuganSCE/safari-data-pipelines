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

# =============================================================================
# Get work order informationt where there were no notifications
# =============================================================================
# sas_trans_wo_query = """
# PROC SQL;
# CREATE TABLE WORK.QUERY_FOR_INT_WOMASTER AS 
# SELECT DISTINCT t1.WorkOrderID, 
# t1.WorkOrderID_NZ, 
# t1.WorkOrder,
# t1.WO_SystemStatusAll, 
# t1.WO_UserStatusAll,
# t1.WO_PlantSection, 
# t1.WO_Workcenter, 
# t1.WO_MatCode, 
# t1.WO_CreatedOnDate, 
# t1.WO_FlocID, 
# t1.WO_FunctionalArea, 
# t1.WO_OrderType, 
# t1.WO_SystemStatus, 
# t1.Region, 
# t1.Grid_District, 
# t1.ABC_Indicator
# FROM TDINTGRN.INT_WOMASTER t1
# WHERE t1.WO_CreatedOnDate >= '""" + full_day_str + """'d AND t1.WO_CreatedOnDate < '""" + today_str + """'d AND t1.WO_MatCodeID IN 
# (
# '11B',
# '12B',
# '13B',
# '14B',
# '15B',
# '16B',
# '1B2',
# '1B3',
# '1B4',
# '1B5',
# '1B6',
# '1B7',
# '1B8',
# '1B9',
# '1BA',
# '1BB',
# '1BC',
# '1BD',
# '1BE',
# '105',
# '106'
# ) AND t1.WO_SystemStatusAll NOT LIKE 'REJC' AND t1.WO_UserStatusAll NOT CONTAINS 'CANC' AND t1.WO_TypeID = 'ETMA' AND t1.WO_NotificationID = ' ';
# QUIT;"""

#Test for larger date range
# sas_trans_wo_query = """
# PROC SQL;
# CREATE TABLE WORK.QUERY_FOR_INT_WOMASTER AS 
# SELECT DISTINCT t1.WorkOrderID, 
# t1.WorkOrderID_NZ, 
# t1.WorkOrder,
# t1.WO_SystemStatusAll, 
# t1.WO_UserStatusAll,
# t1.WO_PlantSection, 
# t1.WO_Workcenter, 
# t1.WO_MatCode, 
# t1.WO_CreatedOnDate, 
# t1.WO_FlocID, 
# t1.WO_FunctionalArea, 
# t1.WO_OrderType, 
# t1.WO_SystemStatus, 
# t1.Region, 
# t1.Grid_District, 
# t1.ABC_Indicator
# FROM TDINTGRN.INT_WOMASTER t1
# WHERE t1.WO_CreatedOnDate >= '18MAR2022'd AND t1.WO_CreatedOnDate < '""" + today_str + """'d AND t1.WO_MatCodeID IN 
# (
# '11B',
# '12B',
# '13B',
# '14B',
# '15B',
# '16B',
# '1B2',
# '1B3',
# '1B4',
# '1B5',
# '1B6',
# '1B7',
# '1B8',
# '1B9',
# '1BA',
# '1BB',
# '1BC',
# '1BD',
# '1BE',
# '105',
# '106'
# ) AND t1.WO_SystemStatusAll NOT LIKE 'REJC' AND t1.WO_UserStatusAll NOT CONTAINS 'CANC' AND t1.WO_TypeID = 'ETMA' AND t1.WO_NotificationID = ' ';
# QUIT;"""
# =============================================================================
# Run code
# =============================================================================
def find_valid_SASsession(username, password):
    server_nums = ['03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    server_num = ['11']
    for sn in server_num[:]:
        try:
            sas = saspy.SASsession(
                cfgfile='config/sascfg_personal.py',
                cfgname='iomcom',
            )
            # sas = saspy.SASsession(iomhost=f'lyxcpsas{sn}.sce.com',
            #                        iomport=8591,
            #                        class_id='440196d4-90f0-11d0-9f41-00a024bb830c',
            #                        provider='sas.iomprovider',
            #                        encoding='utf-8',
            #                        omruser=username,
            #                        omrpw=password,
            #                        m5dsbug=True)
            print(f'SERVER {sn} SUCCESS')
            return sas
        except Exception as e:
            print(f'SERVER {sn} FAILED: {e}')


# credentials = pd.read_excel('C:/Users/brownsjm/Documents/Credentials/Credentials_HANA.xlsx')
with open(AUTH_PATH, "r", encoding="utf-8") as auth_file:
    sas_credentials = (yaml.safe_load(auth_file) or {})["sas"]


try:
    # sas = find_valid_SASsession(sas_credentials["username"], sas_credentials["password"])
    # sas = saspy.SASsession(cfgname='iomcom')

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

#Run code for work order information
    # c = sas.submitLST(sas_trans_wo_query)
    # print(sas.lastlog())
    
    # df_wo = sas.sasdata2dataframe(
    #     table='QUERY_FOR_INT_WOMASTER',
    #     libref='WORK',
    #     method='CSV'
    # )
    
# =============================================================================
# Add in information that doesnt come from SAP
# =============================================================================

    #Testing
    # df.to_excel(r'C:\Users\brownsjm\Desktop\test_Tran_event.xlsx')
    # df = pd.read_excel(r'C:\Users\brownsjm\Desktop\test_Tran_event.xlsx')

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
        sys_assign = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/System_Assignments.xlsx')
        # sys_assign = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/System_Assignments.xlsx')
    #District to System mapping file
        sys_dist_mapping = pd.read_excel('data/SWCNTR_DISTRICT_mapping.xlsx')
    
        sys_dist_mapping = sys_dist_mapping[['SWITCHING_CENTER', 'DISTRICT_NUM']]
        sys_dist_mapping['DISTRICT_NUM'] = sys_dist_mapping['DISTRICT_NUM'].astype(str)
    
#District Num to District Name, Engineer Assigned (Old mapping before going by system)
        dist_inbox = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/District Inboxes.xlsx')
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
#     Create work order info to match notification info
# =============================================================================
    # df_wo_1 = df_wo
    
    # df_wo_1['WorkOrderID'] = df_wo_1['WorkOrderID'].astype(str)
    # df_wo_1['Subject'] = 'Trans_SAP_WO_' + df_wo_1['WorkOrderID']
    # df_wo_1['EntryMethod'] = 'SAP'
    
    # df_wo_1['WO_FlocID'] = df_wo_1['WO_FlocID'].str.slice(0,8)
#WOFlocID may not be the one to link everything, shows up as structure for some and ET Circuit ID for others
    # df_wo_2 = pd.merge(df_wo_1, trans_ckt_info_sub, left_on= ['WO_FlocID'],
    #                      right_on= ['FLOC'], how='left').drop(['FLOC'], axis=1)
    
    
# =============================================================================
# Write to a daily run that will also track duplicates
# =============================================================================
        # df.to_csv('sap_transmission_e1_notifications.csv')
        output_filepath = 'brownstein/transmission/previous_runs_sap/Transmission_Events_' + date_str + ".xlsx"
        df6.to_excel(output_filepath, index=False)

        logger.info(f"Created file -> {output_filepath}")

# =============================================================================
# Write to excel file for PowerAutomate to pick up the data to bring to RO Tool
# =============================================================================
        # output_file = 'C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/Transmission_Data.xlsx'
        output_file = 'C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/Transmission_Data_pd.xlsx'
        output_wb = load_workbook(output_file)
        output_ws = output_wb['Sheet1']
        output_wb.active = output_ws
#Historical information
        # output_file_hist = 'C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/Transmission_Data_hist.xlsx'
        output_file_hist = 'C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/Transmission_Data_hist_pd.xlsx'
        output_wb_hist = load_workbook(output_file_hist)
        output_ws_hist = output_wb_hist['Sheet1']
        output_wb_hist.active = output_ws_hist

        if len(df6) > 0:
            for index, row in df6.iterrows():
                write_data = [row['Subject'], row['Entry_Method'], row['NotfCircuitName'], row['Structure'], row['Voltage (kV)'], row['District'],
                          row['District Name'], row['Engineer Assigned'], row['Substation'], row['Switching_Center_UDF'], row['HFRA_Ckt'],
                          row['Notification'], row['NotfType'], row['Priority'], row['Notf_Createdon_Date'], row['OutageDT'], row['Problem_Stat'],
                          row['Notf_MAT'], row['MAT_Desc'], row['NotfPatrolTypeID'], row['NotfPatrolType'], row['Level_4'], row['Level_3'], row['Level_2'],
                      row['Level_1'], row['NotfCreatedBy'], row['Sys_Status_All'], row['Sys_Status'], row['HFRA_Ckt'], row['LONGTEXT2']]
                output_ws_hist.append(write_data)
    
            output_ws_hist.tables['Data'].ref = "A1:" + output_ws_hist.cell(output_ws_hist.max_row, output_ws_hist.max_column).column_letter + str(output_ws_hist.max_row)
    
            output_wb_hist.save(output_file_hist)
        
        
        if len(df6) > 0:
            for index, row in df6.iterrows():
                write_data = [row['Subject'], row['Entry_Method'], row['NotfCircuitName'], row['Structure'], row['Voltage (kV)'], row['District'],
                          row['District Name'], row['Engineer Assigned'], row['Substation'], row['Switching_Center_UDF'], row['HFRA_Ckt'],
                          row['Notification'], row['NotfType'], row['Priority'], row['Notf_Createdon_Date'], row['OutageDT'], row['Problem_Stat'],
                          row['Notf_MAT'], row['MAT_Desc'], row['NotfPatrolTypeID'], row['NotfPatrolType'], row['Level_4'], row['Level_3'], row['Level_2'],
                      row['Level_1'], row['NotfCreatedBy'], row['Sys_Status_All'], row['Sys_Status'], row['HFRA_Ckt'], row['LONGTEXT2']]
                output_ws.append(write_data)
    
            output_ws.tables['Data'].ref = "A1:" + output_ws.cell(output_ws.max_row, output_ws.max_column).column_letter + str(output_ws.max_row)

            output_wb.save(output_file)

    
        if len(df6) > 0:
            logger.info("Wrote data to excel")
        else:
            logger.info("No data to load")

        if os.path.exists(output_filepath):
            logger.info(f"Successfully created file -> {output_filepath}")
        else:
            logger.error(f"Did not create file -> {output_filepath}")
    else:
        print("No new SAP notifications")

except Exception as e:
    logger.error(e)
finally:
    logger.info('SAS Connection terminated.')
