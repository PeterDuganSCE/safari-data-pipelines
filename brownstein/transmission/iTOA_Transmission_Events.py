# -*- coding: utf-8 -*-
"""
Created on Fri Jul 26 06:26:55 2024

@author: brownsjm
"""
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

warnings.filterwarnings('ignore')
pd.set_option('display.width', 500)
pd.set_option('display.max_column', 500)
pd.set_option('display.max_rows', 200)
pd.set_option('display.max_colwidth', 500)

#pri_workcenter field in iTOA tables (GENERIC)

#%%Load SQL data for iTOA

# SQL Server connection info
DB_DRIVER = "{SQL Server}" # Or whatever driver is installed in your machine
DB_SERVER_2 = 'tcp:AYWCPSQL116\SQLP740,1989'  # Irvine Server
# DB_SERVER_2 = 'tcp:AYWCPSQL651\SQLP648,1989'  # Alhambra Server
DATABASE_NAME_2 = 'ITOA'  # development server, for testing
connection_itoa = pyodbc.connect(f"DRIVER={DB_DRIVER};SERVER={DB_SERVER_2};DATABASE={DATABASE_NAME_2};Trusted_Connection=no;")


itoa_gen_q = '''SELECT DISTINCT APP_ID, LINE_OF_BUSINESS, APPLICATION_STATUS, EQUIPMENT, VOLTAGE_LEVEL, MAIN_ACTUAL_OUT, DISTURBANCE_DURATION, DISTURBANCE_NO, AFFECTED_DISTRICT, WORK_CENTERS, WEATHER, MISCELLANEOUS_COMMENTS, event_activity_comments, INITIATING_CAUSE_CODE, INITIATING_SUB_CAUSE, CAUSE_CATEGORY, OUTAGE_CATEGORY, PROCESS_STATUS FROM dbo.GENERIC_INT_TADS_V'''
itoa_gen = pd.read_sql(itoa_gen_q, connection_itoa)

itoa_gen = itoa_gen.drop_duplicates()

itoa_gen_t = itoa_gen[itoa_gen['LINE_OF_BUSINESS'] == 'Transmission']
itoa_gen_t = itoa_gen_t[itoa_gen_t['PROCESS_STATUS'] != 'VOID']

del itoa_gen_t['PROCESS_STATUS']

itoa_gen_t['iTOA_link'] = itoa_gen_t['APP_ID'].apply(lambda x: f'https://orls.sce.com/itoa/autooutage/view.htmlx?editedOutage.appId={x}&eventAnalysis.analysisId=&referer=interruption')

# =============================================================================
# Need to run info for 2 days prior since notifications loaded at 6pm, so have to make sure all created on dates give the chance for info to get into system
# =============================================================================
# today = datetime.now().strftime("%m/%d/%Y")
today_str = datetime.today().strftime(" %d%b%Y")
today = datetime.strptime(today_str, " %d%b%Y")

#Get date of last run to pull data from days that it didnt run
list_of_files = glob.glob('brownstein/transmission/previous_runs_itoa/*')
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
itoa_gen_t['Out_Date'] = pd.to_datetime(itoa_gen_t['MAIN_ACTUAL_OUT'], errors='coerce')


# today = dt.datetime.today()
# yesterday = today - dt.timedelta(days=1)
# today = today.strftime('%m/%d/%Y')
# formatted_today = dt.datetime.strptime(today, '%m/%d/%Y')

# yesterday = yesterday.strftime('%m/%d/%Y')
itoa_gen_t_2 = itoa_gen_t[itoa_gen_t['Out_Date'] >= full_day]
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
    sys_assign = pd.read_excel('C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/System_Assignments.xlsx')
    # sys_assign = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/System_Assignments.xlsx')

    itoa_gen_t3['WORK_CENTERS'] = itoa_gen_t3['WORK_CENTERS'].str.upper()

    itoa_gen_t4 = pd.merge(itoa_gen_t3, sys_assign, left_on= ['WORK_CENTERS'],
                         right_on= ['Switching Center'], how='left').drop(['Switching Center'], axis=1)

    itoa_gen_t4['Engineer Assigned'] = itoa_gen_t4['Engineer Assigned'].fillna('Unassigned')

    
#Add columns for Engineers to track comments and review times (Old before automation)
# itoa_gen_t['Engineer_Assigned'] = ''
# itoa_gen_t['SAP_Notif'] = ''
# itoa_gen_t['SAP_Notif_Num'] = ''
# itoa_gen_t['Desktop_Review_Time'] = ''
# itoa_gen_t['Days_for_Sr_Patrol_Info'] = ''
# itoa_gen_t['Engineer_Comments'] = ''
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

# =============================================================================
# Change timestamp to account for UTC in PowerAutomate
# =============================================================================

    itoa_gen_t4['Notf_Date_adjust'] = pd.to_datetime(itoa_gen_t4['Notf_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t4['Notf_Date'] = pd.to_datetime(itoa_gen_t4['Notf_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t4['Notf_Date'] = itoa_gen_t4['Notf_Date'].str.strip()

    itoa_gen_t4['Email_Date_adjust'] = pd.to_datetime(itoa_gen_t4['Email_Date']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t4['Email_Date'] = pd.to_datetime(itoa_gen_t4['Email_Date_adjust']).dt.strftime('%m/%d/%Y %H:%M')
    itoa_gen_t4['Email_Date'] = itoa_gen_t4['Email_Date'].str.strip()

# =============================================================================
# Write to previous run to not get duplicates and base pull off of last run
# =============================================================================
    output_filepath = 'brownstein/transmission/previous_runs_itoa/iTOA_Transmission_Events_' + date_str + ".xlsx"
    itoa_gen_t4.to_excel(output_filepath, index=False)

# =============================================================================
# Write to excel file for PowerAutomate to pick up the data to bring to RO Tool
# =============================================================================
#Test on 1 event
# itoa_gen_t4 = itoa_gen_t4[0:1]

    # output_file = 'C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/Transmission_Data_iTOA.xlsx'
    output_file = 'C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/Transmission_Data_iTOA.xlsx'
    output_wb = load_workbook(output_file)
    output_ws = output_wb['Sheet1']
    output_wb.active = output_ws
#Historical information
    # output_file_hist = 'C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/Transmission_Data_hist_iTOA.xlsx'
    output_file_hist = 'C:/Users/duganpr/Southern California Edison/Failure Event Analysis - Keyword Scrape/Transmission_Data_hist_iTOA.xlsx'
    output_wb_hist = load_workbook(output_file_hist)
    output_ws_hist = output_wb_hist['Sheet1']
    output_wb_hist.active = output_ws_hist

    if len(itoa_gen_t4) > 0:
        for index, row in itoa_gen_t4.iterrows():
            write_data = [row['Subject'], row['Entry_Method'], row['NotfCircuitName'], row['Voltage'], row['Engineer_Assigned'], row['Substation'],
                      row['Switching_Center'], row['HFRA_CIRCUIT?'], row['Email_Date'], row['Notf_Date'], row['NotfCreatedBy'],
                      row['WEATHER'], row['DISTURBANCE_DURATION'], row['MISCELLANEOUS_COMMENTS'], row['event_activity_comments'],
                      row['INITIATING_CAUSE_CODE'], row['INITIATING_SUB_CAUSE'], row['CAUSE_CATEGORY'], row['iTOA_link'], row['LONGTEXT1']]
            output_ws_hist.append(write_data)
    
        output_ws_hist.tables['Data'].ref = "A1:" + output_ws_hist.cell(output_ws_hist.max_row, output_ws_hist.max_column).column_letter + str(output_ws_hist.max_row)
# #output_wb.save(output_file)
    
        output_wb_hist.save(output_file_hist)


    if len(itoa_gen_t4) > 0:
        for index, row in itoa_gen_t4.iterrows():
            write_data = [row['Subject'], row['Entry_Method'], row['NotfCircuitName'], row['Voltage'], row['Engineer_Assigned'], row['Substation'],
                      row['Switching_Center'], row['HFRA_CIRCUIT?'], row['Email_Date'], row['Notf_Date'], row['NotfCreatedBy'],
                      row['WEATHER'], row['DISTURBANCE_DURATION'], row['MISCELLANEOUS_COMMENTS'], row['event_activity_comments'],
                      row['INITIATING_CAUSE_CODE'], row['INITIATING_SUB_CAUSE'], row['CAUSE_CATEGORY'], row['iTOA_link'], row['LONGTEXT1']]
        
            output_ws.append(write_data)
        output_ws.tables['Data'].ref = "A1:" + output_ws.cell(output_ws.max_row, output_ws.max_column).column_letter + str(output_ws.max_row)
        output_wb.save(output_file)
        
else:
    print("No new iTOA Events to process")

#read in existing master list
# master = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\FIPA\Transmission_Events_test\01_Transmission_Events_Master.xlsx')

# if len(itoa_gen_t) > 0:
#     master_2 = master.append(itoa_gen_t, ignore_index = True, sort = False)

#     master_2.to_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\FIPA\Transmission_Events_test\01_Transmission_Events_Master.xlsx', index = False)

# #Older write per day, currently replacing with master list
#     write_location = '//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/Transmission_Events_test/iTOA_Transmission_Events_' + month_number + '_' + day_number + '_' + year_number +'.xlsx'

#     itoa_gen_t.to_excel(write_location, index = False)
#===============================================
