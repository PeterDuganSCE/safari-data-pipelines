# -*- coding: utf-8 -*-
"""
Created on Fri Feb 10 06:39:53 2023

@author: brownsjm
"""


from hdbcli import dbapi
# import numpy as np
import pandas as pd
import yaml
# from datetime import date, timedelta
import datetime as dt
# import datetime
# import dateutil
# import math
# from pyproj import Proj
from pandas import DataFrame
# import utm
import os
from pathlib import Path
import saspy
import glob
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.paths import AUTH_PATH


# PRINT TO SHOW SCRIPT HAS BEEN LOADED
print('RUNNING OEIS_OMS.py')

#from dplython import (DplyFrame, X, diamonds, select, sift, sample_n, sample_frac, head, arrange, mutate, group_by,
#                      summarize, DelayFunction)

#Get created date of newest file in report folder
list_of_files = glob.glob('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/*.csv')  # change this to the directory you want to search
latest_file = max(list_of_files, key=os.path.getctime)
# latest_file = 'SCE_2026-01-06-1-29300.csv'
# latest_file = r'C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report\SCE_2026-01-06-1-29300.csv'
t = latest_file.split("SCE_")[1]
t = t.split(".")[0]
t = "-".join(t.split("-")[:3])
# t = t.replace("_", "/")
last_date = pd.to_datetime(t)

#get system date to define yesterday's date and pull all outages from yesterday
today = dt.datetime.today()

delta = today - last_date
delta_days = delta.days

before_weekend = today - dt.timedelta(days=delta_days)
before_weekend_year = before_weekend - dt.timedelta(days=365)
before_weekend_year_year = before_weekend_year.year
today_year=today.year

before_weekend_year_year = str(before_weekend_year_year)

#%% setup credentials for accessing HANA

with open(AUTH_PATH, "r", encoding="utf-8") as auth_file:
    hana_credentials = (yaml.safe_load(auth_file) or {})["hana"]

city_county = pd.read_excel('City_County.xlsx')
pole_census = pd.read_csv('Pole_Census_SurfaceFuels_20230308.csv')

pole_census_1 = pole_census[['SCE_STRUCTURE_NO', 'Definition', 'URBANRURAL']]

#%% Query table in HANA for OMS data
#Connecting to HANA

conn = dbapi.connect(
    address="vp55db51.sce.com",
    port=30015,
    user=hana_credentials["username"],
    password=hana_credentials["password"],
    databasename='P55'
    )

# conn = dbapi.connect(
#     address="vp55db.sce.com",
#     port=30015,
#     user=hana_credentials["username"],
#     password=hana_credentials["password"],
#     databasename='P55'
#     )

#%% Query open and closed calls to get 911 response
#Obtain circuit_id to circt_nam mapping
hana_ckt = '''select "CIRCT_ID" AS "CKT_ID", "CIRCT_NAM", "CIRCT_DESC" from "OMS"."CIRCUIT";'''
ckt_id = pd.read_sql_query(hana_ckt, conn)
# ckt_id = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/HANA_CKT.xlsx')

print("Finished pulling data from HANA, exiting program")
sys.exit()
#table from non hist incidents location (did not receive enough zip codes)
#hana_loc = '''select "INCIDENT_ID" AS "Incident_ID", "ZIP_CD" from "OMS"."LOCATION"
#WHERE YEAR(CREATION_DATETIME) >= ''' + yesterday_year_year + ''';'''
#loc = pd.read_sql_query(hana_loc, conn)
#
#loc_1 = loc[loc['ZIP_CD'] != '']
#loc_1 = loc[loc['ZIP_CD'] != 'None']
#loc_1 = loc[~loc['ZIP_CD'].isna()]
#
#loc_2 = loc_1.groupby(['Incident_ID'], as_index=False).first()

#Obtain call comments from HANA OMS
hana_call = '''select CUST_NAM, CITY_NAM, CONTACT_NAM, CLUE_CD, CALL_TYPE_CD, AFF_DATETIME, STRCTUR_NO,
CIRCT_ID, INCIDENT_ID from "OMS"."CALL"
WHERE YEAR(AFF_DATETIME) >= ''' + before_weekend_year_year + '''
AND "CLUE_CD" LIKE '%9';'''
call = pd.read_sql_query(hana_call, conn)
# call = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/HANA_CALL.xlsx')


call = call[call['AFF_DATETIME'] >= before_weekend]

#For adding later since separates in SAS join
hana_call_text = '''select INCIDENT_ID, COMMENT_TEXT, AFF_DATETIME from "OMS"."CALL"
WHERE YEAR(AFF_DATETIME) >= ''' + before_weekend_year_year + '''
AND "CLUE_CD" LIKE '%9';'''
call_text = pd.read_sql_query(hana_call_text, conn)
# call_text = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/HANA_CALL_TEXT.xlsx')

#call_text['CLUE_DESC1'] = "Unknown"
#call_text['TOT_LOSS_POWER_FLG'] = "Unknown"

call_text = call_text[call_text['AFF_DATETIME'] >= before_weekend]
del call_text['AFF_DATETIME']


hana_his_call = '''select CUST_NAM, CITY_NAM, CONTACT_NAM, CLUE_CD, CALL_TYPE_CD, AFF_DATETIME, STRCTUR_NO,
CIRCT_ID, INCIDENT_ID from "OMS"."HIS_CALL"
WHERE YEAR(AFF_DATETIME) >= ''' + before_weekend_year_year + '''
AND "CLUE_CD" LIKE '%9';'''
his_call = pd.read_sql_query(hana_his_call, conn)
# his_call = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/HANA_HIST_CALL.xlsx')

his_call = his_call[his_call['AFF_DATETIME'] >= before_weekend]


#For adding later since separates in SAS join
hana_his_call_text = '''select INCIDENT_ID, AFF_DATETIME, CLUE_DESC1, TOT_LOSS_POWER_FLG, COMMENT_TEXT from "OMS"."HIS_CALL"
WHERE YEAR(AFF_DATETIME) >= ''' + before_weekend_year_year + '''
AND "CLUE_CD" LIKE '%9';'''
his_call_text = pd.read_sql_query(hana_his_call_text, conn)
# his_call_text = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/HANA_HIS_CALL_TEXT.xlsx')

his_call_text = his_call_text[his_call_text['AFF_DATETIME'] >= before_weekend]
del his_call_text['AFF_DATETIME']


call_1 = pd.concat([call,his_call], axis=0, sort= True).drop_duplicates()
call_text_1 = pd.concat([call_text,his_call_text], axis=0, sort= True).drop_duplicates()

call_1['CITY_NAM'] = call_1['CITY_NAM'].str.upper()
call_1['CITY_NAM'] = call_1['CITY_NAM'].str.replace('COUNTY OF ', '')

#Add county from flat table for city to county information (wikipedia)
call_1 = pd.merge(call_1, city_county, left_on= ['CITY_NAM'],
                         right_on= ['City'], how='left').drop(['City'], axis=1)

#Get circuit name from circuit table in OMS
call_1 = pd.merge(call_1, ckt_id, left_on= ['CIRCT_ID'],
                         right_on= ['CKT_ID'], how='left').drop(['CKT_ID'], axis=1)




#Get structure from Equip_Structure column in OMS table (Seems to have a STRCTUR_NO)
#call_1['Structure'] = call_1['EQUIP_STN_NO'].str.split(':').str[1]
#
#call_1['Structure'] = call_1['Structure'].str.replace("V", "")
#call_1['Structure'] = call_1['Structure'].str.replace("S", "")
#call_1['Structure'] = call_1['Structure'].str.replace("P", "")
#call_1['Structure'] = call_1['Structure'].str.replace("M", "")
#call_1['Structure'] = call_1['Structure'].str.replace("T1", "")
#call_1['Structure'] = call_1['Structure'].str.replace("T2", "")
#
#call_1['Structure'] = call_1['Structure'].str.split('.').str[0]
#
#
#def structure(Structure):
#    if Structure.startswith('X'):
#        return Structure.str.replace("X", "")
#    else:
#            return Structure
#        
#
#call_1['Structure1'] = call_1.apply(lambda x: structure(x['Structure']), axis=1)
#
#del call_1['Structure']




#%%Use SAS to link structure information (lat, long, etc)

#Log into SAS
def find_valid_SASsession(username, password):
    server_nums = ['03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for sn in server_nums[:]:
        try:
            sas = saspy.SASsession(iomhost=f'lyxcpsas{sn}.sce.com',
                                   iomport=8591,
                                   class_id='440196d4-90f0-11d0-9f41-00a024bb830c',
                                   provider='sas.iomprovider',
                                   encoding='utf-8',
                                   omruser=username,
                                   omrpw=password,
                                   m5dsbug=True)
            print(f'SERVER{sn} SUCCESS')
            return sas
        except Exception as e:
            print(f'SERVER {sn} FAILED: {e}')


credentials = pd.read_excel('C:/Users/brownsjm/Documents/Credentials/Credentials_HANA.xlsx')

# sas = saspy.SASsession(iomhost=f'lyxcpsas07.sce.com',
#                                    iomport=8591,
#                                    class_id='440196d4-90f0-11d0-9f41-00a024bb830c',
#                                    provider='sas.iomprovider',
#                                    encoding='utf-8',
#                                    omruser=credentials.iat[0,0],
#                                    omrpw=credentials.iat[0,1],
#                                    m5dsbug=True)

sas = find_valid_SASsession(credentials.iat[0,0], credentials.iat[0,1])

#If SAS cannot log in, export to excel and run SAS EG code

# call_1.to_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\Jbrow\OEIS_OMS\call_1.xlsx')

# sas = saspy.SASsession(cfgname='iomcom', m5dsbug = [True])
#checks log
#print(sas.lastlog())
#Write call_1 to SAS to use in query; Saves a lot of time of reading out larger list from SAS
CALL_2 = sas.df2sd(call_1, 'CALL_2')
# call_1.to_csv('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/OEIS_OMS/midput/call_1.csv', index = False)

c = sas.submitLST("""PROC SQL;
                      CREATE TABLE WORK.STRUCT AS
                      SELECT DISTINCT FLOC_StructureID, FLOC_Latitude, FLOC_Longitude, FLOC_Tier_Level, FLOC_District, FLOC_DistrictID
                      FROM TDSTAGNG.STG_FLOCMASTER
                      WHERE FLOC_InacInd NE 'X'
                      AND ((FLOC_PlannerGroup LIKE '2%' OR FLOC_PlannerGroup LIKE'1%')
                      AND (FLOC_FunctionalLocationID LIKE 'OH%' OR FLOC_FunctionalLocationID LIKE 'UG%')
                      AND (FLOC_TechObjectType LIKE 'ED%' OR FLOC_TechObjectType LIKE 'ET%' OR FLOC_TechObjectType LIKE 'EZ%'))
                      AND FLOC_TechObjectType NOT LIKE 'ED_HH'
                      AND FLOC_TechObjectType NOT LIKE 'ED_SL'
                      AND FLOC_TechObjectType NOT LIKE 'ED_TRENCH'
                      AND FLOC_TechObjectType NOT LIKE 'ED_TUNNEL';
                      QUIT;
                      """)

                      
c = sas.submitLST("""PROC SQL;
                      CREATE TABLE WORK.CALL_3 AS
                      SELECT DISTINCT t1.*, t2.*
                      FROM WORK.CALL_2 t1
                      LEFT JOIN WORK.STRUCT t2 ON t1.STRCTUR_NO = t2.FLOC_StructureID;
                      QUIT;
                      """)



c = sas.submitLST("""PROC SQL;
                      CREATE TABLE WORK.WEATHER AS
                      SELECT DISTINCT FLOC, STATION_NUMBER, SUBSTR(FLOC,INDEX(FLOC,"-")+1,20) AS FLOC_2
                      FROM TDSTAGNG.STG_SCE_WS_TO_FLOC_MAPPING;
                      QUIT;
                      """)
                      

                      
c = sas.submitLST("""PROC SQL;
                     CREATE TABLE WORK.CALL_4 AS
                     SELECT DISTINCT t1.*, t2.STATION_NUMBER
                     FROM WORK.CALL_3 t1
                     LEFT JOIN WORK.WEATHER t2 ON t1.STRCTUR_NO = t2.FLOC_2;
                     QUIT;
                      """)



# c = sas.submitLST("""PROC SQL;
#                    CREATE TABLE WORK.CKT_NAM AS
#                    SELECT DISTINCT _BIC_ZTDBUFLOC AS CKT_NUM, Circuit
#                    FROM TDSTAGNG.CIRCUIT_LIST;
#                    QUIT;
#                    """)
                   
c = sas.submitLST("""PROC SQL;
                   CREATE TABLE WORK.CKT_NAM AS
                   SELECT DISTINCT Circuit_Number AS CKT_NUM, Circuit_Name AS Circuit
                   FROM TDSTAGNG.CIRCUIT_NAMENUMBER;
                   QUIT;
                   """)

# print(sas.lastlog())

c = sas.submitLST("""PROC SQL;
                   CREATE TABLE WORK.SUB AS
                   SELECT DISTINCT CKT_NAME, SUBSTATION, SUB_NO, Region_Name AS Region_Name
                   FROM TDSTAGNG.CIRCUITSUBSTAT;
                   QUIT;
                   """)
                   
print(sas.lastlog())

# call_t = sas.sasdata2dataframe('CALL_3', libref='WORK', method='CSV')
call_4 = sas.sasdata2dataframe('CALL_4', libref='WORK', method='CSV') #sd2df() is a short hand version
ckt_nam = sas.sasdata2dataframe('CKT_NAM', libref='WORK', method='CSV')
sub = sas.sasdata2dataframe('SUB', libref='WORK', method='CSV')

del call_4['FLOC_StructureID']

sas.endsas()

#If SAS not working, go to SAS EG program and run to get excel

# call_4 = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\Jbrow\OEIS_OMS\CALL_4.xlsx')
# ckt_nam = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\Jbrow\OEIS_OMS\CKT_NAM.xlsx')
# sub = pd.read_excel(r'\\sce\workgroup\TDBU7\RESO-CRT\PERFORMANCE MANAGEMENT AND ANALYSIS\Jbrow\OEIS_OMS\SUB.xlsx')

#Redesign SAP Tier designation to Tier 2, Tier 3, and Non-HFTD
def hftd(Tier):
    if Tier == 'Elevated':
        return 'Tier 2'
    elif Tier == 'Extreme':
        return 'Tier 3'
    elif Tier == '':
        return 'Unknown'
    else:
            return 'Non-HFTD'

call_4['HFTD Class'] = call_4.apply(lambda x: hftd(x['FLOC_Tier_Level']), axis=1)


#%% Start matching original OEIS components
#Merge in clue desc, tot_loss_flg, and comment text
call_5 = pd.merge(call_4, call_text_1, left_on= ['INCIDENT_ID'],
                         right_on= ['INCIDENT_ID'], how='left')


#Combine CUST_NAM and CONTACT_NAM to get Fire department out of it

call_5['CUST_NAM'] = call_5['CUST_NAM'].fillna('NONE')
call_5['CONTACT_NAM'] = call_5['CONTACT_NAM'].fillna('NONE')
call_5['Suppressing Agency'] = call_5['CUST_NAM'] + '_' + call_5['CONTACT_NAM']
call_5['Suppressing Agency'] = call_5['Suppressing Agency'].str.upper()

call_5['FIRE'] = call_5['Suppressing Agency'].str.contains('FIRE')
call_5['FD'] = call_5['Suppressing Agency'].str.contains('FD')

def fire(FIRE, FD):
    if FIRE == True or FD == True:
        return 'YES'
    else:
            return 'NO'

call_5['FIRE_FD'] = call_5.apply(lambda x: fire(x['FIRE'],x['FD']), axis=1)

del call_5['FIRE']
del call_5['FD']

call_5 = call_5[call_5['FIRE_FD'] == 'YES']
#Add in selected columns from pole census (URBANRURAL, Definition)
call_6 = pd.merge(call_5, pole_census_1, left_on= ['STRCTUR_NO'],
                         right_on= ['SCE_STRUCTURE_NO'], how='left').drop(['SCE_STRUCTURE_NO'], axis=1)

call_6['URBANRURAL'] = call_6['URBANRURAL'].fillna('Undetermined at time of reporting')
call_6 = call_6.rename(columns = {"URBANRURAL" : "Origin Land Use"})

call_6['Definition'] = call_6['Definition'].fillna('Undetermined at time of reporting')
call_6 = call_6.rename(columns = {"Definition" : "Fuel Bed Description"})
call_6['Fuel Bed Description Comment'] = "N/A"

#Start matching original OEIS components
call_6['Notification Date']  = today
call_6['Notification Type'] = '29300(a)(1)'
call_6['Utility ID'] = 'SCE'
call_6['Confidential'] = 'No'
call_6 = call_6.rename(columns = {"AFF_DATETIME" : "Incident Start Date and Time"})
#Get circuit voltage
call_6['Voltage'] = call_6['CIRCT_DESC'].str[-4:]
call_6['Voltage'] = call_6['Voltage'].str.replace(" ", "")

call_6['Material at Origin'] = "Undetermined at time of reporting"
call_6['Material at Origin Comment'] = "N/A"

#Maybe can get (comes from National Weather service)
#Get a notice that RFW for yesterday and mark all events that day
call_6['RFW Status'] = 'Undetermined at time of reporting'
call_6['RFW Issue Date and Time'] = ""
call_6['FWW Status'] = 'Undetermined at time of reporting'
call_6['FWW Issue Date and Time'] = ""
call_6['HWW Status'] = 'Undetermined at time of reporting'
call_6['HWW Issue Date and Time'] = ""
call_6['Detection Method'] = "Agency"
call_6['Detection Method Comment'] = "N/A"
call_6['Fire Size'] = "Undetermined at time of reporting"
call_6['Suppressed By'] = "Fire agency"

call_6['Fire Investigation'] = "Will be done after event"
call_6['Fire AHJ'] = ""
call_6['Toutage ID'] = ""
call_6['Determination'] = "Utility Personnel"
call_6['Determination Comment'] = "N/A"
call_6['SuspectedInitiatingCauseComment'] = "N/A"
call_6['Equipment Failure'] = "Undetermined at time of reporting"
call_6['Equipment Failure Comment'] = "N/A"
call_6['Object Contact'] = "Undetermined at time of reporting"
call_6['Object Contact Comment'] = "N/A"
call_6['Facility Contacted'] = "Undetermined at time of reporting"
call_6['Facility Contacted Comment'] = "N/A"


#Add Outage Status as TOT_LOSS_POWER_FLG

def out_status(TOT_LOSS):
    if TOT_LOSS == 'F':
        return 'NO'
    elif TOT_LOSS == 'T':
        return 'YES'
    else:
        return 'Unknown'

call_6['Outage Status'] = call_6.apply(lambda x: out_status(x['TOT_LOSS_POWER_FLG']), axis=1)


#%%Wrangle information brought back

call_6 = call_6.rename(columns = {"FLOC_Latitude" : "Latitude",
                                  "FLOC_Longitude" : "Longitude"})
    
call_6['FLOC_District'] = call_6['FLOC_District'].str.upper()
call_6['FLOC_District'] = call_6['FLOC_District'].str.replace(" DISTRICT", "")
call_6['FLOC_DistrictID'] = call_6['FLOC_DistrictID'].str.replace("ED", "")


ckt_sub = pd.merge(sub, ckt_nam, left_on= ['CKT_NAME'],
                        right_on= ['Circuit'], how='left').drop(['Circuit'], axis=1)

# ckt_sub = sub

call_7 = pd.merge(call_6, ckt_sub, left_on= ['CIRCT_NAM'],
                      right_on= ['CKT_NAME'], how= 'left').drop(['CKT_NAME'], axis=1)

call_7['TOT_LOSS_POWER_FLG'] = call_7['TOT_LOSS_POWER_FLG'].fillna('Unknown')
#call_7['CLUE_DESC1'] = call_7['CLUE_DESC1'].fillna('Unknown')
call_7['CLUE_DESC1'] = "Undetermined at time of reporting"

call_7 = call_7.rename(columns = {"SUB_NO" : "Substation",
                                  "STATION_NUMBER" : "Nearest Weather Station ID",
                                  "WEATHER_STN_ID" : "Nearest Weather Station ID",
                                  "FLOC_District" : "District",
                                  "CLUE_DESC1" : "SuspectedInitiatingCause",
                                  "COMMENT_TEXT" : "Additional Notes",
                                  "INCIDENT_ID" : "Doutage ID",
                                  "CKT_NUM" : "Circuit ID",
                                  "Voltage" : "Circuit Voltage"})




#%% Re-order to match OEIS Report

call_8 = call_7[['Notification Date', 'Notification Type', 'Utility ID', 'Confidential', 'Incident Start Date and Time',
                 'County', 'District', 'Latitude', 'Longitude', 'HFTD Class', 'Origin Land Use', 'Material at Origin',
                 'Material at Origin Comment', 'Fuel Bed Description', 'Fuel Bed Description Comment', 'Circuit ID',
                 'Circuit Voltage', 'Substation', 'Nearest Weather Station ID', 'RFW Status', 'RFW Issue Date and Time',
                 'FWW Status', 'FWW Issue Date and Time', 'HWW Status', 'HWW Issue Date and Time', 'Detection Method',
                 'Detection Method Comment', 'Fire Size', 'Suppressed By', 'Suppressing Agency', 'Fire Investigation',
                 'Fire AHJ', 'Outage Status', 'Toutage ID', 'Doutage ID', 'SuspectedInitiatingCause', 'SuspectedInitiatingCauseComment',
                 'Determination', 'Determination Comment', 'Equipment Failure', 'Equipment Failure Comment', 'Object Contact',
                 'Object Contact Comment', 'Facility Contacted', 'Facility Contacted Comment', 'Additional Notes']].reset_index()

call_8 = call_8.rename(columns = {"index" : "ID",
                                "Notification Date" : "NotificationDate",
                                  "Notification Type" : "NotificationType",
                                  "Utility ID" : "UtilityID",
                                  "Incident Start Date and Time" : "IncidentStartDateTime",
                                  "HFTD Class" : "HFTDClass",
                                  "Origin Land Use" : "OriginLandUse",
                                  "Material at Origin" : "MaterialAtOrigin",
                                  "Material at Origin Comment": "MaterialAtOriginComment",
                                  "Fuel Bed Description" : "FuelBedDescription",
                                  "Fuel Bed Description Comment" : "FuelBedDescriptionComment",
                                  "Circuit ID" : "CircuitID",
                                  "Circuit Voltage" : "CircuitVoltage",
                                  "Substation" : "SubstationID",
                                  "Nearest Weather Station ID" : "NearestWeatherStationID",
                                  "RFW Status" : "RFWStatus",
                                  "RFW Issue Date and Time" : "RFWIssueDateTime",
                                  "FWW Status" : "FWWStatus",
                                  "FWW Issue Date and Time" : "FWWIssueDateTime",
                                  "HWW Status" : "HWWStatus",
                                  "HWW Issue Date and Time" : "HWWIssueDateTime",
                                  "Detection Method" : "DetectionMethod",
                                  "Detection Method Comment" : "DetectionMethodComment",
                                  "Fire Size" : "FireSize",
                                  "Suppressed By" : "SuppressedBy",
                                  "Suppressing Agency" : "SuppressingAgency",
                                  "Fire Investigation" : "FireInvestigation",
                                  "Fire AHJ" : "FireAHJ",
                                  "Outage Status" : "OutageStatus",
                                  "Toutage ID" : "TOutageID",
                                  "Doutage ID" : "DOutageID",
                                  "Determination Comment" : "DeterminationComment",
                                  "Equipment Failure" : "EquipmentFailure",
                                  "Equipment Failure Comment" : "EquipmentFailureComment",
                                  "Object Contact" : "ObjectContact",
                                  "Object Contact Comment" : "ObjectContactComment",
                                  "Facility Contacted" : "FacilityContacted",
                                  "Facility Contacted Comment" : "FacilityContactedComment",
                                  "Additional Notes" : "AdditionalNotes"})

call_8['Temp'] = call_8['IncidentStartDateTime'].dt.strftime('%m/%d/%Y')

call_8['AdditionalNotes'] = "An event occurred on " + call_8['Temp'] + " that impacted SCE facilities in our " + call_8['District'] + " District on circuit " + call_8['CircuitID'] + "."
del call_8['Temp']




as_of = pd.Timestamp.today().strftime('%Y-%m-%d')

# =============================================================================
# Read in master lists (add call8(new data) to master list and look for only new line items entered into wf and esir)
# =============================================================================
#Tracking file for what has already been submitted
master_list = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', sheet_name = '29300(a)(1)')
master_list_a2 = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', sheet_name = '29300(a)(2)')
master_list_b = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', sheet_name = '29300(b)')

#Read in to see if anything manually added
wf = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_WF/WF_29300_a2.xlsx', sheet_name = '29300(a)(2)')
#Read in to see if anything manually added
esir = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_ESIR/ESIR_29300_b.xlsx', sheet_name = '29300(b)')

#if manually added there would be line items in wf2 and esir2, otherwise will be a blank file
wf2 = wf[~wf['ID'].isin(master_list_a2['ID'])]
esir2 = esir[~esir['ID'].isin(master_list_b['ID'])]

master_list = master_list.loc[:, ~master_list.columns.str.contains('^Unnamed', case=False)]
master_list_a2 = master_list_a2.loc[:, ~master_list_a2.columns.str.contains('^Unnamed', case=False)]
master_list_b = master_list_b.loc[:, ~master_list_b.columns.str.contains('^Unnamed', case=False)]

# =============================================================================
# Write call 8 file to both folders, FIPA - FIPA Reporting will generate email from Ray
# =============================================================================
#Write call 8 data, new wf or new esir data to daily sheet
# This one will not email out, copy for ourselves
with pd.ExcelWriter('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/OEIS_' + as_of + '.xlsx', engine='openpyxl') as writer:
    # Write each DataFrame to a different sheet
    call_8.to_excel(writer, sheet_name='29300(a)(1)', index=False)
    wf2.to_excel(writer, sheet_name='29300(a)(2)', index=False)
    esir2.to_excel(writer, sheet_name='29300(b)', index=False)
    
#Write to file with Ray
# This has to change to CSV files to get uploaded to SharePoint
# with pd.ExcelWriter('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_' + as_of + '.xlsx', engine='openpyxl') as writer:
#     # Write each DataFrame to a different sheet
#     call_8.to_excel(writer, sheet_name='29300(a)(1)', index=False)
#     wf2.to_excel(writer, sheet_name='29300(a)(2)', index=False)
#     esir2.to_excel(writer, sheet_name='29300(b)', index=False)


#Append wf2 and esir2 to call_8 which should have the reporting criteria built into the notification type
call_9 = call_8.append(wf2, ignore_index = True, sort = False)
call_9 = call_9.append(esir2, ignore_index = True, sort = False)
del call_9['ID']
    
#New format saved to Shared Drive
call_9['NotificationDate'] = pd.to_datetime(call_9['NotificationDate'], errors='coerce').dt.strftime('%Y/%m/%d %H:%M')
call_9.to_csv('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/SCE_' + as_of + '-1-29300.csv', index=False)


#New format to upload to SharePoint
call_9.to_csv('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/SCE_' + as_of + '-1-29300.csv', index=False)
call_9.to_csv('C:/Users/brownsjm/Office of Energy Infrastructure Safety/Energy Safety - External Stakeholders - GIS QDR Submissions - 29300 Notifications/SCE_' + as_of + '-1-29300.csv', index=False)


# call_8.to_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/OEIS_' + as_of + '.xlsx', sheet_name = '29300(a)(1)')
# call_8.to_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_' + as_of + '.xlsx', sheet_name = '29300(a)(1)')

# master_list = pd.read_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/OEIS_Master.xlsx', sheet_name = '29300(a)(1)')
# master_list = pd.read_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', sheet_name = '29300(a)(1)')

# =============================================================================
# Append data to all lists so call8 gets a history and wf and esir will not keep sending the line item
# =============================================================================
#Call8 appended to master list
call_8 = call_8.rename(columns = {"index" : "ID"})
master_list2 = master_list.append(call_8, ignore_index = True, sort = False)
# del master_list2['Unnamed: 0']

master_list2 = master_list2.sort_values(['NotificationDate']).drop_duplicates('DOutageID')

# wf events new appended to master list
master_list_a2_hist = master_list_a2.append(wf2, ignore_index = True, sort = False)

# esir events new appended to master list
master_list_b_hist = master_list_b.append(esir2, ignore_index = True, sort = False)


#write to PMA folder
# with pd.ExcelWriter('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/OEIS_Master.xlsx', engine='openpyxl') as writer:
#     # Write each DataFrame to a different sheet ((Need to update to new version once wf and esir would be appended))
#     master_list2.to_excel(writer, sheet_name='29300(a)(1)', index=False)
#     master_list_a2_hist.to_excel(writer, sheet_name='29300(a)(2)', index=False)
#     master_list_b_hist.to_excel(writer, sheet_name='29300(b)', index=False)


#Write to FIPA-Reporting folder ((Need to update to new version once wf and esir would be appended))
# with pd.ExcelWriter('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', engine='openpyxl') as writer:
#     # Write each DataFrame to a different sheet
#     master_list2.to_excel(writer, sheet_name='29300(a)(1)', index=False)
#     master_list_a2_hist.to_excel(writer, sheet_name='29300(a)(2)', index=False)
#     master_list_b_hist.to_excel(writer, sheet_name='29300(b)', index=False)
    
# master_list2.to_excel('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/OEIS_Master.xlsx', sheet_name = '29300(a)(1)')
# master_list2.to_excel('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/OEIS_Master.xlsx', sheet_name = '29300(a)(1)')

print('Completed OEIS_OMS.py')

# =============================================================================
# Test out writing directly to SharePoint
# =============================================================================

# call_9_test = pd.read_csv('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/SCE_2026-02-09-1-29300.csv')

# call_9_test.to_csv('C:/Users/brownsjm/Office of Energy Infrastructure Safety/Energy Safety - External Stakeholders - GIS QDR Submissions - 29300 Notifications/SCE_' + as_of + '-1-29300test.csv', index=False)