# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 06:35:14 2023

@author: brownsjm

Intent of report:
    
The intent of this automated report is to look into the Repair Order review event and equipment data to identify any Equipment failure that
involved covered conductor involved in an ignition in order to determine if a stated mitigation has not worked in preventing an ignition.
"""

import pandas as pd
import datetime as dt
import warnings
import pyodbc

# PRINT TO SHOW SCRIPT HAS BEEN LOADED
print('RUNNING CC_Failure_Review.py')

#from dplython import (DplyFrame, X, diamonds, select, sift, sample_n, sample_frac, head, arrange, mutate, group_by,
#                      summarize, DelayFunction)
warnings.filterwarnings('ignore')
#get system date to define yesterday's date and pull all outages from yesterday

today = dt.datetime.today()
yesterday = today - dt.timedelta(days=1)
yesterday_year = yesterday - dt.timedelta(days=365)
yesterday_year_year = yesterday_year.year
today_year=today.year

yesterday_year_year = str(yesterday_year_year)

#Only take events that happened in the last week since run every Friday
last_week = today - dt.timedelta(days=7)

as_of = pd.Timestamp.today().strftime('%m_%d_%Y')

#%%Load SQL Event info, bounce against list to filter down on events to write to event tables

# SQL Server connection info
DB_DRIVER = "{SQL Server}"  # Or whatever driver is installed in your machine
DB_SERVER = 'tcp:D259321,49172'  # ip address or name\\instance.  include port if any
DATABASE_NAME = 'RepairOrder'  # development server, for testing

connection = pyodbc.connect(f"DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DATABASE_NAME};Trusted_Connection=no;")

#repair_order view all from SQL
ro_q = '''SELECT DISTINCT * FROM dbo.vw_AllRepairOrderData
            WHERE Category = 'Conductor'
            AND Specifics = 'Covered'
            AND "Voltage Classification" = 'Primary';'''
ro = pd.read_sql(ro_q, connection)

ro = ro.rename(columns = {"Voltage Classification" : "Voltage_class"})

#sanity check on duplicate id's
ro_event_list = ro[['ID']].drop_duplicates()

#Capture if RO had a different primary equipment from the RO AL Covered Conductor
prim_equip_q = '''SELECT DISTINCT ID, Category, Specifics, "Voltage Classification", "Material" FROM dbo.vw_AllRepairOrderData
            WHERE "Primary Failed Equipment" = 'True';'''
prim_equip = pd.read_sql(prim_equip_q, connection)

prim_equip = prim_equip.rename(columns = {"Voltage Classification" : "Prim_equip_volt_class",
                                          "Category" : "Prim_equip_category",
                                          "Specifics" : "Prim_equip_specifics",
                                          "Material" : "Prim_equip_material"})

ro_1 = pd.merge(ro, prim_equip, left_on= ['ID'],
                         right_on= ['ID'], how='left')

#Only want fire events
ro_1 = ro_1[ro_1['Event Type'].str.contains("Fire", na=False)]

ro_1['2 Week'] = ro_1["Incident Date"] + pd.Timedelta(weeks=2)

#Only need some columns
fire_cc_1 = ro_1[['ID', 'Incident Date', '2 Week', 'Structure', 'Category', 'Specifics', 'Material', 'size', 'Prim_equip_volt_class', 'Prim_equip_category', 'Prim_equip_specifics', 'Prim_equip_material', 'Event Driver', 'Event Driver Specifics','HFRA Circuit', 'Circuit', 'Voltage (kV)', 'Substation', 'District Name', 'District Number', 'Switching Center', 'Reviewed Date', 'Event Type', 'RO Problem', 'RO Cause', 'CAD ID', 'Comments']]


fire_cc_1 = fire_cc_1.rename(columns = {"Category" : "Equipment",
                                  "Voltage (kV)" : "Voltage",
                                  "Specifics" : "Type"})


fire_cc_1 = fire_cc_1[(fire_cc_1['Reviewed Date'] >= last_week)]

del fire_cc_1['Reviewed Date']

# fire_cc_1.to_excel(r'C:/Users/brownsjm/Downloads/wd_CC_070926_WD.xlsx', index=False)

if len(fire_cc_1) > 0:
    fire_cc_1.to_excel('C:/Users/duganpr/Southern California Edison/FIPA - FIPA_Reporting/CC_Failure_Review_PD/CC_Failure_Review_' + as_of + '.xlsx', sheet_name = 'Fire_CC', index=False)
else:
    fire_cc_1.to_excel('C:/Users/duganpr/Southern California Edison/FIPA - FIPA_Reporting/CC_Failure_Review_No_Events_PD/CC_Failure_Review_' + as_of + '.xlsx', sheet_name = 'Fire_CC', index=False)

    
print('Completed CC_Failure_Review.py Processing')