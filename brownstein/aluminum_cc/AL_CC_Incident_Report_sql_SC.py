# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 06:35:14 2023

@author: brownsjm

Intent of report:
    
The intent of this automated report is to look into the Repair Order review event and equipment data to identify any Equipment failure that
involved covered conductor that was also a safety event defined by a wire down or ignition flag.  This is based on engineering review 
of repair orders and flags that they have for the data.  This doesnot reference CC failure but any equipment failure where the CC 
was impacted.
"""

import numpy as np
import pandas as pd
import datetime as dt
import warnings
import pyodbc

# PRINT TO SHOW SCRIPT HAS BEEN LOADED
print('RUNNING AL_CC_Incident_Report.py')

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

#Read in event data, equipment data, ro scrape data, and coastal structures list used from a structure geospatial join
coast_struct = pd.read_csv('data/Structure_within_coast.csv')

#%%Load SQL Event info, bounce against list to filter down on events to write to event tables

# SQL Server connection info
DB_DRIVER = "{SQL Server}" # Or whatever driver is installed in your machine
DB_SERVER = 'tcp:D259321,49172' # ip address or name\\instance.  include port if any
DATABASE_NAME = 'RepairOrder'  # development server, for testing

connection = pyodbc.connect(f"DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DATABASE_NAME};Trusted_Connection=no;")

#repair_order view all from SQL
ro_q = '''SELECT DISTINCT * FROM dbo.vw_AllRepairOrderData
            WHERE Category = 'Conductor'
            AND Specifics = 'Covered'
            AND "Voltage Classification" = 'Primary'
            AND (Material = 'Aluminum' OR Material = 'AL');'''
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

#Only want consequence events
ro_1 = ro_1[ro_1['Event Type'].str.contains("Fire") | ro_1['Event Type'].str.contains("Wire Down")]

ro_1['2 Week'] = ro_1["Incident Date"] + pd.Timedelta(weeks=2)

#Only need some columns
al_cc_1 = ro_1[['ID', 'Incident Date', '2 Week', 'Structure', 'Category', 'Specifics', 'Material', 'size', 'Prim_equip_volt_class', 'Prim_equip_category', 'Prim_equip_specifics', 'Prim_equip_material', 'Event Driver', 'Event Driver Specifics','HFRA Circuit', 'Circuit', 'Voltage (kV)', 'Substation', 'District Name', 'District Number', 'Switching Center', 'Reviewed Date', 'Event Type', 'RO Problem', 'RO Cause', 'CAD ID', 'Comments']]


def corrosion_Al(Event_Driver, prim_equip):
    if Event_Driver == 'Corrosion' and prim_equip == 'Conductor':
        return 'Y'
    else:
            return 'N'

al_cc_1['Al Conductor Corrosion'] = al_cc_1.apply(lambda x: corrosion_Al(x['Event Driver Specifics'],x['Prim_equip_category']), axis=1)


al_cc_1 = al_cc_1.rename(columns = {"Category" : "Equipment",
                                  "Voltage (kV)" : "Voltage",
                                  "Specifics" : "Type"})

#Remove CFO events as this isnt how constructed
al_cc_1 = al_cc_1[al_cc_1['Event Driver'] != 'Contact From Object']

#Add in if structure on coast
coast_struct_1 = coast_struct[['STRUCTURE_NUMBER']].reset_index()
coast_struct_1['Coast_Structure'] = 'Y'

al_cc_2 = pd.merge(al_cc_1, coast_struct_1, left_on= ['Structure'],
                         right_on= ['STRUCTURE_NUMBER'], how='left').drop('STRUCTURE_NUMBER', axis = 1)

al_cc_2['Coast_Structure'] = al_cc_2['Coast_Structure'].fillna('N')


al_cc_2 = al_cc_2[(al_cc_2['Reviewed Date'] >= last_week)]

del al_cc_2['Reviewed Date']

if len(al_cc_2) > 0:
    al_cc_2.to_excel('C:/Users/duganpr/Southern California Edison/FIPA - FIPA_Reporting/SED_AL_CC_Report_PD/AL_CC_' + as_of + '.xlsx', sheet_name = 'AL_CC', index=False)
else:
    al_cc_2.to_excel('C:/Users/duganpr/Southern California Edison/FIPA - FIPA_Reporting/SED_AL_CC_Report_No_Events_PD/AL_CC_No_Event_' + as_of + '.xlsx', sheet_name = 'AL_CC', index=False)

    
print('Completed AL_CC_Report.py Processing')
