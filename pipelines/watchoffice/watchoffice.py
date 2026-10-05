from datetime import datetime
import os
from pprint import pprint
import sys
import tempfile
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import openpyxl
import numpy as np
import re


PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from pipelines.watchoffice.extract import extract_incident_data
from shared.logging import setup_logging
from shared.sharepoint.graph_api import authenticate_user, pull_sharepoint_excel_file
from shared.sharepoint.utils import download_file, upload_file
from shared.sharepoint.utils import upload_dataframe_to_sharepoint_list
from shared.utils import distribution_circuits


def _remove_processed_source_rows(workbook_path, processed_source_rows):
    workbook = load_workbook(workbook_path)
    try:
        worksheet = workbook.worksheets[0]
        source_table = None
        for table in worksheet.tables.values():
            min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(table.ref)
            headers = [worksheet.cell(min_row, column).value for column in range(min_col, max_col + 1)]
            if min_row == 1 and 'WatchOffice' in headers:
                source_table = table
                break
        if source_table is None:
            raise RuntimeError('The source workbook must contain an existing WatchOffice table on the first sheet.')

        totals_rows = int(bool(source_table.totalsRowCount))
        rows_to_delete = sorted({int(row) + 2 for row in processed_source_rows}, reverse=True)
        if any(row <= min_row or row > max_row - totals_rows for row in rows_to_delete):
            raise ValueError('Processed source rows fall outside the WatchOffice table.')

        for row in rows_to_delete:
            worksheet.delete_rows(row)

        new_max_row = max_row - len(rows_to_delete)
        if new_max_row < min_row + 1 + totals_rows:
            if totals_rows:
                worksheet.insert_rows(min_row + 1)
            new_max_row = min_row + 1 + totals_rows
        start_cell = f'{openpyxl.utils.get_column_letter(min_col)}{min_row}'
        end_column = openpyxl.utils.get_column_letter(max_col)
        source_table.ref = f'{start_cell}:{end_column}{new_max_row}'
        if source_table.autoFilter is not None:
            source_table.autoFilter.ref = f'{start_cell}:{end_column}{new_max_row - totals_rows}'
        workbook.save(workbook_path)
    finally:
        workbook.close()


logger = setup_logging("watchoffice")
logger.info('Running outlook_WatchOffice_fire.py')

# =============================================================================
# Read in WatchOffice Data
# =============================================================================

site_url = 'https://edisonintl.sharepoint.com/teams/grpmo/FIPA'
sharepoint_drive = 'Documents'
sharepoint_file_path = '/MPR - FIPA/FIPA_Reporting/power_app_automation/WatchOffice'
sharepoint_file_name = 'Prelim_WatchOffice_Data.xlsx'
source_row_column = '_watchoffice_source_row'
sharepoint_folder_url = (
    f'https://edisonintl.sharepoint.com/teams/grpmo/FIPA/'
    f'{sharepoint_drive}/{sharepoint_file_path.strip("/")}'
)

logger.info('Authenticating to SharePoint for WatchOffice data.')
authenticate_user()
logger.info('Loading WatchOffice workbook from SharePoint: %s/%s', sharepoint_file_path, sharepoint_file_name)
read_watchoffice = pull_sharepoint_excel_file(
    site_path='/teams/grpmo/FIPA',
    drive_name=sharepoint_drive,
    file_path=sharepoint_file_path,
    file_name=sharepoint_file_name
)
source_is_sharepoint = not read_watchoffice.empty

if read_watchoffice.empty:
    logger.warning('SharePoint file was empty; falling back to local synced workbook path.')
    local_source_path = Path(
        'C:/Users/duganpr/Southern California Edison/'
        'FIPA - FIPA_Reporting/power_app_automation/WatchOffice/'
        'Prelim_WatchOffice_Data.xlsx'
    )
    read_watchoffice = pd.read_excel(local_source_path)

source_workbook = read_watchoffice.copy()
source_workbook[source_row_column] = range(len(source_workbook))
read_watchoffice = source_workbook.copy()

read_watchoffice = read_watchoffice.drop_duplicates()

# Remove rows where all values are missing or empty
read_watchoffice = read_watchoffice.dropna(how='all')

logger.info('Loaded WatchOffice dataset with %s rows.', len(read_watchoffice))

# Remove rows where all string elements have length zero
read_watchoffice = read_watchoffice.map(lambda x: x.strip() if isinstance(x, str) else x)
read_watchoffice = read_watchoffice.loc[(read_watchoffice.astype(bool).sum(axis=1) != 0)]

# Reset the index of the DataFrame
read_watchoffice = read_watchoffice.reset_index(drop=True)

read_watchoffice = read_watchoffice[~read_watchoffice['WatchOffice'].isna()]
read_watchoffice = read_watchoffice[~read_watchoffice['WatchOffice'].isnull()]
read_watchoffice = read_watchoffice[read_watchoffice['WatchOffice'] != '']
read_watchoffice = read_watchoffice[read_watchoffice['WatchOffice'] != ' ']
read_watchoffice = read_watchoffice[read_watchoffice['WatchOffice'] != 'nan']

read_watchoffice = read_watchoffice[[
    source_row_column,
    'Entry_Method',
    'email_date',
    'Reported_By_Email',
    'WatchOffice',
]]

# =============================================================================
# Set excel file in SharePoint where to open excel and paste WatchOffice information
# =============================================================================

# sharepoint_folder = 'Repair Order Event Tracking/Keyword Scrape'
# logger.info('Downloading workbook template from SharePoint folder: %s', sharepoint_folder)
# temp_dir = tempfile.TemporaryDirectory()
# temp_path = Path(temp_dir.name)
# output_file = temp_path / 'WatchOffice_Data.xlsx'
# if not download_file(output_file.name, sharepoint_folder, output_file):
#     logger.error('Unable to download %s from SharePoint', output_file.name)
#     raise FileNotFoundError(f"Unable to download {output_file.name} from SharePoint")
# output_wb = load_workbook(output_file)
# output_ws = output_wb['Sheet1']
# output_wb.active = output_ws

#Load in files with additional information to join to circuit 
# ckt_list = pd.read_csv('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/Jbrow/PowerAppAutomation/Distribution Circuit List.csv')
ckt_list = distribution_circuits()
ckt_list = ckt_list[['Title', 'SAP_HFRA', 'Switching Center']]
ckt_dist = pd.read_csv('data/Circuits_District.csv')
ckt_dist = ckt_dist[['CKT_NAME', 'DISTRICT','KVOLT', 'SUBSTATION', 'Region Name']]
# dist_inbox = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Keyword Scrape/District Inboxes.xlsx')
# dist_inbox = pd.read_excel('C:/Users/brownsjm/Southern California Edison/Failure Event Analysis - Documents (1)/Repair Order Event Tracking/Keyword Scrape/District Inboxes.xlsx')
from shared.utils import district_inboxes
dist_inbox = district_inboxes()
dist_inbox = dist_inbox[['District Number', 'District Name', 'Engineer Assigned']]

today = datetime.today()
today_year = today.year
try:
    new_rows = []
    for _, wo in read_watchoffice.iterrows():
        data = extract_incident_data(
            raw_text=wo['WatchOffice'],
            email_date=wo['email_date']
        )

        incident_date_raw = data.get('incident_date')
        incident_time_raw = data.get('incident_time')

        if incident_date_raw in (None, '', 'Unknown', 'unknown'):
            parsed_email_date = pd.to_datetime(wo.get('email_date'), errors='coerce')
            if pd.notna(parsed_email_date):
                incident_date_raw = parsed_email_date.strftime('%Y-%m-%d')

        if incident_time_raw in (None, '', 'Unknown', 'unknown'):
            parsed_email_date = pd.to_datetime(wo.get('email_date'), errors='coerce')
            if pd.notna(parsed_email_date):
                incident_time_raw = parsed_email_date.strftime('%H:%M')

        has_known_date = incident_date_raw not in (None, '', 'Unknown', 'unknown')
        has_known_time = incident_time_raw not in (None, '', 'Unknown', 'unknown')
        incident_date = datetime.strptime(f"{incident_date_raw} {incident_time_raw}", '%Y-%m-%d %H:%M') if has_known_date and has_known_time else None
        circuit = (
            "Encrypted"
            if isinstance(data.get('incident_narrative'), str)
            and data['incident_narrative'].startswith("Southern California Edison Secure eMail Portal")
            else data['circuit']
        )

        new_rows.append({
            "Title": f"Watch_Office_{incident_date.strftime('%m/%d/%Y')}_{incident_date.strftime('%H%M')}" if incident_date else "Watch_Office_Unknown_Unknown",
            "Circuit": circuit,
            "Entry_Method": "WO",
            "Email Date": wo['email_date'],
            "Incident Date": incident_date,
            "Reported By Email": 'WatchOffice',
            "Comments": data['incident_narrative'],
            "HFRA": data['hfra']
        })

    df = pd.DataFrame(new_rows)

# =============================================================================
# Set excel file in SharePoint where to open excel and paste WatchOffice information
# =============================================================================
  
    df['Circuit_key'] = df['Circuit'].str.upper()
    df['Circuit_key'] = df['Circuit_key'].str.strip()

    ckt_list = ckt_list.rename(columns={'Title': 'Circuit_key'})

    df_2 = df.merge(
        ckt_list,
        on="Circuit_key",
        how="left"
    )  # .drop(['Circuit_key'], axis=1)  
    
    # df_2['Time'] = df_2['Incident_Time'].str.replace(":", '')
       
    # PowerAutomate changes to UTC so need to subtract 7 hours        
    # df_2['Incident_Time_adjust'] = pd.to_datetime(df_2['Incident_Date_time']).dt.tz_localize('US/Pacific').dt.tz_convert('UTC').dt.strftime('%m/%d/%Y %H:%M')
    # df_2['Incident_Date_time'] = df_2['Incident_Time_adjust'] + ' '
    # df_2['Incident_Date_time'] = df_2['Incident_Date_time'].str.strip()


    # Being to link other data to bring into lists for SharePoint
    
    def hfra(Tier, hfra):
        if Tier == 'Elevated':
            return 'T2'
        elif Tier == 'Extreme':
            return 'T3'
        elif Tier == '' or Tier == 'Missing' or Tier == np.nan:
            return 'NON-HFRA' if hfra == 'No' else hfra
        else:
            return 'NON-HFRA'

    df_2['HFRA Circuit'] = df_2.apply(lambda x: hfra(x['SAP_HFRA'], x['HFRA']), axis=1)
    df_2.drop(['SAP_HFRA', 'HFRA'], axis=1, inplace=True)

# =============================================================================
# Set excel file in SharePoint where to open excel and paste WatchOffice information
# =============================================================================
    df_3 = df_2.merge(
        ckt_dist,
        left_on= ['Circuit_key'],
        right_on= ['CKT_NAME'], how= 'left'
    ).drop(['CKT_NAME', 'Circuit_key'], axis=1)

# =============================================================================
# Set excel file in SharePoint where to open excel and paste WatchOffice information
# =============================================================================
    df_4 = df_3.merge(
        dist_inbox,
        left_on= ['DISTRICT'],
        right_on= ['District Number'], how= 'left'
    ).drop(['District Number'], axis=1)

    df_4['Engineer Assigned'] = (
        df_4['Engineer Assigned']
        .replace(r"^\s*$", pd.NA, regex=True)
        .fillna('Unassigned')
    )
    # df_4['DISTRICT'].fillna(99, inplace=True)
    # df_4.fillna('Unknown', inplace=True)

    print(df_4)
    df_4.to_excel('output.xlsx', index=False)
    # sys.exit()

    # Set this when the destination SharePoint list name is available.
    sharepoint_site_url_2 = 'https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis'
    sharepoint_list_name = 'Repair Order Events'

    # SharePoint internal field names may differ from DataFrame column names.
    # Add entries here when the destination list's internal names are known.
    sharepoint_field_mapping = {
        'Title': 'Title',
        'Circuit': 'Circuit',
        'Entry_Method': 'Entry_Method',
        'Email Date': 'Email_x0020_Date',
        'Incident Date': 'Incident_x0020_Date',
        'Reported By Email': 'Reported_x0020_By_x0020_Email',
        'Comments': 'Comments',
        'HFRA Circuit': 'HFRA_x0020_Circuit',
        'Switching Center': 'Switching_x0020_Center',
        'DISTRICT': 'District_x0020_Number',
        'District Name': 'District_x0020_Name',
        'KVOLT': 'Voltage_x0020__x0028_kV_x0029_',
        'SUBSTATION': 'Substation',
        'Region Name': 'Region',
        'Engineer Assigned': 'Engineer_x0020_Assigned'
    }

    if sharepoint_list_name:
        uploaded_count = upload_dataframe_to_sharepoint_list(
            sharepoint_site=sharepoint_site_url_2,
            sharepoint_list_name=sharepoint_list_name,
            dataframe=df_4,
            sharepoint_field_mapping=sharepoint_field_mapping,
        )
        if uploaded_count != len(df_4):
            raise RuntimeError(
                f'Expected to upload {len(df_4)} rows but uploaded {uploaded_count}.'
            )

        processed_source_rows = set(read_watchoffice[source_row_column])

        with tempfile.TemporaryDirectory() as temp_dir:
            if source_is_sharepoint:
                replacement_path = Path(temp_dir) / sharepoint_file_name
                if not download_file(sharepoint_file_name, sharepoint_folder_url, replacement_path):
                    raise RuntimeError('Unable to download the source workbook for row cleanup.')
                _remove_processed_source_rows(replacement_path, processed_source_rows)
                uploaded_source = upload_file(replacement_path, sharepoint_folder_url)
                if not uploaded_source:
                    raise RuntimeError(
                        'Destination upload succeeded, but the processed source '
                        'workbook could not be replaced.'
                    )
            else:
                    _remove_processed_source_rows(local_source_path, processed_source_rows)
        logger.info(
            'Removed %s successfully processed rows from %s.',
            len(processed_source_rows),
            sharepoint_file_name,
        )
    else:
        logger.warning(
            'SharePoint list upload skipped because sharepoint_list_name is not set.'
        )
    
    print('Completed running outlook_WatchOffice_fire.py')
except Exception as e:
    print('outlook_WatchOffice_fire.py failed to run')
    print(e)
    error_message = {
        'Error': ['outlook_WatchOffice_fire.py failed to run'],
                     'Details': [str(e)]
                     }
    raise
