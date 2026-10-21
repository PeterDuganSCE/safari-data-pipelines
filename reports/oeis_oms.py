import sys
from pathlib import Path
import re
from urllib.parse import quote, urlparse

from hdbcli import dbapi
import pandas as pd
import datetime as dt
import saspy
import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared.sharepoint.graph_api import authenticate_user, pull_sharepoint_excel_file
from shared.sharepoint.utils import upload_file
from shared.sharepoint.graph import GraphClient
from shared.logging import setup_logging
from shared.paths import AUTH_PATH, CONFIG_PATH

# ---------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------
logger = setup_logging("safari_to_hana_ignition")


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
logger.info("Project root directory: %s", PROJECT_ROOT)
logger.info("Using config file: %s", CONFIG_PATH)
logger.info("Using auth file: %s", AUTH_PATH)

fipa_sharepoint_url = "https://edisonintl.sharepoint.com/teams/grpmo/FIPA"
oeis_sharepoint_folder = "MPR - FIPA/FIPA_Reporting/OEIS_Report"

def main():
    # authenticate_user()
    last_run_date = get_last_run_date(
        site_url=fipa_sharepoint_url,
        folder_path=oeis_sharepoint_folder
    )
    logger.info("Last run date: %s", last_run_date)
    circuit_df, call_1, call_text_1 = get_hana_data(last_run_date)
    call_4, ckt_nam, sub = get_sas_data(call_1)
    logger.info("Saving call_4 to output/call_4.xlsx")
    call_4.to_excel('output/call_4.xlsx', index=False)
    call_6 = merge_data(call_4, call_text_1)
    call_6.to_excel('output/call_6.xlsx', index=False)
    call_8 = format_data(call_6, ckt_nam, sub)
    call_8.to_excel('output/call_8.xlsx', index=False)
    wf2, esir2 = read_in_manual_entry()
    upload_to_oeis(call_8, wf2, esir2)


def get_last_run_date(
    site_url: str,
    drive_name: str = "Documents",
    folder_path: str = "",
) -> dt.datetime:
    """Return the date embedded in the most recently modified SharePoint CSV."""
    parsed_site_url = urlparse(site_url)
    if not parsed_site_url.netloc or not parsed_site_url.path:
        raise ValueError(f"Invalid SharePoint site URL: {site_url!r}")

    graph = GraphClient(
        app_config_path=str(PROJECT_ROOT / "config" / "sharepoint_auth.json")
    )
    site = graph.get(
        f"sites/{parsed_site_url.netloc}:{parsed_site_url.path.rstrip('/')}"
    )
    site_id = site["id"]

    drives = graph.get_all(f"sites/{site_id}/drives")
    drive = next(
        (
            item
            for item in drives
            if item.get("name", "").casefold() == drive_name.casefold()
        ),
        None,
    )
    if drive is None:
        raise ValueError(f"SharePoint drive not found: {drive_name!r}")

    normalized_folder = folder_path.strip("/")
    if normalized_folder:
        encoded_folder = quote(normalized_folder, safe="/")
        endpoint = (
            f"sites/{site_id}/drives/{drive['id']}"
            f"/root:/{encoded_folder}:/children"
        )
    else:
        endpoint = f"sites/{site_id}/drives/{drive['id']}/root/children"

    csv_files = [
        item
        for item in graph.get_all(endpoint)
        if item.get("file") and item.get("name", "").lower().endswith(".csv")
    ]
    if not csv_files:
        raise FileNotFoundError(
            f"No CSV files found in SharePoint folder {folder_path or '/'}"
        )

    latest_file = max(csv_files, key=lambda item: item["lastModifiedDateTime"])
    date_match = re.search(r"(?<!\d)(\d{4}-\d{2}-\d{2})(?!\d)", latest_file["name"])
    if date_match is None:
        raise ValueError(
            f"No YYYY-MM-DD date found in latest CSV filename: {latest_file['name']}"
        )

    return dt.datetime.strptime(date_match.group(1), "%Y-%m-%d")
    # return dt.datetime.strptime("2026-10-09", "%Y-%m-%d")


def get_hana_data(last_run_date):
    """
    Placeholder function to demonstrate how to retrieve data from HANA.
    Replace this with actual HANA connection and query logic.
    """

    last_date = pd.to_datetime(last_run_date)

    #get system date to define yesterday's date and pull all outages from yesterday
    today = dt.datetime.today()

    delta = today - last_date
    delta_days = delta.days

    before_weekend = today - dt.timedelta(days=delta_days)
    before_weekend_year = before_weekend - dt.timedelta(days=365)
    before_weekend_year_year = before_weekend_year.year
    today_year=today.year

    # pole_census = pd.read_csv('data/Pole_Census_SurfaceFuels_20230308.csv')
    # pole_census_1 = pole_census[['SCE_STRUCTURE_NO', 'Definition', 'URBANRURAL']]


    with open(AUTH_PATH, "r", encoding="utf-8") as auth_file:
        hana_credentials = (yaml.safe_load(auth_file) or {})["hana"]

    connection = dbapi.connect(
        address="vp55db51.sce.com",
        port=30015,
        user=hana_credentials["username"],
        password=hana_credentials["password"],
        databasename="P55"
    )

    circuit_query = (
        '''
        select "CIRCT_ID" AS "CKT_ID", "CIRCT_NAM", "CIRCT_DESC" from "OMS"."CIRCUIT";
        '''
    )
    call_query = (
        '''
        select CUST_NAM, CITY_NAM, CONTACT_NAM, CLUE_CD, CALL_TYPE_CD, AFF_DATETIME, STRCTUR_NO,
        CIRCT_ID, INCIDENT_ID from "OMS"."CALL"
        WHERE YEAR(AFF_DATETIME) >= ?
        AND "CLUE_CD" LIKE '%9';
        '''
    )
    call_text_query = (
        '''
        select INCIDENT_ID, COMMENT_TEXT, AFF_DATETIME from "OMS"."CALL"
        WHERE YEAR(AFF_DATETIME) >= ?
        AND "CLUE_CD" LIKE '%9';
        '''
    )
    his_call_query = (
        '''
        select CUST_NAM, CITY_NAM, CONTACT_NAM, CLUE_CD, CALL_TYPE_CD, AFF_DATETIME, STRCTUR_NO,
        CIRCT_ID, INCIDENT_ID from "OMS"."HIS_CALL"
        WHERE YEAR(AFF_DATETIME) >= ?
        AND "CLUE_CD" LIKE '%9';
        '''
    )
    his_call_text_query = (
        '''
        select INCIDENT_ID, AFF_DATETIME, CLUE_DESC1, TOT_LOSS_POWER_FLG, COMMENT_TEXT from "OMS"."HIS_CALL"
        WHERE YEAR(AFF_DATETIME) >= ?
        AND "CLUE_CD" LIKE '%9';
        '''
    )

    try:
        circuit_df = pd.read_sql_query(circuit_query, connection)
        year_params = (before_weekend_year_year,)
        call_df = pd.read_sql_query(call_query, connection, params=year_params)
        call_text_df = pd.read_sql_query(call_text_query, connection, params=year_params)
        his_call_df = pd.read_sql_query(his_call_query, connection, params=year_params)
        his_call_text_df = pd.read_sql_query(his_call_text_query, connection, params=year_params)

        call_df = call_df[call_df['AFF_DATETIME'] >= before_weekend]
        call_text_df = call_text_df[call_text_df['AFF_DATETIME'] >= before_weekend]
        his_call_df = his_call_df[his_call_df['AFF_DATETIME'] >= before_weekend]
        his_call_text_df = his_call_text_df[his_call_text_df['AFF_DATETIME'] >= before_weekend]

        # del call_df['AFF_DATETIME']
        del call_text_df['AFF_DATETIME']
        del his_call_text_df['AFF_DATETIME']
        # del his_call_df['AFF_DATETIME']

        call_1 = pd.concat([call_df, his_call_df], axis=0, sort= True).drop_duplicates()
        call_text_1 = pd.concat([call_text_df, his_call_text_df], axis=0, sort= True).drop_duplicates()

        call_1['CITY_NAM'] = call_1['CITY_NAM'].str.upper()
        call_1['CITY_NAM'] = call_1['CITY_NAM'].str.replace('COUNTY OF ', '')

        #Add county from flat table for city to county information (wikipedia)
        city_county = pd.read_excel('data/City_County.xlsx')
        call_1 = pd.merge(call_1, city_county, left_on= ['CITY_NAM'],
                         right_on= ['City'], how='left').drop(['City'], axis=1)

        #Get circuit name from circuit table in OMS
        call_1 = pd.merge(call_1, circuit_df, left_on= ['CIRCT_ID'],
                         right_on= ['CKT_ID'], how='left').drop(['CKT_ID'], axis=1)
        
        return circuit_df, call_1, call_text_1
    finally:
        connection.close()


def get_sharepoint_data():
    """
    Placeholder function to demonstrate how to retrieve data from SharePoint.
    Replace this with actual SharePoint connection and query logic.
    """
    # Example: Connect to SharePoint and fetch data
    # sharepoint_client = SharePointClient(...)
    # data = sharepoint_client.get_list_items("your_list")
    # return data
    pass

def get_sas_data(call_1: pd.DataFrame):
    """
    Placeholder function to demonstrate how to retrieve data from SAS.
    Replace this with actual SAS connection and query logic.
    """
    # Example: Connect to SAS and fetch data
    # sas_client = SASClient(...)
    # data = sas_client.get_data("your_dataset")
    # return data
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

        CALL_2 = sas.df2sd(call_1, 'CALL_2')

        query_1 = (
            """
            PROC SQL;
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
            """
        )

        query_2 = (
            """
            PROC SQL;
            CREATE TABLE WORK.CALL_3 AS
            SELECT DISTINCT t1.*, t2.*
            FROM WORK.CALL_2 t1
            LEFT JOIN WORK.STRUCT t2 ON t1.STRCTUR_NO = t2.FLOC_StructureID;
            QUIT;
            """
        )

        query_3 = (
            """
            PROC SQL;
            CREATE TABLE WORK.WEATHER AS
            SELECT DISTINCT FLOC, STATION_NUMBER, SUBSTR(FLOC,INDEX(FLOC,"-")+1,20) AS FLOC_2
            FROM TDSTAGNG.STG_SCE_WS_TO_FLOC_MAPPING;
            QUIT;
            """
        )

        query_4 = (
            """
            PROC SQL;
            CREATE TABLE WORK.CALL_4 AS
            SELECT DISTINCT t1.*, t2.STATION_NUMBER
            FROM WORK.CALL_3 t1
            LEFT JOIN WORK.WEATHER t2 ON t1.STRCTUR_NO = t2.FLOC_2;
            QUIT;
            """
        )

        query_5 = (
            """
            PROC SQL;
            CREATE TABLE WORK.CKT_NAM AS
            SELECT DISTINCT Circuit_Number AS CKT_NUM, Circuit_Name AS Circuit
            FROM TDSTAGNG.CIRCUIT_NAMENUMBER;
            QUIT;
            """
        )

        query_6 = (
            """
            PROC SQL;
            CREATE TABLE WORK.SUB AS
            SELECT DISTINCT CKT_NAME, SUBSTATION, SUB_NO, Region_Name AS Region_Name
            FROM TDSTAGNG.CIRCUITSUBSTAT;
            QUIT;
            """
        )

        c = sas.submitLST(query_1)
        c = sas.submitLST(query_2)
        c = sas.submitLST(query_3)
        c = sas.submitLST(query_4)
        c = sas.submitLST(query_5)
        c = sas.submitLST(query_6)

        call_4 = sas.sd2df('CALL_4', libref='WORK', method='CSV')
        ckt_nam = sas.sd2df('CKT_NAM', libref='WORK', method='CSV')
        sub = sas.sd2df('SUB', libref='WORK', method='CSV')

        del call_4['FLOC_StructureID']

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

        return call_4, ckt_nam, sub
    except Exception as e:
        logger.error(f"Error occurred while fetching SAS data: {e}")
        return None, None, None
    finally:
        sas.endsas()


def merge_data(
        call_4: pd.DataFrame,
        call_text_1: pd.DataFrame,
    ):

    # Start matching original OEIS components
    # Merge in clue desc, tot_loss_flg, and comment text
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
    pole_census = pd.read_csv('data/Pole_Census_SurfaceFuels_20230308.csv')
    pole_census_1 = pole_census[['SCE_STRUCTURE_NO', 'Definition', 'URBANRURAL']]

    call_6 = pd.merge(call_5, pole_census_1, left_on= ['STRCTUR_NO'],
                            right_on= ['SCE_STRUCTURE_NO'], how='left').drop(['SCE_STRUCTURE_NO'], axis=1)

    call_6['URBANRURAL'] = call_6['URBANRURAL'].fillna('Undetermined at time of reporting')
    call_6 = call_6.rename(columns = {"URBANRURAL" : "Origin Land Use"})

    call_6['Definition'] = call_6['Definition'].fillna('Undetermined at time of reporting')
    call_6 = call_6.rename(columns = {"Definition" : "Fuel Bed Description"})
    call_6['Fuel Bed Description Comment'] = "N/A"

    #Start matching original OEIS components
    today = dt.datetime.today()
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

    return call_6


def format_data(
        call_7: pd.DataFrame,
        ckt_nam: pd.DataFrame,
        sub: pd.DataFrame,
    ) -> pd.DataFrame:

    ckt_sub = pd.merge(sub, ckt_nam, left_on= ['CKT_NAME'],
                        right_on= ['Circuit'], how='left').drop(['Circuit'], axis=1)

    call_7 = pd.merge(call_7, ckt_sub, left_on= ['CIRCT_NAM'],
                      right_on= ['CKT_NAME'], how= 'left').drop(['CKT_NAME'], axis=1)

    call_7['TOT_LOSS_POWER_FLG'] = call_7['TOT_LOSS_POWER_FLG'].fillna('Unknown')
    call_7['CLUE_DESC1'] = "Undetermined at time of reporting"

    call_7 = call_7.rename(
        columns = {
            "SUB_NO" : "Substation",
            "STATION_NUMBER" : "Nearest Weather Station ID",
            "WEATHER_STN_ID" : "Nearest Weather Station ID",
            "FLOC_District" : "District",
            "CLUE_DESC1" : "SuspectedInitiatingCause",
            "COMMENT_TEXT" : "Additional Notes",
            "INCIDENT_ID" : "Doutage ID",
            "CKT_NUM" : "Circuit ID",
            "Voltage" : "Circuit Voltage"
        }
    )

    # Re-order to match OEIS Report
    call_8 = call_7[
        ['Notification Date', 'Notification Type', 'Utility ID', 'Confidential', 'Incident Start Date and Time',
        'County', 'District', 'Latitude', 'Longitude', 'HFTD Class', 'Origin Land Use', 'Material at Origin',
        'Material at Origin Comment', 'Fuel Bed Description', 'Fuel Bed Description Comment', 'Circuit ID',
        'Circuit Voltage', 'Substation', 'Nearest Weather Station ID', 'RFW Status', 'RFW Issue Date and Time',
        'FWW Status', 'FWW Issue Date and Time', 'HWW Status', 'HWW Issue Date and Time', 'Detection Method',
        'Detection Method Comment', 'Fire Size', 'Suppressed By', 'Suppressing Agency', 'Fire Investigation',
        'Fire AHJ', 'Outage Status', 'Toutage ID', 'Doutage ID', 'SuspectedInitiatingCause', 'SuspectedInitiatingCauseComment',
        'Determination', 'Determination Comment', 'Equipment Failure', 'Equipment Failure Comment', 'Object Contact',
        'Object Contact Comment', 'Facility Contacted', 'Facility Contacted Comment', 'Additional Notes'
        ]
    ].reset_index()

    call_8.to_excel('output/call_8.xlsx', index=False)

    call_8 = call_8.rename(
        columns = {
            "index" : "ID",
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
            "Additional Notes" : "AdditionalNotes"
        }
    )
    
    call_8['Temp'] = call_8['IncidentStartDateTime'].dt.strftime('%m/%d/%Y')
    call_8['AdditionalNotes'] = "An event occurred on " + call_8['Temp'] + " that impacted SCE facilities in our " + call_8['District'] + " District on circuit " + call_8['CircuitID'] + "."
    del call_8['Temp']
    
    return call_8


def read_in_manual_entry():
    authenticate_user()

    master_folder = "MPR - FIPA/FIPA_Reporting/OEIS_Report"
    master_file = "OEIS_Master.xlsx"
    master_list = pull_sharepoint_excel_file(
        site_path="/teams/grpmo/FIPA",
        drive_name="Documents",
        file_path=master_folder,
        file_name=master_file,
        sheet_name="29300(a)(1)",
    )
    master_list_a2 = pull_sharepoint_excel_file(
        site_path="/teams/grpmo/FIPA",
        drive_name="Documents",
        file_path=master_folder,
        file_name=master_file,
        sheet_name="29300(a)(2)",
    )
    master_list_b = pull_sharepoint_excel_file(
        site_path="/teams/grpmo/FIPA",
        drive_name="Documents",
        file_path=master_folder,
        file_name=master_file,
        sheet_name="29300(b)",
    )

    # Read in to see if anything manually added
    wf = pull_sharepoint_excel_file(
        site_path="/teams/grpmo/FIPA",
        drive_name="Documents",
        file_path="MPR - FIPA/FIPA_Reporting/OEIS_WF",
        file_name="WF_29300_a2.xlsx",
        sheet_name="29300(a)(2)",
    )
    # Read in to see if anything manually added
    esir = pull_sharepoint_excel_file(
        site_path="/teams/grpmo/FIPA",
        drive_name="Documents",
        file_path="MPR - FIPA/FIPA_Reporting/OEIS_ESIR",
        file_name="ESIR_29300_b.xlsx",
        sheet_name="29300(b)",
    )

    # if manually added there would be line items in wf2 and esir2, otherwise will be a blank file
    wf2 = wf[~wf['ID'].isin(master_list_a2['ID'])]
    esir2 = esir[~esir['ID'].isin(master_list_b['ID'])]

    master_list = master_list.loc[:, ~master_list.columns.str.contains('^Unnamed', case=False)]
    master_list_a2 = master_list_a2.loc[:, ~master_list_a2.columns.str.contains('^Unnamed', case=False)]
    master_list_b = master_list_b.loc[:, ~master_list_b.columns.str.contains('^Unnamed', case=False)]

    return wf2, esir2

def upload_to_oeis(call_8, wf2, esir2):
    as_of = pd.Timestamp.today().strftime('%Y-%m-%d')
    folder = f"output/oeis_uploads"
    filename = f"SCE_{as_of}-1-29300.csv"

    excel_path = Path(f"{folder}/SCE_{as_of}.xlsx")
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Write each DataFrame to a different sheet
        call_8.to_excel(writer, sheet_name='29300(a)(1)', index=False)
        wf2.to_excel(writer, sheet_name='29300(a)(2)', index=False)
        esir2.to_excel(writer, sheet_name='29300(b)', index=False)

    # Append manual entries to call_8 while preserving the existing row order.
    call_9 = pd.concat([call_8, wf2, esir2], ignore_index=True, sort=False)
    if 'ID' in call_9.columns:
        del call_9['ID']

    # New format saved to Shared Drive
    call_9['NotificationDate'] = pd.to_datetime(call_9['NotificationDate'], errors='coerce').dt.strftime('%Y/%m/%d %H:%M')
    # call_9.to_csv('//sce/workgroup/TDBU7/RESO-CRT/PERFORMANCE MANAGEMENT AND ANALYSIS/FIPA/OEIS_Reports/Peter/SCE_' + as_of + '-1-29300.csv', index=False)

    excel_path_2 = Path(f"{folder}/SCE_{as_of}-1-29300.csv")
    with pd.ExcelWriter(excel_path_2, engine='openpyxl') as writer:
        # Write each DataFrame to a different sheet
        call_9.to_excel(writer, sheet_name='in', index=False)

    authenticate_user()
    sharepoint_folder_url = (
        f"{fipa_sharepoint_url}/Documents/{oeis_sharepoint_folder}"
    )
    uploaded_url = upload_file(excel_path_2, sharepoint_folder_url)
    if not uploaded_url:
        raise RuntimeError(
            f"Failed to upload {excel_path_2.name} to {sharepoint_folder_url}"
        )
    logger.info("Uploaded %s to %s", excel_path_2, sharepoint_folder_url)

    # New format to upload to SharePoint
    # call_9.to_csv('C:/Users/brownsjm/Southern California Edison/FIPA - FIPA_Reporting/OEIS_Report/SCE_' + as_of + '-1-29300.csv', index=False)
    # call_9.to_csv('C:/Users/brownsjm/Office of Energy Infrastructure Safety/Energy Safety - External Stakeholders - GIS QDR Submissions - 29300 Notifications/SCE_' + as_of + '-1-29300.csv', index=False)

if __name__ == "__main__":
    main()
