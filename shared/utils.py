"""Utilities shared by Safari data pipelines."""

from typing import Any

import pandas as pd

from shared.sharepoint.graph_api import ensure_authenticated, pull_sharepoint_list_data


def district_inboxes() -> pd.DataFrame:
	"""Return all items from the District Inboxes SharePoint list as a DataFrame.

	Args:
		None

	Returns:
		pd.DataFrame: A DataFrame containing all items from the District Inboxes SharePoint list.
	"""
	ensure_authenticated()
	df = pull_sharepoint_list_data('https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis', 'District Inboxes')

	df = df[
		['fields.Title', 'fields.GMC_x0020_Zone', 'fields.qlqg', 'fields.District_x0020_Number', 'fields.District_x0020_Name', 'fields.Region', 'fields.Engineer_x0020_Assigned', 'fields.Wire_x0020_Down_x0020_Analyst_x0']
	].rename(columns={
        'fields.Title': 'District Inbox',
        'fields.GMC_x0020_Zone': 'GMC Zone',
        'fields.qlqg': 'Inbox Email',
        'fields.District_x0020_Number': 'District Number',
        'fields.District_x0020_Name': 'District Name',
        'fields.Region': 'Region',
        'fields.Engineer_x0020_Assigned': 'Engineer Assigned',
        'fields.Wire_x0020_Down_x0020_Analyst_x0': 'Wire Down Analyst Assigned'
    })
	return df


def system_assignments() -> pd.DataFrame:
	"""Return all items from the System Assignments SharePoint list as a DataFrame.

	Args:
		None

	Returns:
		pd.DataFrame: A DataFrame containing all items from the System Assignments SharePoint list.
	"""
	ensure_authenticated()
	df = pull_sharepoint_list_data('https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis', 'System Assignments')

	df = df[
		['fields.Title', 'fields.oqwq']
	].rename(columns={
        'fields.Title': 'Switching Center',
        'fields.oqwq': 'Engineer Assigned',
    })
	return df


def distribution_circuits() -> pd.DataFrame:
	"""Return all items from the Distribution Circuits SharePoint list as a DataFrame.

	Args:
		None

	Returns:
		pd.DataFrame: A DataFrame containing all items from the Distribution Circuits SharePoint list.
	"""
	ensure_authenticated()
	df = pull_sharepoint_list_data('https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis', 'Distribution Circuit List')

	df = df[
		['fields.Title', 'fields.field_1', 'fields.field_2', 'fields.field_3', 'fields.field_4', 'fields.field_5', 'fields.field_6']
	].rename(columns={
        'fields.Title': 'Title',
        'fields.field_1': 'Substation Name',
        'fields.field_2': 'Voltage (kV)',
        'fields.field_3': 'HFRA CIRCUIT?',
        'fields.field_4': 'HFRA TIERS',
        'fields.field_5': 'Switching Center',
        'fields.field_6': 'SAP_HFRA',
    })
	return df
