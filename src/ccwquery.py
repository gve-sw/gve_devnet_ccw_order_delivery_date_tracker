#!/usr/bin/env python3
"""
Copyright (c) 2023 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

__author__ = "Trevor Maco <tmaco@cisco.com>"
__copyright__ = "Copyright (c) 2023 Cisco and/or its affiliates."
__license__ = "Cisco Sample Code License, Version 1.1"

import json
import os
import sys
from datetime import datetime, date, timezone
from multiprocessing import Pool

import pandas as pd
import requests
from dotenv import load_dotenv
from pandas.io.formats.style import jinja2
from rich.console import Console
from rich.panel import Panel

import ccwparser
import config

# Rich Console Instance
console = Console()

# Load ENV Variable
load_dotenv()
CLIENT_KEY = os.getenv("CLIENT_KEY")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Number of Workers for Processing
NUM_OF_WORKERS = 10


def get_access_token(client_id: str, client_secret: str) -> str:
    """
    Get Access Token for CCW Order API (see README to ensure proper access to API and the creation of an App)
    :param client_id: App Client ID
    :param client_secret: App Client Secret
    :return:
    """
    url = "https://id.cisco.com/oauth2/default/v1/token"
    headers = {
        'accept': "application/json",
        'content-type': "application/x-www-form-urlencoded",
        'cache-control': "no-cache"
    }
    payload = "client_id=" + client_id + \
              "&client_secret=" + client_secret + \
              "&grant_type=client_credentials"

    response = requests.request("POST", url, data=payload, headers=headers)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        response.raise_for_status()


def get_order_details(access_token: str, order_number: str) -> str | None:
    """
    Get Order Details from CCW API using the Order Number (Sales, Web, Purchase number)
    :param access_token: CCW API Access Token
    :param order_number: Order Number to query
    :return: Order Details in Str format, or None if Order not found
    """
    # Create Payload template
    raw_payload = """
    {
        "GetPurchaseOrder": {
            "value": {
                "DataArea": {
                    "PurchaseOrder": [
                        {
                            "PurchaseOrderHeader": {
                                "Description": [
                                    {
                                        "value": true,
                                        "typeCode": "details"
                                    }
                                ],
                                {% if reference_type == "PurchaseReference" %}
                                "ID": {
                                    "value": "{{ order_number }}"
                                }
                                {% else %}
                                "{{ reference_type }}": [
                                    {
                                        "ID": {
                                            "value": {{ order_number }}
                                        }
                                    }
                                ]
                                {% endif %}
                            }
                        }
                    ]
                },
                "ApplicationArea": {
                    "CreationDateTime": "datetime",
                    "BODID": {
                        "value": "Q4FY23-CCWDeliveryUpdater",
                        "schemeVersionID": "V1"
                    }
                }
            }
        }
    }
    """
    payload_template = jinja2.Template(raw_payload)

    # Modify payload based on Order Number type (Sales, Purchase, Web)
    payload_data = {}
    if config.ORDER_ID_TYPE == "Sales Order":
        payload_data['reference_type'] = "SalesOrderReference"
        payload_data['order_number'] = order_number
    elif config.ORDER_ID_TYPE == "Web Order":
        payload_data['reference_type'] = "DocumentReference"
        payload_data['order_number'] = order_number
    else:
        # Assume anything else is a Purchase Order Number
        payload_data['reference_type'] = "PurchaseReference"
        payload_data['order_number'] = order_number

    # Generate finalized payload query
    payload_str = payload_template.render(payload_data)
    payload = json.loads(payload_str)

    url = "https://apix.cisco.com/commerce/ORDER/v2/sync/checkOrderStatus"
    headers = {
        'authorization': "Bearer " + access_token,
        'accept': "application/json",
        'content-type': "application/json",
        'cache-control': "no-cache"
    }
    response = requests.request("POST", url, json=payload, headers=headers)

    if response.status_code == 200:
        # Check if order not found or unauthorized for access (error code OSA001)
        if 'OSA001' in response.json()['ShowPurchaseOrder']['value']['DataArea']['Show']['ResponseCriteria'][0][
            'ResponseExpression']['value']:
            return None
        else:
            return response.text
    else:
        response.raise_for_status()


def process_line_item(idx: int, row: pd.Series, order_number: str, access_token: str) -> dict:
    """
    Process Line Item from Excel Sheet, extract relevant data to query CCW API Order Details with, find line item, extract target field data
    :param idx: Index of the row in the Excel Sheet
    :param row: Row of the Excel Sheet
    :param order_number: Order Number to query
    :param access_token: CCW API Access Token
    :return: Dictionary of extracted data from CCW API Order Details
    """

    # Get the mandatory input values from the Excel columns
    ship_set = str(row[config.SHIP_SET_COLUMN_NAME]).strip()
    sku = str(row[config.SKU_COLUMN_NAME]).strip()

    # Populate line item dictionary with default values
    excel_line = {"idx": idx, "order_number": order_number, "sku": sku}
    for item in config.FIELDS_TO_TRACK:
        excel_line[item] = "No Data"

    # Skip rows with missing data
    if order_number == '' or sku == '' or ship_set == '':
        return excel_line

    # Get Order Details from CCW using the Sales Order Number
    order_details = get_order_details(access_token, order_number)

    # Order details are none, order = not found or unauthorized access
    if not order_details:
        for item in config.FIELDS_TO_TRACK:
            excel_line[item] = "Order Not Found"

        return excel_line

    # Build Order Parser Object, includes a number of methods for extracting out relevant Response Fields and
    # potentially writing them to the tracker file
    order = ccwparser.CCWOrderParser(order_details)

    # Get line items from order
    orderDetails = order.getOrderDetail()

    # Iterate through order line items, find the specific line item
    line_item = None
    for detail in orderDetails:
        # Sanity check pieces are here
        if detail['sku'] and detail['shipSetNumber']:
            # Match correct sku and ship set in line orders, found line item!
            if detail['sku'] == sku and detail['shipSetNumber'] == str(ship_set):
                line_item = detail
                break

    # Extract values using tracked structure keys
    for item in config.FIELDS_TO_TRACK:
        # Check if field is a class variable
        if hasattr(order, item):
            excel_line[item] = getattr(order, item)

        # Check if we found a line item
        if line_item:
            # Check if field is in line item
            if item in line_item:
                excel_line[item] = line_item[item]
        else:
            excel_line[item] = 'SKU/SS Not Found'

    return excel_line


def append_df_to_results(result_df: pd.DataFrame, parse_output: dict, keep_his_col: str | None) -> pd.DataFrame:
    """
    Append the processed line output to the result DataFrame, result DF ultimately written back to Excel sheet
    :param result_df: DataFrame to append to
    :param parse_output: Output from the processing of the line item
    :param keep_his_col: Column name we are keeping history of (date stamped)
    :return: Updated DataFrame
    """
    console.print(
        f"Processed [blue]Order[/]: [yellow]{parse_output['order_number']}[/] - [blue]SKU[/]: [yellow]{parse_output['sku']}[/]")

    new_row = {}
    for key, col_name in config.FIELDS_TO_TRACK.items():
        # If keep history enabled, swap default name for keep history column
        if keep_his_col and key == config.KEEP_HISTORY_FIELD:
            col_name = keep_his_col

        if key == 'deliveryDate' and parse_output.get('deliveryDate', 'No Data') not in ['SKU/SS Not Found',
                                                                                         'Order Not Found',
                                                                                         'Missing Data', 'No Data']:
            # Special Parsing for Delivery Date (make it more readable)
            try:
                d = datetime.fromisoformat(parse_output['deliveryDate'][:-1]).astimezone(timezone.utc)
                new_row[col_name] = d.strftime('%m/%d/%Y')
            except ValueError:
                new_row[col_name] = 'Invalid Date'
        else:
            new_row[col_name] = parse_output.get(key, 'No Data')

    # Print different outputs if data not found for a variety of reasons
    final_values = list(new_row.values())
    if 'SKU/SS Not Found' in final_values:
        console.print(f"- [red]SKU/SS Not Found in Order![/]")
    elif 'Order Not Found' in final_values:
        console.print(f"- [red]Order Not Found![/]")
    else:
        console.print(f"- Found the following values: {new_row}")

    result_df = pd.concat([result_df, pd.DataFrame([new_row])], ignore_index=True)

    return result_df


def main():
    """
    Main method to process the Excel file(s), extract order(s) details, and write back to the Excel file
    """
    console.print(Panel.fit("CCW Order Tracker"))

    console.print(Panel.fit("GET CCW Access Token", title="Step 1"))

    # Get Access Token for CCW API Requests
    access_token = get_access_token(CLIENT_KEY, CLIENT_SECRET)
    console.print("[green]Obtained Access Token for CCW API[/]")

    console.print(Panel.fit("Process Orders from Excel File(s)", title="Step 2"))

    # Read all input Excel sheets into list of Pandas Dataframe
    all_sheets = pd.read_excel(config.EXCEL_FILE_NAME, sheet_name=None, na_filter=False)

    # Initialize columns list
    columns = list(config.FIELDS_TO_TRACK.values())

    # Check if KEEP_HISTORY is true and deliveryDate is in FIELDS_TO_TRACK
    keep_his_col = None
    if config.KEEP_HISTORY and config.KEEP_HISTORY_FIELD in config.FIELDS_TO_TRACK:
        # Create a new column name with the date stamp, replace the column with the new date stamped column name
        keep_his_col = f"{config.FIELDS_TO_TRACK[config.KEEP_HISTORY_FIELD]}: {date.today().strftime('%m.%d.%Y')}"
        columns = [keep_his_col if col == config.FIELDS_TO_TRACK[config.KEEP_HISTORY_FIELD] else col for col in columns]

    # Result DataFrame with updated columns
    result_df = pd.DataFrame(columns=columns)

    if config.SINGLE_SHEET:
        # Multiple Orders on the first sheet!
        first_sheet_name = list(all_sheets.keys())[0]
        df = all_sheets[first_sheet_name]

        # Check if bare minimum columns are present, kill processing of single sheet
        if (config.ORDER_COLUMN_NAME not in df.columns) or (config.SHIP_SET_COLUMN_NAME not in df.columns) or (
                config.SKU_COLUMN_NAME not in df.columns):
            console.print('[red]Error: One or more mandatory columns are not present in the Excel Sheet, '
                          'check the column values in `config.py`[/]')
            sys.exit(-1)

        # Process up to NUM_OF_WORKERS orders simultaneously
        with Pool(NUM_OF_WORKERS) as pool:
            # Replace with custom parser depending on desired field
            results = [pool.apply_async(process_line_item, [idx, row, row[config.ORDER_COLUMN_NAME], access_token]) for
                       idx, row in df.iterrows()]

            # Iterate through results, process returned order details from processing
            for result in results:
                # Get results of parsing from custom parser method
                output = result.get()

                # Process parsing results using customer processor method, write to result DF
                result_df = append_df_to_results(result_df, output, keep_his_col)

        # Concatenate or update the existing and result DataFrame along columns
        for col in result_df.columns:
            # Handle Tracking Case
            if keep_his_col and col == keep_his_col:
                # Check if this is not the first addition of columns to Excel sheet (3 mandatory + additional columns)
                # Skip if tracking column already exists (same day)
                if keep_his_col not in df.columns and len(df.columns) >= 3 + len(config.FIELDS_TO_TRACK):
                    # Insert the column after fixed columns (push any other columns to the right)
                    df.insert(3 + len(config.FIELDS_TO_TRACK) - 1, col, result_df[col])
                    continue

            df[col] = result_df[col]

        # Save the updated DataFrame to the Excel file
        with pd.ExcelWriter(config.EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=first_sheet_name, index=False)

    else:
        # Multiple Sheets, take Order Number from sheet name
        for sheet_name, df in all_sheets.items():
            # Check if bare minimum columns are present
            if (config.SHIP_SET_COLUMN_NAME not in df.columns) or (config.SKU_COLUMN_NAME not in df.columns):
                console.print(
                    f'[red]Error: One or more mandatory columns are not present in the Excel Sheet[/]: {sheet_name}')
                continue

            # Process up to NUM_OF_WORKERS orders simultaneously
            with Pool(NUM_OF_WORKERS) as pool:
                # Replace with custom parser depending on desired field
                results = [pool.apply_async(process_line_item, [idx, row, sheet_name, access_token])
                           for idx, row in df.iterrows()]

                # Iterate through results, process returned order details from processing
                for result in results:
                    # Get results of parsing from custom parser method
                    output = result.get()

                    # Process parsing results using customer processor method, write to result DF
                    result_df = append_df_to_results(result_df, output, keep_his_col)

            # Concatenate or update the existing and result DataFrame along columns
            for col in result_df.columns:
                # Handle Tracking Case
                if keep_his_col and col == keep_his_col:
                    # Check if this is not the first addition of columns to Excel sheet (2 mandatory + additional columns)
                    # Skip if tracking column already exists (same day)
                    if keep_his_col not in df.columns and len(df.columns) >= 2 + len(config.FIELDS_TO_TRACK):
                        # Insert the column after fixed columns (push any other columns to the right)
                        df.insert(2 + len(config.FIELDS_TO_TRACK) - 1, col, result_df[col])
                        continue

                df[col] = result_df[col]

            # Save the updated DataFrame to the Excel file
            with pd.ExcelWriter(config.EXCEL_FILE_NAME, engine="openpyxl", mode='a',
                                if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Empty Values from DF
            result_df = result_df[0:0]


if __name__ == '__main__':
    main()
