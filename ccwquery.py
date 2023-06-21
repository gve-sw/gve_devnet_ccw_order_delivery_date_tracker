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

import sys
from datetime import datetime, date, timezone
from openpyxl.utils.cell import column_index_from_string
from openpyxl import load_workbook

from multiprocessing import Pool

import pandas as pd
import requests
from rich.console import Console
from rich.panel import Panel

import ccwparser
import config

# Rich Console Instance
console = Console()


def getAccessToken(client_id, client_secret):
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


def getOrderDetails(access_token, id):
    """
    Get Order details from CCW API for a particular Sales Order Number, returns a number of fields which are parsed using CCWOrderParser class
    :param access_token: CCW Order API Token
    :param id: Sales Order Number
    :return:
    """
    payload = {
        "GetPurchaseOrder": {
            "value": {
                "DataArea": {
                    "PurchaseOrder": [
                        {
                            "PurchaseOrderHeader": {
                                "Description": [
                                    {
                                        "value": "YES",
                                        "typeCode": "details"
                                    }
                                ],
                                "SalesOrderReference": [
                                    {
                                        "ID": {
                                            "value": id
                                        }
                                    }
                                ]
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
        return


def find_delivery_date(idx, row, access_token):
    """
    Find delivery date for specific SKU in the order
    :param idx: Row Index to write result too
    :param row: Excel row, used to extract fields for processing
    :param access_token: CCW Access Token, necessary for API Calls
    :return: Delivery Date for SKU
    """
    # Get the input values from the Excel columns
    sales_order_number = row[config.SALES_ORDER_NUMBER_COLUMN_NAME]
    ship_set = str(row[config.SHIP_SET_COLUMN_NAME]).strip()
    sku = str(row[config.SKU_COLUMN_NAME]).strip()

    # Skip rows with missing data
    if sales_order_number == '' or sku == '' or ship_set == "":
        return {"idx": idx, "sales_order_number": sales_order_number, "sku": sku, "delivery_date": 'Missing Data'}

    # Get Order Details from CCW using the Sales Order Number
    order_details = getOrderDetails(access_token, sales_order_number)

    # Order details are none, order = not found or unauthorized access
    if not order_details:
        return {"idx": idx, "sales_order_number": sales_order_number, "sku": sku, "delivery_date": 'Order Not Found'}

    # Build Order Parser Object, includes a number of methods for extracting out relevant Response Fields and
    # potentially writing them to the tracker file
    order = ccwparser.CCWOrderParser(order_details)

    # Get line items from order
    orderDetails = order.getOrderDetail()

    # Iterate through order line items, extract delivery date corresponding to SKU
    for details in orderDetails:
        # Sanity check pieces are here
        if details['sku'] and details['shipSetNumber'] and details['deliveryDate']:
            # Match correct sku and ship set in line orders
            if details['sku'] == sku and details['shipSetNumber'] == str(ship_set):
                return {"idx": idx, "sales_order_number": sales_order_number, "sku": sku, "delivery_date": details['deliveryDate']}

    # If no deliveryDate found
    return {"idx": idx, "sales_order_number": sales_order_number, "sku": sku, "delivery_date": "SKU/SS Not Found"}


def process_delivery_date(result_df, parse_output):
    """
    Process each delivery date, write results to new Dataframe
    :param result_df: Results Dataframe (written to excel)
    :param parse_output: Output of parsing this order
    :return: Results Dataframe with new row entry
    """
    new_row = {result_df.columns[0]: ''}

    if parse_output['delivery_date'] == 'Missing Data':
        # Skip blank rows/blank entries
        new_row[result_df.columns[0]] = ''
    else:
        console.print(f"Processing Sales Order: [blue]'{int(parse_output['sales_order_number'])}'[/]")

        if parse_output['delivery_date'] == 'SKU/SS Not Found' or parse_output['delivery_date'] == 'Order Not Found':
            console.print(f' - [red]No delivery date for target sku[/] ({parse_output["sku"]})... {parse_output["delivery_date"]}')
        else:
            # Convert Delivery Date to more readable format
            d = datetime.fromisoformat(parse_output['delivery_date'][:-1]).astimezone(timezone.utc)
            parse_output["delivery_date"] = d.strftime('%m/%d/%Y')

            console.print(f' - [green]Found a delivery date for target sku[/] ({parse_output["sku"]}): {parse_output["delivery_date"]}')

        # Update the 'output_column' row with the result
        new_row[result_df.columns[0]] = parse_output["delivery_date"]

    # Combine new row with output Dataframe
    result_df = pd.concat([result_df, pd.DataFrame([new_row])], ignore_index=True)

    return result_df


def main():
    console.print(Panel.fit("CCW Order Delivery Date Tracker"))

    console.print(Panel.fit("GET CCW Access Token", title="Step 1"))

    # Get Access Token for CCW API Requests
    access_token = getAccessToken(config.CLIENT_KEY, config.CLIENT_SECRET)
    console.print("[green]Obtained Access Token for CCW API[/]")

    console.print(Panel.fit("Process Orders from Excel File", title="Step 2"))

    # Read in input Excel into Pandas Dataframe
    df = pd.read_excel(config.EXCEL_FILE_NAME, sheet_name=config.EXCEL_SHEET_NAME, na_filter=False)

    # Result Dataframe (if keep_history flag is true, then create a new DF with Output Column - Date Stamped)
    if config.KEEP_HISTORY:
        new_col = f"{config.OUTPUT_COLUMN_NAME}: {date.today().strftime('%m.%d.%Y')}"

        # Skip adding processing if column already exists for the day
        if new_col in df.columns:
            console.print('[green]Column already exists for today! No need to run again.[/]')
            sys.exit(0)

        new_df = pd.DataFrame(columns=[new_col])
    else:
        # Check if bare minimum columns are present
        if (config.SALES_ORDER_NUMBER_COLUMN_NAME not in df.columns) or (config.OUTPUT_COLUMN_NAME not in df.columns):
            console.print('[red]Error: Sales Order column and/or Output Column are not present in the Excel File, '
                          'check the column values in `config.py`')
            sys.exit(-1)

        new_df = pd.DataFrame(columns=[config.OUTPUT_COLUMN_NAME])

    NUM_OF_WORKERS = 10

    # Process up to NUM_OF_WORKERS orders simultaneously
    with Pool(NUM_OF_WORKERS) as pool:
        # Replace with custom parser depending on desired field
        results = [pool.apply_async(find_delivery_date, [idx, row, access_token]) for idx, row in df.iterrows()]

        # Iterate through results, process returned order details from processing
        for result in results:
            # Get results of parsing from custom parser method
            output = result.get()

            # Process parsing results using customer processor method, write to new DF
            new_df = process_delivery_date(new_df, output)

    # If we want the history of the column, append a blank new column before the insertion point (set in config) -
    # columns tracked from newest -> oldest, left to right
    if config.KEEP_HISTORY:
        # Get index of column letter
        column_index = column_index_from_string(config.INSERT_BEFORE_COLUMN)

        # Load workbook and sheet
        wb = load_workbook(config.EXCEL_FILE_NAME)
        sheet = wb[config.EXCEL_SHEET_NAME]

        # Insert new blank column, shifts columns to the right
        sheet.insert_cols(column_index)

        # Save workbook
        wb.save(config.EXCEL_FILE_NAME)

        # Set the column index to the previous column for writing data (due to 0 index in df.to_excel)
        column_index = column_index - 1
    else:
        # Set column index to value of target column
        column_index = df.columns.get_loc(config.OUTPUT_COLUMN_NAME)

    # Save the updated DataFrame to a new Excel file
    with pd.ExcelWriter(config.EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        new_df.to_excel(writer, sheet_name=config.EXCEL_SHEET_NAME, startcol=column_index, index=False)


if __name__ == '__main__':
    main()
