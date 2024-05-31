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
from pprint import pprint


def get_nested_value(data: list | dict, keys: list, default="No Data"):
    """
    This function extracts a value nested within a list or dictionary, or if not found, returns "No Data"
    :param data: list or dictionary
    :param keys: list of keys to traverse
    :param default: default value to return if not found
    :return: value or default
    """
    for key in keys:
        if isinstance(data, list) and key < len(data):
            data = data[key]
        elif isinstance(data, dict) and key in data:
            data = data[key]
        else:
            return default
    return data


class CCWOrderParser(object):
    """
    This class is used to parse the raw data from the CCW API into a more usable format.
    """

    def __init__(self, variable):
        """
        This init method is called whenever a new object is instantiated. This will take the raw text
        from CCW Output and parse the appropriate data from the records into the appropriate fields.
        :param variable: Raw text output from CCW API
        """

        self.rawdata = variable
        self.jsondata = json.loads(variable)

        headerlocation = self.jsondata['ShowPurchaseOrder']['value']['DataArea']['PurchaseOrder'][0][
            'PurchaseOrderHeader']

        self.billtoparty = get_nested_value(headerlocation, ['BillToParty', 'Name', 0, 'value'])
        self.party = get_nested_value(headerlocation, ['Party', 0, 'Name', 0, 'value'])
        self.status = get_nested_value(headerlocation, ['Status', 0, 'Description', 'value'])
        self.salesordernum = get_nested_value(headerlocation, ['SalesOrderReference', 0, 'ID', 'value'])
        self.shiptoparty = get_nested_value(headerlocation, ['ShipToParty', 'Name', 0, 'value'])
        self.amount = get_nested_value(headerlocation, ['TotalAmount', 'value'])
        self.currencycode = get_nested_value(headerlocation, ['TotalAmount', 'currencyCode'])

        self.linelocation = self.jsondata['ShowPurchaseOrder']['value']['DataArea']['PurchaseOrder'][0][
            'PurchaseOrderLine']

        self.orderdetail = []
        for i in self.linelocation:
            line = {
                "sku": get_nested_value(i, ['Item', 'ID', 'value']),
                "description": get_nested_value(i, ['Item', 'Description', 0, 'value']),
                "quantity": get_nested_value(i, ['Item', 'Lot', 0, 'Quantity', 'value']),
                "line": get_nested_value(i, ['SalesOrderReference', 'LineNumberID', 'value']),
                "amount": get_nested_value(i, ['ExtendedAmount', 'value']),
                "deliveryDate": get_nested_value(i, ['PromisedDeliveryDateTime']),
                "shipSetNumber": get_nested_value(i, ['LineIDSet', 0, 'ID', 0, 'value'])
            }
            self.orderdetail.append(line)

    def getDisplay(self):
        """
        display - This method will display the raw json in pretty print format
        :return: nothing
        """
        pprint(self.jsondata)

    def getStatus(self):
        """
        status - This method provides access to the status
        :return: status (String)
        """
        return self.status

    def getBillToParty(self):
        """
        billtoparty - This method provides access to the bill-to-party
        :return: billtoparty (String)
        """
        return self.billtoparty

    def getParty(self):
        """
        party - This method provides access to the party (Seems to be the end customer)
        :return: party (String)
        """
        return self.party

    def getSalesOrderNum(self):
        """
        salesordernum - This method provides access to the sales order number (SO#)
        :return: salesordernum (String)
        """
        return self.salesordernum

    def getShipToParty(self):
        """
        shiptoparty - This method provides access to the ship-to-party (Where the equipment is shipped)
        :return: shiptoparty (String)
        """
        return self.shiptoparty

    def getAmount(self):
        """
        amount - This method provides access to the total amount of the order
        :return: amount (Float)
        """
        return self.amount

    def getCurrencyCode(self):
        """
        currencycode - This method provides access to the currency code of the oder
        :return: currencycode (String)
        """
        return self.currencycode

    def getLineLocation(self):
        return self.linelocation

    def getOrderDetail(self):
        """
        orderdetail - This method will return the dictionary that represents the details of the order.
        :return: orderdetail (Dictionary)
        """
        return self.orderdetail
