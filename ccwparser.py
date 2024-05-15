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


class CCWOrderParser(object):

    def __init__(self, variable):
        """
        init - This init method is called whenever a new object is instantiated.   This will take the the raw text
               from CCW Output and parse the appropriate data from the records into the appropriate fields.
        :return: nothing
        """

        self.rawdata = variable
        self.jsondata = json.loads(variable)

        headerlocation = self.jsondata['ShowPurchaseOrder']['value']['DataArea']['PurchaseOrder'][0][
            'PurchaseOrderHeader']

        self.billtoparty = headerlocation['BillToParty']['Name'][0]['value']
        self.party = headerlocation['Party'][0]['Name'][0]['value']
        self.status = headerlocation['Status'][0]['Description']['value']
        self.salesordernum = headerlocation['SalesOrderReference'][0]['ID']['value']
        self.shiptoparty = headerlocation['ShipToParty']['Name'][0]['value']

        # self.amount = headerlocation['TotalAmount']['value']
        # self.currencycode = headerlocation['TotalAmount']['currencyCode']

        self.linelocation = self.jsondata['ShowPurchaseOrder']['value']['DataArea']['PurchaseOrder'][0][
            'PurchaseOrderLine']

        self.orderdetail = []
        for i in self.linelocation:
            line = {"sku": i['Item']['ID']['value'],
                    "description": i['Item']['Description'][0]['value'],
                    "quantity": i['Item']['Lot'][0]['Quantity']['value'],
                    "line": i['SalesOrderReference']['LineNumberID']['value'],
                    # "amount": i['ExtendedAmount']['value'],
                    "deliveryDate": i['PromisedDeliveryDateTime'] if 'PromisedDeliveryDateTime' in i else None,
                    "shipSetNumber": i['LineIDSet'][0]['ID'][0]['value'] if 'LineIDSet' in i else None
                    }
            self.orderdetail.append(line)

    def getDisplay(self):
        """
        dispay - This method will display the raw json in pretty print format
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
