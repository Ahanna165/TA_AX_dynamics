# Price Group for Dynamics AX
Tool creates Trade Agreements for AX Dynamics in N number of currencies currencies

## Content Description
* creates new version directory in your folder system according to the date
* reads the source files and recognizes currency. renames the file according to the currency inside
* class TradeAgreemen cleans the data, shapes the TA for each input file and safes the TAs
* class SampleData selects items randomly (files with volume prices) than it displys the values in file and opens the URL link to that product

## Import

*    pandas
*    os
*    re
*    xlsxwriter
*    openpyxl
*    datetime, timedelta, date
*    numpy
*    webbrowser
*    from pathlib import Path

## Output Files Columns
['Price/discount journal number', 'Line number', 'Relation', 'Account code', 'Account selection', 
'Item code', 'Item relation', 'Configuration', 'Site', 'Warehouse', 'Location', 'From', 'To', 'Unit',
'Amount in currency', 'From date', 'To date', 'Currency', 'Price charges', 'Price unit', 
'Discount percentage 1', 'Discount percentage 2', 'Lead time', 'Working days', 'Disregard lead time', 
'Include generic currency', 'Find next']

## Sample for Test
* filter items having volume prices only
* sample size each 600th item will be opened as URL
