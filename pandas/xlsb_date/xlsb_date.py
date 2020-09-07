# -*- coding: utf-8 -*-
"""
Created on Mon Sep  7 11:36:25 2020

@author: lsrpequeno@gmail.com

For *.xlsb Excel files, date comes as a numerical data (float).
This script find a way to solve this issue and parse the numeric value to a date type

reference: https://stackoverflow.com/questions/31359150/convert-date-from-excel-in-number-format-to-date-format-python#:~:text=The%20module%20xlrd%20provides%20a,tuple%20into%20a%20datetime%20%2Dobject.

For dates before 1904 we have an issue due the starting year (1900) which is a Leap year,  
the date 1900-02-29 is not considered, resulting in a mistaken calculation

"""

import pandas as pd 
from datetime import datetime

# xlsb file
excel_file =  pd.read_excel('xlsb_date.xlsb', header=0, index_col=False, 
                            engine='pyxlsb', sheet_name='sheet1', 
                            skiprows=(0), names=['DATE_EXAMPLE'], usecols=('A'))

# this is to store treated data
fixed_date = []

# lets run the entire df(excel_file)
for ind in excel_file.index:
        fixed_date.append(datetime.fromordinal(datetime(1900,1,1).toordinal()+ int(excel_file['DATE_EXAMPLE'][ind])-2).strftime("%Y-%m-%d"))

# add the nova_data to excel_file    
excel_file.insert(1, 'DATE_EXAMPLE_FIXED', fixed_date)
