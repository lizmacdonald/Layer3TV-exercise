#!/usr/bin/env python2

#############################################################
# Layer3 Exercise
#############################################################

import openpyxl as px
import pandas as pd

dat = px.load_workbook('subscriber-data.xlsx', data_only=True, parse_dates = [0])
ws = dat['data']

df = pd.DataFrame(ws.values)
df.columns = [df.iloc[0]]
df = df.drop([0, ])
