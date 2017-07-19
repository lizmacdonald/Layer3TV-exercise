#!/usr/bin/env python2

#############################################################
# Layer3 Exercise
#############################################################

import openpyxl as px
import pandas as pd
import datetime


dat = px.load_workbook('subscriber-data.xlsx', data_only=True)
ws = dat['data']

df = pd.DataFrame(ws.values)
df.columns = [df.iloc[0]]
df = df.drop([0, ])

dat_form = "%Y-%m-%d %H:%S:%M"

## Set up week number column

max_day = pd.to_datetime(df.iloc[1,0], format = dat_form)
days = df.iloc[:,0]

wk = []

for day in days:
	wk.append(pd.to_datetime(day, format = dat_form))

dif = []

for day in wk:
	dif.append(int((max_day - day).days))

wknum = []

for i in dif:
	wknum.append(i / 7)

df['dif'] = wknum


