#!/usr/bin/env python2

#############################################################
# Layer3 Exercise
#############################################################

import openpyxl as px
import pandas as pd
import datetime
import numpy as np
import transposer

## read in data
dat = px.load_workbook('subscriber-data.xlsx', data_only=True)
ws = dat['data']

df = pd.DataFrame(ws.values)
df.columns = [df.iloc[0]]
df = df.drop([0, ])


## Set up week number column
### need to find the min date so week 1 is the first week not the last week
dat_form = "%Y-%m-%d %H:%S:%M"
max_day = pd.to_datetime(df.iloc[1,0], format = dat_form)
days = df.iloc[:,0]

wk = []
dif = []
wknum = []

for d in days:
	wk.append(pd.to_datetime(d, format = dat_form))

for w in wk:
	dif.append(int((max_day - w).days))

## need to add a 1 to the wk number so the week doesn't start at zero
for i in dif:
	wknum.append(i / 7)


df['dif'] = wknum
df['dif'] = 'Week ' + df['dif'].astype(str)


## Set up column names to be the format I want
df.columns = ['activity_date', 'market', 'Total Connects', 'Total Disconnects', 'Self Installs', 'Pro Installs', 'Disconnects', 'Post Install Returns', 'Total Subscribers', 'dif']

df['Net Gain'] = df['Total Connects'].astype(int) - df['Total Disconnects'].astype(int)

df['Beginning Sub']



dfout = pd.transpose(df)

