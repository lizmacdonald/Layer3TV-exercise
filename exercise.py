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
## Need to find the min date so week 1 is the first week not the last week
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

## Need to add a 1 to the wk number so the week doesn't start at zero
for i in dif:
	wknum.append(i / 7)


df['dif'] = wknum
df['dif'] = 'Week ' + df['dif'].astype(str)


## Change the type of the columns to numeric
for col in ['new_subscriptions', 'self_install', 'professional_install', 'disconnects', 'post_install_returns']:
    df[col] = df[col].astype(int)


## Set up net gain and net loss columns
df['Total Disconnects'] = df['post_install_returns'] + df['disconnects']

df['Net Gain'] = df['new_subscriptions'] - df['Total Disconnects']


## Set up cumulative sum of subscriptions
sumtot = df['Net Gain'].cumsum()

tot = []
tot2 = []

for i in sumtot:
	tot.append(i)

for i in reversed(tot):
	tot2.append(i)

df['Total Subscriptions'] = tot2

## Update column names
df.columns = ['activity_date', 'market', 'Total Connects', 'tot_discon', 'Self Installs', 'Pro Installs', 'Disconnects', 'Post Install Returns', 'tot_sub', 'dif', 'Total Disconnects', 'Net Gain']


## Create dataframe for aggregate data 
df_ag = df.groupby(['dif']).sum().reset_index()




## Set up column names to be the format I want
df.columns = ['activity_date', 'market', 'Total Connects', 'Total Disconnects', 'Self Installs', 'Pro Installs', 'Disconnects', 'Post Install Returns', 'Total Subscribers', 'dif']




dfout = np.transpose(df)

