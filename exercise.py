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
df['dif'] = df['dif'].astype(int)
df['dif'] = 1 + df['dif']
df['dif'] = 'Week ' + df['dif'].astype(str)


## Change the type of the columns to numeric
for col in ['new_subscriptions', 'self_install', 'professional_install', 'disconnects', 'post_install_returns']:
    df[col] = df[col].astype(int)


## Set up net gain and net loss columns
df['Total Disconnects'] = df['post_install_returns'] + df['disconnects']

df['Net Gain'] = df['new_subscriptions'] - df['Total Disconnects']


## Set up cumulative sum of subscriptions
dc = df
dc = df.groupby(['dif']).sum().reset_index()

daily = []
daily2 = []

for i in dc['Net Gain']:
	daily.append(i)

for i in reversed(daily):
	daily2.append(i)

dc['dayR'] = daily2
dc['culday'] = dc['dayR'].cumsum()

tot = []
tot2 = []

for i in dc['culday']:
	tot.append(i)

for i in reversed(tot):
	tot2.append(i)

dc['Ending Subs'] = tot2

dc['Beginning Subscribers'] = dc['Ending Subs'] - dc['Net Gain']

dc_col = ['dif', 'Ending Subs', 'Beginning Subscribers']
dc = dc[dc_col]

df = pd.merge(df, dc, on = 'dif')
df = df.append(dc)

## Select desired columns Update column names

col = ['dif', 'market', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

df = df[col]

df.columns = ['dif', 'market', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## Create dataframe for aggregate data 
df_ag = df.groupby(['dif']).sum().reset_index()
df_ag = np.transpose(df_ag)
df_ag = df_ag.reset_index()
df_ag = df_ag.set_value(0, 'index', 'Aggregate')
df_ag = df_ag.drop([1, ])
x = pd.Series([""], index = df_ag.index)
df_ag = df_ag.append(x, ignore_index=True)

## Create Dataframe fro Atlanta data
df_atl = df[df['market'] == 'Atlanta']
df_atl = df_atl.groupby(['dif', 'market']).sum().reset_index()
df_atl = pd.DataFrame(np.transpose(df_atl))
df_atl = df_atl.reset_index()
df_atl = df_atl.set_value(0, 'index', 'Atlanta')
df_atl = df_atl.drop([1, ])
x = pd.Series([" "], index = df_atl.index)
df_atl = df_atl.append(x, ignore_index=True)


## Create Dataframe for Seattle data
df_sea = df[df['market'] == 'Seattle']
df_sea = df_sea.groupby(['dif', 'market']).sum().reset_index()
df_sea = pd.DataFrame(np.transpose(df_sea))
df_sea = df_sea.reset_index()
df_sea = df_sea.set_value(0, 'index', 'Seattle')
df_sea = df_sea.drop([1, ])

## Bind datasets together and remove NaNs
frame = [df_ag, df_atl, df_sea]
df_f = pd.concat(frame)

df_f = df_f.fillna("")






