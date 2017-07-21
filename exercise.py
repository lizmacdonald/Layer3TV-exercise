#!/usr/bin/env python2

#############################################################
# Layer3 Exercise
#############################################################

import openpyxl as px
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import datetime


## read in data
dat = px.load_workbook('subscriber-data.xlsx', data_only=True)
ws = dat['data']

df = pd.DataFrame(ws.values)
df.columns = [df.iloc[0]]
df = df.drop([0, ])


## Set up week number column
dat_form = "%Y-%m-%d %H:%S:%M"
max_day = pd.to_datetime(df.iloc[1,0], format = dat_form)
days = df.iloc[:,0]

week = []
dif = []
wknum = []

for d in days:
	week.append(pd.to_datetime(d, format = dat_form))

for w in week:
	dif.append(int((max_day - w).days))

## Need to add a 1 to the wk number so the week doesn't start at zero
for i in dif:
	wknum.append(i / 7)

df['Week'] = wknum
df['Week'] = df['Week'].astype(int)
df['Week'] = 1 + df['Week']


## Set up monthy number columns
df['Month'] = df['activity_date'].apply(lambda x: x.strftime('%m-%Y'))

## Set up quarterly number column
quart = pd.to_datetime(df['activity_date']).dt.quarter

year = []

for d in days:
	year.append(d.year)

df['Quarter'] = year
df['Quarter'] = df['Quarter'].astype(str) + ' ' + 'Q' + quart.astype(str)


## Change the type of the columns to numeric
for col in ['new_subscriptions', 'self_install', 'professional_install', 'disconnects', 'post_install_returns']:
    df[col] = df[col].astype(int)


## Set up net gain and net loss columns
df['Total Disconnects'] = df['post_install_returns'] + df['disconnects']

df['Net Gain'] = df['new_subscriptions'] - df['Total Disconnects']


## Create a list for adding in the word "week" before the week number

df['Week'] = df['Week'].astype(str)
wk_unique = df.Week.unique()
wk_word = []

for i in wk_unique:
	wk_word.append('Week ' + str(i))

wk_word = np.array(wk_word)


## Set up separate dataframes for each market
## Set up column names to be used for each data subset
wkcol = ['Week', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

wk_fin = ['Week', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']

moncol = ['Month', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

mon_fin = ['Month', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']

qucol = ['Quarter', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

qu_fin = ['Quarter', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## AGGREGATE: subset for the weekly data
dc = df.groupby(['Week']).sum().reset_index()
dc['Week'] = dc['Week'].astype(int)
dc = dc.sort_values('Week', ascending = False)

## AGGREGATE: Set up weekly cumulatvie beg and end subscription numbers
dc['Ending Subs'] = dc['Net Gain'].cumsum()
dc['Beginning Subscribers'] = dc['Ending Subs'] - dc['Net Gain']

dc = dc[wkcol]
dc.columns = wk_fin

## AGGREGATE: transpose weekly the data
dc = dc.sort_values('Week', ascending = True)
dc = np.transpose(dc)

### vector for subsequent datasets
blankrow = pd.Series([""], index = dc.index)

## AGGREGATE: add the word 'week' to the week data, clean up indexing
dc = dc.T.reset_index(drop=True).T
dc.loc[['Week'], 0:132] = wk_word
dc = dc.reset_index()
dc = dc.set_value(0, 'index', 'Aggregate Market')
dc = dc.append(blankrow, ignore_index=True)


## AGGREGATE: select the monthly datasbusets
dcm = df.groupby(['Month']).sum().reset_index()
dcm['Month'] = pd.to_datetime(dcm['Month'])
dcm = dcm.sort_values('Month', ascending = True)
dcm['Month'] = dcm['Month'].apply(lambda x: x.strftime('%b-%Y'))
dcm = dcm.reset_index(drop = True)

## AGGREGATE: Set up monthly cumulatvie beg and end subscription numbers
dcm['Ending Subs'] = dcm['Net Gain'].cumsum()
dcm['Beginning Subscribers'] = dcm['Ending Subs'] - dcm['Net Gain']

dcm = dcm[moncol]
dcm.columns = mon_fin


## AGGREGATE: transpose the monthly data
dcm['Month'] = pd.to_datetime(dcm['Month'])
dcm = dcm.sort_values('Month', ascending = False)
dcm['Month'] = dcm['Month'].apply(lambda x: x.strftime('%b-%Y'))
dcm = np.transpose(dcm)

## AGGREGATE: clean up indexing
dcm = dcm.reset_index()
dcm = dcm.set_value(0, 'index', 'Aggregate Market')
dcm = dcm.append(blankrow, ignore_index=True)

## AGGREGATE: select the quarterly datasbusets
dcq = df.groupby(['Quarter']).sum().reset_index()
dcq = dcq.sort_values('Quarter', ascending = True)

## AGGREGATE: Set up quarterly cumulatvie beg and end subscription numbers
dcq['Ending Subs'] = dcq['Net Gain'].cumsum()
dcq['Beginning Subscribers'] = dcq['Ending Subs'] - dcq['Net Gain']

dcq = dcq[qucol]
dcq.columns = qu_fin

## AGGREGATE: transpose the quarterly data
dcq = dcq.sort_values('Quarter', ascending = False)
dcq = np.transpose(dcq)

## AGGREGATE: clean up quarterly indexing
dcq = dcq.reset_index()
dcq = dcq.set_value(0, 'index', 'Aggregate Market')
dcq = dcq.append(blankrow, ignore_index=True)



## ATLANTA: select datasubset
da = df[df['market'] == 'Atlanta']

## ATLANTA: subset for the weekly data
daw = da.groupby(['Week']).sum().reset_index()
daw['Week'] = daw['Week'].astype(int)
daw = daw.sort_values('Week', ascending = False)


## ATLANTA: Set up weekly cumulatvie beg and end subscription numbers
daw['Ending Subs'] = daw['Net Gain'].cumsum()
daw['Beginning Subscribers'] = daw['Ending Subs'] - daw['Net Gain']

daw = daw[wkcol]
daw.columns = wk_fin

## ATLANTA: transpose weekly the data
daw = daw.sort_values('Week', ascending = True)
daw = np.transpose(daw)


## ATLANTA: add the word 'week' to the week data, clean up indexing
daw = daw.T.reset_index(drop=True).T
daw.loc[['Week'], 0:132] = wk_word
daw['index'] = daw.index
daw = daw.reset_index(drop = True)
cols = daw.columns.tolist()
cols = [cols[-1]] + cols[ : -1]
daw = daw.reindex(columns = cols)
daw = daw.set_value(0, 'index', 'Atlanta Market')
blankrow = pd.Series([""], index = daw.index)
daw = daw.append(blankrow, ignore_index=True)



## ATLANTA: select the monthly datasbusets
dam = da.groupby(['Month']).sum().reset_index()
dam['Month'] = pd.to_datetime(dam['Month'])
dam = dam.sort_values('Month', ascending = True)
dam['Month'] = dam['Month'].apply(lambda x: x.strftime('%b-%Y'))
dam = dam.reset_index(drop = True)

## ATLANTA: Set up monthly cumulatvie beg and end subscription numbers
dam['Ending Subs'] = dam['Net Gain'].cumsum()
dam['Beginning Subscribers'] = dam['Ending Subs'] - dam['Net Gain']

dam = dam[moncol]
dam.columns = mon_fin


## ATLANTA: transpose the monthly data
dam['Month'] = pd.to_datetime(dam['Month'])
dam = dam.sort_values('Month', ascending = False)
dam['Month'] = dam['Month'].apply(lambda x: x.strftime('%b-%Y'))
dam = np.transpose(dam)

## ATLANTA: clean up monthly indexing
dam = dam.reset_index()
dam = dam.set_value(0, 'index', 'Aggregate Market')
blankrow = pd.Series([""], index = dam.index)
dam = dam.append(blankrow, ignore_index=True)



## AGGREGATE: select the quarterly datasbusets
daq = da.groupby(['Quarter']).sum().reset_index()
daq = daq.sort_values('Quarter', ascending = True)

## AGGREGATE: Set up quarterly cumulatvie beg and end subscription numbers
daq['Ending Subs'] = daq['Net Gain'].cumsum()
daq['Beginning Subscribers'] = daq['Ending Subs'] - daq['Net Gain']

daq = daq[qucol]
daq.columns = qu_fin

## AGGREGATE: transpose the quarterly data
daq = daq.sort_values('Quarter', ascending = False)
daq = np.transpose(daq)

## AGGREGATE: clean up quarterly indexing
daq = daq.reset_index()
daq = daq.set_value(0, 'index', 'Aggregate Market')
blankrow = pd.Series([""], index = daq.index)
daq = daq.append(blankrow, ignore_index=True)



## SEATTLE: Select data subset
ds = df[df['market'] == 'Seattle']
ds['Week'] = ds['Week'].astype(int)
ds = ds.groupby(['Week']).sum().reset_index()


## SEATTLE: Calculate the atlanta cumulative numbers
daily_se = []
daily_se2 = []
tot_se = []
tot_se2 = []

for i in ds['Net Gain']:
	daily_se.append(i)

for i in reversed(daily_se):
	daily_se2.append(i)

ds['Se_cul'] = daily_se2
ds['Se_cul'] = ds['Se_cul'].cumsum()

for i in ds['Se_cul']:
	tot_se.append(i)

for i in reversed(tot_se):
	tot_se2.append(i)

ds['Ending Subs'] = tot_se2
ds['Beginning Subscribers'] = ds['Ending Subs'] - ds['Net Gain']


## SEATTLE: Select desired columns Update column names
ds = ds[col]

ds.columns = ['Week', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## SEATTLE: transpose add the word 'week' to the week data, clean up indexing
df_sea = np.transpose(ds)
df_sea.loc[['Week'], 0:132] = wk_word

## SEATTLE: add index names to first column of dataset
df_sea['index'] = df_sea.index
df_sea = df_sea.reset_index(drop = True)
df_sea = df_sea.reindex(columns = cols)

df_sea = df_sea.set_value(0, 'index', 'Seattle Market')



## BIND datasets together and remove NaNs
frame = [df_ag, df_atl, df_sea]
df_f = pd.concat(frame)


## BIND: add header row to describe report
title = ' Subscriber Report Week-Over-Week 2017'
d = pd.DataFrame(columns = df_f.columns)
d = d.set_value(len(d), 'index', " ")
df_f.index = df_f.index + 1
df_f = d.append(df_f)
df_f = df_f.set_value(0, 'index', title)




df_f = df_f.fillna("")


## Write out the report
writer = pd.ExcelWriter('report.xlsx', engine = 'xlsxwriter')

df_f.to_excel(writer, sheet_name = 'Weekly Data Report', index = False, header = False)

workbook = writer.book

worksheet = writer.sheets['Weekly Data Report']
worksheet.set_column('A:A', 18)


writer.save()

