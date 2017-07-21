#!/usr/bin/env python2

#############################################################
# Layer3 Exercise
#############################################################

import openpyxl as px
import pandas as pd
import numpy as np
import xlsxwriter


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
wkdat = np.array(wk_word)


## Set up separate dataframes for each market

## Aggregate
dc = df.groupby(['Week']).sum().reset_index()

## Seattle
ds = df[df['market'] == 'Seattle']
ds['Week'] = ds['Week'].astype(int)
ds = ds.groupby(['Week']).sum().reset_index()


## AGGREGATE: Set up cumulatvie beg and end subscription numbers

daily_ag = []
daily_ag2 = []
tot_ag = []
tot_ag2 = []

for i in dc['Net Gain']:
	daily_ag.append(i)

for i in reversed(daily_ag):
	daily_ag2.append(i)

dc['Ag_cul'] = daily_ag2
dc['Ag_cul'] = dc['Ag_cul'].cumsum()

for i in dc['Ag_cul']:
	tot_ag.append(i)

for i in reversed(tot_ag):
	tot_ag2.append(i)

dc['Ending Subs'] = tot_ag2
dc['beg_sub_ag'] = dc['Ending Subs'] - dc['Net Gain']


## AGGREGATE: Select the desired columns
col = ['Week', 'market', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

dc = dc[col]
dc.columns = ['Week', 'market', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## AGGREGATE: transpose the data
df_ag = np.transpose(dc)

## AGGREGATE: add the word 'week' to the week data, clean up indexing
df_ag.loc[['Week'], 0:132] = wkdat
df_ag = df_ag.reset_index()
df_ag = df_ag.set_value(0, 'index', 'Aggregate')
x = pd.Series([""], index = df_ag.index)
df_ag = df_ag.append(x, ignore_index=True)




## ATLANTA: select data
da = df[df['market'] == 'Atlanta']
da['Week'] = da['Week'].astype(int)
da = da.groupby(['Week']).sum().reset_index()


## ATLANTA: Calculate the atlanta cumulative numbers
daily_at = []
daily_at2 = []
tot_at = []
tot_at2 = []

for i in da['Net Gain']:
	daily_at.append(i)

for i in reversed(daily_at):
	daily_at2.append(i)

da['At_cul'] = daily_at2
da['At_cul'] = da['At_cul'].cumsum()

for i in da['At_cul']:
	tot_at.append(i)

for i in reversed(tot_at):
	tot_at2.append(i)

da['end_sub_at'] = tot_at2
da['beg_sub_at'] = da['end_sub_at'] - da['Net Gain']


## Calculate the atlanta cumulative numbers
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

ds['end_sub_se'] = tot_se2
ds['beg_sub_se'] = ds['end_sub_se'] - ds['Net Gain']


## Select the beginning and end subscriber columns from each dataset
dc_col = ['Week', 'beg_sub_ag', 'end_sub_ag']
dc = dc[dc_col]

da_col = ['Week', 'beg_sub_at', 'end_sub_at']
da = da[da_col]

ds_col = ['Week', 'beg_sub_se', 'end_sub_se']
ds = ds[ds_col]

df = pd.merge(df, dc, on = 'Week')
df = pd.merge(df, da, on = "Week")
df = pd.merge(df, ds, on = "Week")

## Select desired columns Update column names

col = ['Week', 'market', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

df = df[col]

df.columns = ['Week', 'market', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## Create dataframe for aggregate data 
df_ag = df.groupby(['Week']).sum().reset_index()

## Create Dataframe fro Atlanta data
df_atl = df[df['market'] == 'Atlanta']
df_atl['Week'] = df_atl['Week'].astype(int)
df_atl = df_atl.groupby(['Week']).sum().reset_index()
df_atl = pd.DataFrame(np.transpose(df_atl))
df_atl.loc[['Week'], 0:132] = wkdat
df_atl = df_atl.reset_index()
df_atl = df_atl.set_value(0, 'index', 'Atlanta')
x = pd.Series([" "], index = df_atl.index)
df_atl = df_atl.append(x, ignore_index=True)


## Create Dataframe for Seattle data
df_sea = df[df['market'] == 'Seattle']
df_sea['Week'] = df_sea['Week'].astype(int)
df_sea = df_sea.groupby(['Week']).sum().reset_index()
df_sea = pd.DataFrame(np.transpose(df_sea))
df_sea.loc[['Week'], 0:132] = wkdat
df_sea = df_sea.reset_index()
df_sea = df_sea.set_value(0, 'index', 'Seattle')


## Bind datasets together and remove NaNs
frame = [df_ag, df_atl, df_sea]
df_f = pd.concat(frame)

df_f = df_f.fillna("")


## Write out the report
writer = pd.ExcelWriter('report.xlsx', engine = 'xlsxwriter')

df_f.to_excel(writer, sheet_name = 'Weekly Data Report', index = False, header = False)

workbook = writer.book
worksheet = writer.sheets['Weekly Data Report']
worksheet.set_column('A:A', 18)

writer.save()






