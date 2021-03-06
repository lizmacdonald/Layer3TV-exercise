#!/usr/bin/env python2

#############################################################
# Layer3tv Exercise
#############################################################

import openpyxl as px
import pandas as pd
import numpy as np
import xlsxwriter as xlsxwriter


## Read in data
dat = px.load_workbook('subscriber-data.xlsx', data_only=True)
ws = dat['data']

df = pd.DataFrame(ws.values)
df.columns = [df.iloc[0]]
df = df.drop([0, ])


### set up week number column
df['activity_date'] = pd.to_datetime(df['activity_date'])

week = pd.to_datetime(df['activity_date']).dt.week
wk = []
for i in week:
	wk.append(str("%02d" % (i, )))

df['wksort'] = df['activity_date'].apply(lambda x: x.strftime('%Y')).astype(str) + wk

### Change the first couple of days in 2016 to be 2015 and 2017 to be 2016 due to them landing in the middle of the acutal week
d1 = df.wksort == '201653'
d2 = df.wksort == '201752'
colname = 'wksort'
df.loc[d1, colname] = '201553'
df.loc[d2, colname] = '201652'



## Set up month column
df['Month'] = df['activity_date'].apply(lambda x: x.strftime('%m-%Y'))


## Set up quarter column
quart = pd.to_datetime(df['activity_date']).dt.quarter
df['Quarter'] = df['activity_date'].apply(lambda x: x.strftime('%Y'))
df['Quarter'] = df['Quarter'].astype(str) + ' Q' + quart.astype(str)


## Change the type of the columns to numeric
for col in ['new_subscriptions', 'self_install', 'professional_install', 'disconnects', 'post_install_returns']:
    df[col] = df[col].astype(int)


## Set up net gain and net loss columns
df['Total Disconnects'] = df['post_install_returns'] + df['disconnects']

df['Net Gain'] = df['new_subscriptions'] - df['Total Disconnects']


## Set up column names to be used for each data subset
wkcol = ['Week', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

wk_fin = ['Week', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']

moncol = ['Month', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

mon_fin = ['Month', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']

qucol = ['Quarter', 'Beginning Subscribers', 'new_subscriptions','self_install', 'professional_install', 'Total Disconnects', 'post_install_returns', 'disconnects' ,'Net Gain', 'Ending Subs']

qu_fin = ['Quarter', 'Beginning Subscribers', 'Total Connects', 'Self Installs', 'Pro Installs', 'Total Disconnects', 'Post Install Returns', 'Disconnects',  'Net Gain', 'Ending Sub']


## Set up separate dataframes for each market

## AGGREGATE: subset for the weekly data
dc = df.groupby(['wksort']).sum().reset_index()
dc = dc.sort_values('wksort', ascending = True)
dc['Week'] = 'Week ' + dc['wksort'].map(lambda x: x[4:6]).astype(str) + " " + dc['wksort'].map(lambda x: x[0:4]).astype(str)


## AGGREGATE: Set up weekly cumulatvie beg and end subscription numbers and select appropriate columns
dc['Ending Subs'] = dc['Net Gain'].cumsum()
dc['Beginning Subscribers'] = dc['Ending Subs'] - dc['Net Gain']



## AGGREGATE: transpose weekly the data
dc = dc[wkcol]
dc.columns = wk_fin
dc = np.transpose(dc)


## AGGREGATE: add the word 'week' to the week data, clean up indexing
dc = dc.T.reset_index(drop=True).T
dc = dc.reset_index()
dc = dc.set_value(0, 'index', 'Aggregate Market')
blankrow = pd.Series([""], index = dc.index)
dc = dc.append(blankrow, ignore_index=True)


## AGGREGATE: select the monthly databases
dcm = df.groupby(['Month']).sum().reset_index()
dcm['Month'] = pd.to_datetime(dcm['Month'])
dcm = dcm.sort_values('Month', ascending = True)
dcm['Month'] = dcm['Month'].apply(lambda x: x.strftime('%b-%Y'))
dcm = dcm.reset_index(drop = True)


## AGGREGATE: Set up monthly cumulatvie beg and end subscription numbers and select appropriate columns
dcm['Ending Subs'] = dcm['Net Gain'].cumsum()
dcm['Beginning Subscribers'] = dcm['Ending Subs'] - dcm['Net Gain']

dcm = dcm[moncol]
dcm.columns = mon_fin


## AGGREGATE: transpose the monthly data
dcm['Month'] = pd.to_datetime(dcm['Month'])
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
dcq = np.transpose(dcq)


## AGGREGATE: clean up quarterly indexing
dcq = dcq.reset_index()
dcq = dcq.set_value(0, 'index', 'Aggregate Market')
dcq = dcq.append(blankrow, ignore_index=True)



## ATLANTA: select datasubset
da = df[df['market'] == 'Atlanta']



## ATLANTA: subset for the weekly data
daw = da.groupby(['wksort']).sum().reset_index()
daw = daw.sort_values('wksort', ascending = True)
daw['Week'] = 'Week ' + daw['wksort'].map(lambda x: x[4:6]).astype(str) + " " + daw['wksort'].map(lambda x: x[0:4]).astype(str)

## ATLANTA: Set up weekly cumulatvie beg and end subscription numbers
daw['Ending Subs'] = daw['Net Gain'].cumsum()
daw['Beginning Subscribers'] = daw['Ending Subs'] - daw['Net Gain']


## ATLANTA: transpose weekly the data
daw = daw[wkcol]
daw.columns = wk_fin

daw = np.transpose(daw)


## ATLANTA: add the word 'week' to the week data, clean up indexing
daw = daw.T.reset_index(drop=True).T
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
dam['Month'] = dam['Month'].apply(lambda x: x.strftime('%b-%Y'))
dam = np.transpose(dam)


## ATLANTA: clean up monthly indexing
dam = dam.reset_index()
dam = dam.set_value(0, 'index', 'Atlanta Market')
blankrow = pd.Series([""], index = dam.index)
dam = dam.append(blankrow, ignore_index=True)


## ATLANTA: select the quarterly datasbusets
daq = da.groupby(['Quarter']).sum().reset_index()
daq = daq.sort_values('Quarter', ascending = True)

## ATLANTA: Set up quarterly cumulatvie beg and end subscription numbers
daq['Ending Subs'] = daq['Net Gain'].cumsum()
daq['Beginning Subscribers'] = daq['Ending Subs'] - daq['Net Gain']

daq = daq[qucol]
daq.columns = qu_fin


## ATLANTA: transpose the quarterly data
daq = np.transpose(daq)


## ATLANTA: clean up quarterly indexing
daq = daq.reset_index()
daq = daq.set_value(0, 'index', 'Atlanta Market')
blankrow = pd.Series([""], index = daq.index)
daq = daq.append(blankrow, ignore_index=True)



## SEATTLE: Select data subset
ds = df[df['market'] == 'Seattle']



## SEATTLE: subset for the weekly data
dsw = ds.groupby(['wksort']).sum().reset_index()
dsw = dsw.sort_values('wksort', ascending = True)
dsw['Week'] = 'Week ' + dsw['wksort'].map(lambda x: x[4:6]).astype(str) + " " + dsw['wksort'].map(lambda x: x[0:4]).astype(str)



## SEATTLE: Set up weekly cumulatvie beg and end subscription numbers
dsw['Ending Subs'] = dsw['Net Gain'].cumsum()
dsw['Beginning Subscribers'] = dsw['Ending Subs'] - dsw['Net Gain']


## SEATTLE: transpose add the word 'week' to the week data, clean up indexing
dsw = dsw[wkcol]
dsw.columns = wk_fin

dsw = np.transpose(dsw)


## SEATTLE: add the word 'week' to the week data, clean up indexing
dsw = dsw.T.reset_index(drop=True).T
dsw['index'] = dsw.index
dsw = dsw.reset_index(drop = True)
cols = dsw.columns.tolist()
cols = [cols[-1]] + cols[ : -1]
dsw = dsw.reindex(columns = cols)
dsw = dsw.set_value(0, 'index', 'Seattle Market')
blankrow = pd.Series([""], index = dsw.index)
dsw = dsw.append(blankrow, ignore_index=True)


## SEATTLE: select the monthly datasbusets
dsm = ds.groupby(['Month']).sum().reset_index()
dsm['Month'] = pd.to_datetime(dsm['Month'])
dsm = dsm.sort_values('Month', ascending = True)
dsm['Month'] = dsm['Month'].apply(lambda x: x.strftime('%b-%Y'))
dsm = dsm.reset_index(drop = True)


## SEATTLE: Set up monthly cumulatvie beg and end subscription numbers
dsm['Ending Subs'] = dsm['Net Gain'].cumsum()
dsm['Beginning Subscribers'] = dsm['Ending Subs'] - dsm['Net Gain']

dsm = dsm[moncol]
dsm.columns = mon_fin


## SEATTLE: transpose the monthly data
dsm['Month'] = pd.to_datetime(dsm['Month'])
dsm['Month'] = dsm['Month'].apply(lambda x: x.strftime('%b-%Y'))
dsm = np.transpose(dsm)


## SEATTLE: clean up monthly indexing
dsm = dsm.reset_index()
dsm = dsm.set_value(0, 'index', 'Seattle Market')
blankrow = pd.Series([""], index = dsm.index)
dsm = dsm.append(blankrow, ignore_index=True)


## SEATTLE: select the quarterly datasbusets
dsq = ds.groupby(['Quarter']).sum().reset_index()
dsq = dsq.sort_values('Quarter', ascending = True)


## SEATTLE: Set up quarterly cumulatvie beg and end subscription numbers
dsq['Ending Subs'] = dsq['Net Gain'].cumsum()
dsq['Beginning Subscribers'] = dsq['Ending Subs'] - dsq['Net Gain']

dsq = dsq[qucol]
dsq.columns = qu_fin


## SEATTLE: transpose the quarterly data
dsq = np.transpose(dsq)


## SEATTLE: clean up quarterly indexing
dsq = dsq.reset_index()
dsq = dsq.set_value(0, 'index', 'Seattle Market')
blankrow = pd.Series([""], index = dsq.index)
dsq = dsq.append(blankrow, ignore_index=True)



## BIND datasets together and remove NaNs

## WEEK: create the weekly report
weekly_report = [dc, daw, dsw]
weekly_report = pd.concat(weekly_report)


## WEEK: add a header row to report
wk_title = ' Subscriber Report Week-Over-Week 2015-2017'
wt = pd.DataFrame(columns = weekly_report.columns)
wt = wt.set_value(len(wt), 'index', " ")
weekly_report.index = weekly_report.index + 1
weekly_report = wt.append(weekly_report)
weekly_report = weekly_report.set_value(0, 'index', wk_title)
weekly_report = weekly_report.fillna("")


## MONTH: create the monthly report 
monthly_report = [dcm, dam, dsm]
monthly_report = pd.concat(monthly_report)


## MONTH: add a header row to report
mon_title = ' Subscriber Report Month-Over-Month 2015-2017'
mt = pd.DataFrame(columns = monthly_report.columns)
mt = mt.set_value(len(mt), 'index', " ")
monthly_report.index = monthly_report.index + 1
monthly_report = mt.append(monthly_report)
monthly_report = monthly_report.set_value(0, 'index', mon_title)
monthly_report = monthly_report.fillna("")


## QUARTER: create the quarterly report 
quarterly_report = [dcq, daq, dsq]
quarterly_report = pd.concat(quarterly_report)


## QUARTER: add a header row to report
qu_title = ' Subscriber Report Quarter-Over-Quarter 2015-2017'
qt = pd.DataFrame(columns = quarterly_report.columns)
qt = qt.set_value(len(qt), 'index', " ")
quarterly_report.index = quarterly_report.index + 1
quarterly_report = qt.append(quarterly_report)
quarterly_report = quarterly_report.set_value(0, 'index', qu_title)
quarterly_report = quarterly_report.fillna("")



## Write out the report
writer = pd.ExcelWriter('Layer3tv_market_report.xlsx', engine = 'xlsxwriter')

weekly_report.to_excel(writer, sheet_name = 'Weekly Report', index = False, header = False)
monthly_report.to_excel(writer, sheet_name = 'Monthly Report', index = False, header = False)
quarterly_report.to_excel(writer, sheet_name = 'Quarterly Report', index
 = False, header = False)

workbook = writer.book


## format the weekly report
worksheet = writer.sheets['Weekly Report']
bold = workbook.add_format({'bold': True, 'font_size': 14})
bold2 = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 17)
worksheet.set_column('B:EE', 11)
worksheet.write(0, 0, wk_title, bold)
worksheet.write(1, 0, 'Aggregate Market', bold2)
worksheet.write(12, 0, 'Atlanta Market', bold2)
worksheet.write(23, 0, 'Seattle Market', bold2)


## format the monthly report
worksheet = writer.sheets['Monthly Report']
bold = workbook.add_format({'bold': True, 'font_size': 14})
bold2 = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 17)
worksheet.write(0, 0, mon_title, bold)
worksheet.write(1, 0, 'Aggregate Market', bold2)
worksheet.write(12, 0, 'Atlanta Market', bold2)
worksheet.write(23, 0, 'Seattle Market', bold2)


## format the quarterly report
worksheet = writer.sheets['Quarterly Report']
bold = workbook.add_format({'bold': True, 'font_size': 14})
bold2 = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 17)
worksheet.write(0, 0, qu_title, bold)
worksheet.write(1, 0, 'Aggregate Market', bold2)
worksheet.write(12, 0, 'Atlanta Market', bold2)
worksheet.write(23, 0, 'Seattle Market', bold2)


writer.save()
