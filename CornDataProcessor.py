import pandas as pd
import datetime as dtt
import numpy as np
import matplotlib.pyplot as plt
import copy

CORNPRICE_EXCEL = "E:\\Desktop\\workstation\\database\\CornPice.xlsx"
FUTURES_EXCEL = "E:\\Desktop\\workstation\\database\\FuturesData.xlsx"
PIGPRICE_EXCEL = "E:\\Desktop\\workstation\\database\\PigData.xlsx"
SALERATE_EXCEL = "E:\\Desktop\\workstation\\database\\CornSalerate.xlsx"
NORTHPORT_EXCEL = "E:\\Desktop\\workstation\\database\\NorthPort.xlsx"
SORTHPORT_EXCEL = "E:\\Desktop\\workstation\\database\\SorthPort.xlsx"
DEEPPROCESSING_EXCEL = "E:\\Desktop\\workstation\\database\\DeepProcessOprationAndProfit.xlsx"
FACTORY_EXCEL = "E:\\Desktop\\workstation\\database\\FactoryInventories.xlsx"
ANALYSIS_EXCEL = "E:\\Desktop\\workstation\\Analysis.xlsx"
IMPORTANDEXPORT_EXCEL = "E:\\Desktop\\workstation\\database\\ImportAndExport.xlsx"
CFTC_EXCEL = "E:\\Desktop\\workstation\\database\\CFTCopi.xlsx"
CBOTFUTURES_EXCEL = "E:\\Desktop\\workstation\\database\\CBOTFuturesData.xlsx"

WEBDATA_EXCEL = "E:\\Desktop\\PyCode\\webdata.xlsx"

# df = pd.read_excel(excel_path)
# print(df.shape)
# print(df.iloc[3062,6].date())
# if not pd.isnull(df.iloc[3062,0]):
# 	print(type(df.iloc[3062,0]))
# else:
# 	print('Nodata')

def df_clean(df):
	delta1 = 0
	delta2 = 0
	delta3 = 0
	dflens = len(df.iloc[:,0])
	print(df.shape)
	clean_df = pd.DataFrame(np.zeros((dflens,8)))
	for i in range(dflens):
		if pd.notnull(df.iloc[i,0]):
			# print(df.iloc[i,0])
			clean_df.iloc[i,0] = df.iloc[i,0].date()
			clean_df.iloc[i,1] = df.iloc[i,1]
			try:
				while df.iloc[i,0] > df.iloc[i+delta1,2]:
					delta1 += 1
				while df.iloc[i,0] > df.iloc[i+delta2,4]:
					delta2 += 1
				while df.iloc[i,0] > df.iloc[i+delta3,6]:
					delta3 += 1
			except Exception as e:
				print(e)
				print(i+delta1)
				print(df.iloc[i,0])
				print(df.iloc[i+delta1-1,2])
			
			if df.iloc[i,0] < df.iloc[i+delta1,2]:
				clean_df.iloc[i,2] = np.NaN
				clean_df.iloc[i,3] = np.NaN
				delta1 -=1
			elif df.iloc[i,0] == df.iloc[i+delta1,2]:
				clean_df.iloc[i,2] = df.iloc[i+delta1,2].date()
				clean_df.iloc[i,3] = df.iloc[i+delta1,3]

			if df.iloc[i,0] < df.iloc[i+delta2,4]:
				clean_df.iloc[i,4] = np.NaN
				clean_df.iloc[i,5] = np.NaN
				delta2 -=1
			elif df.iloc[i,0] == df.iloc[i+delta2,4]:
				clean_df.iloc[i,4] = df.iloc[i+delta2,4].date()
				clean_df.iloc[i,5] = df.iloc[i+delta2,5]

			if df.iloc[i,0] < df.iloc[i+delta3,6]:
				print( df.iloc[i,0], df.iloc[i+delta3,6])
				clean_df.iloc[i,6] = np.NaN
				clean_df.iloc[i,7] = np.NaN
				delta3 -=1
			elif df.iloc[i,0] == df.iloc[i+delta3,6]:
				clean_df.iloc[i,6] = df.iloc[i+delta3,6].date()
				clean_df.iloc[i,7] = df.iloc[i+delta3,7]

	clean_df = clean_df[clean_df.iloc[:,0]!=0]
	return clean_df
######################################################################
# corn basis process
def basis_cal(clean_df_orig):
	clean_df = clean_df_orig.copy()
	basis_df = pd.DataFrame(np.zeros((len(clean_df),4)), index=clean_df.index, columns=['basis1', 'basis5', 'basis9', 'monthday'])
	for i in range(len(basis_df.index)):
		# print(i.month)
		# print(i.day)
		# print("8"*30)
		# print(basis_df.index[i])
		basis_df.iloc[i,3] = basis_df.index[i].strftime('%m-%d')
		if basis_df.index[i].month <= 3:
			basis_df.iloc[i,0] = np.NaN	
			if basis_df.index[i].month == 1:
				if basis_df.index[i].day < 10:
					basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]
				# else:
				# 	basis_df.iloc[i,0] = np.NaN			
			basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]			
			basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]		
		elif basis_df.index[i].month == 4:
			basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]			
			basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]			
			basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]
		elif basis_df.index[i].month > 4 and basis_df.index[i].month < 8:
			basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]	
			basis_df.iloc[i,1] = np.NaN	
			if basis_df.index[i].month == 5:
				if basis_df.index[i].day < 10:
					basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]
				# else:		
				# 	basis_df.iloc[i,1] = np.NaN			
			basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]
		elif basis_df.index[i].month == 8:
			basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]			
			basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]			
			basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]
		elif basis_df.index[i].month > 8 and basis_df.index[i].month < 12:
			basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]			
			basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]	
			basis_df.iloc[i,2] = np.NaN
			if basis_df.index[i].month == 9:
				if basis_df.index[i].day < 10:
					basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]
				# else:		
				# 	basis_df.iloc[i,2] = np.NaN
		elif basis_df.index[i].month == 12:
			basis_df.iloc[i,0] = clean_df.iloc[i,6] - clean_df.iloc[i,0]			
			basis_df.iloc[i,1] = clean_df.iloc[i,6] - clean_df.iloc[i,1]			
			basis_df.iloc[i,2] = clean_df.iloc[i,6] - clean_df.iloc[i,2]		
	return basis_df
######################################################################
# corn year basis process
def basis_year_cal(basis_df_orig, contract_month=1):
	basis_df = basis_df_orig.copy()
	basis_df['date'] = basis_df.index
	# print(basis_df.loc[basis_df.index[0],'date'].date())
	# print(type(basis_df.loc[basis_df.index[0],'date'].date()))
	minyear = basis_df.loc[basis_df.index[0],'date'].year
	maxyear = basis_df.loc[basis_df.index[-1],'date'].year
	df_d = {}
	for year in range(minyear,maxyear+1,1):
		d1 = dtt.datetime(year,contract_month+3,1)
		d2 = dtt.datetime(year+1,contract_month+3,1)
		df_d['%s'%year] = basis_df[(basis_df['date']>=d1) & (basis_df['date']<d2)]

	startday = dtt.datetime(2000,contract_month+3,1)
	endday = dtt.datetime(2001,contract_month+3,1)
	oneday = dtt.timedelta(days=1)
	columns = ['%sbasis%s'% (i, contract_month) for i in range(minyear+1,maxyear+2)]
	rows = pd.date_range(startday, endday-oneday)
	basis_year_df = pd.DataFrame(np.zeros([len(rows),len(columns)]), index=rows, columns=columns)

	d = startday
	while d<endday:
		dstr = d.strftime('%m-%d')
		l = []
		for key,value in df_d.items():
			if dstr in list(value['monthday']):
				basis = value[value['monthday']==dstr].loc[:,'basis%s'%contract_month].values[0]
				if pd.isnull(basis):
					basis = np.NaN
			else:
				basis = np.NaN
			l.append(basis)
		if l.count(np.NaN)<len(columns):
			basis_year_df.loc[d,:] = l	
		else:
			basis_year_df.drop(labels=[d], axis=0, inplace=True)
		d += oneday
	# basis_year_df.fillna(method='bfill', limit=2, inplace=True)
	return basis_year_df
######################################################################
# corn month spread process
def month_spread_cal(clean_df_orig, recent=9, far=1, contract='corn'):
	clean_df = clean_df_orig.copy()
	clean_df['date'] = clean_df.index
	for i in clean_df.index:
		clean_df.loc[i,'monthday'] = i.strftime('%m-%d')
	minyear = clean_df.loc[clean_df.index[0],'date'].year
	maxyear = clean_df.loc[clean_df.index[-1],'date'].year
	if contract=='corn':
		a = 0
	elif contract =='cs':
		a = 3
	# spread fenpian
	df_d = {}
	for year in range(minyear,maxyear+1,1):
		if recent==9 and far==1:
			d1 = dtt.datetime(year,far+1,1)
			d2 = dtt.datetime(year,recent,1)
			spread_df = clean_df.iloc[:,a+2]-clean_df.iloc[:,a+0]
		elif recent==1 and far==5:
			d1 = dtt.datetime(year,far+1,1)
			d2 = dtt.datetime(year+1,recent,1)
			spread_df = clean_df.iloc[:,a+0]-clean_df.iloc[:,a+1]
		elif recent==5 and far==9:
			d1 = dtt.datetime(year,far+1,1)
			d2 = dtt.datetime(year+1,recent,1)
			spread_df = clean_df.iloc[:,a+1]-clean_df.iloc[:,a+2]
		spread_df = spread_df.to_frame()
		spread_df['date'] = clean_df.index
		spread_df['monthday'] = clean_df['monthday']
		df_d['%s'%year] = spread_df[(spread_df['date']>=d1) & (spread_df['date']<d2)]

	# spread year cal
	if recent==9 and far==1:
		startday = dtt.datetime(2000,far+1,1)
		endday = dtt.datetime(2000,recent,1)
		columns = ['%s%s_%s'% (i, recent, far) for i in range(minyear,maxyear+1)]
	elif recent==1 and far==5:
		startday = dtt.datetime(2000,far+1,1)
		endday = dtt.datetime(2001,recent,1)
		columns = ['%s%s_%s'% (i, recent, far) for i in range(minyear+1,maxyear+2)]
	elif recent==5 and far==9:
		startday = dtt.datetime(2000,far+1,1)
		endday = dtt.datetime(2001,recent,1)
		columns = ['%s%s_%s'% (i, recent, far) for i in range(minyear+1,maxyear+2)]
	# startday = dtt.datetime(2000,d1.month,1)
	# endday = dtt.datetime(2001,d2.month,10)
	oneday = dtt.timedelta(days=1)
	rows = pd.date_range(startday, endday-oneday)
	spread_year_df = pd.DataFrame(np.zeros([len(rows),len(columns)]), index=rows, columns=columns)

	d = startday
	while d<endday:
		dstr = d.strftime('%m-%d')
		l = []
		for key,value in df_d.items():
			if dstr in list(value['monthday']):
				spread = value[value['monthday']==dstr].iloc[0,0]
				if pd.isnull(spread):
					spread = np.NaN
			else:
				spread = np.NaN
			l.append(spread)
		if l.count(np.NaN)<len(columns):
			spread_year_df.loc[d,:] = l	
		else:
			spread_year_df.drop(labels=[d], axis=0, inplace=True)
		d += oneday
	# spread_year_df.fillna(method='bfill', limit=2, inplace=True)

	return spread_year_df
######################################################################
# cs month spread process
# def cs_month_spread_cal(clean_df_orig):
# 	clean_df = clean_df_orig.copy()
######################################################################
# corn cs spread process
def corn_cs_spread_cal(clean_df_orig):
	clean_df = clean_df_orig.copy()
	corn_cs_spread_df = pd.DataFrame(np.zeros((len(clean_df),4)), index=clean_df.index, columns=['corncs1', 'corncs5', 'corncs9', 'monthday'])
	for i in range(len(corn_cs_spread_df.index)):
		corn_cs_spread_df.iloc[i,3] = corn_cs_spread_df.index[i].strftime('%m-%d')
		if pd.notnull(clean_df.iloc[i,0]) and pd.notnull(clean_df.iloc[i,3]):
			corn_cs_spread_df.iloc[i,0] = clean_df.iloc[i,0] - clean_df.iloc[i,3]	
		else:
			corn_cs_spread_df.iloc[i,0] = np.NaN

		if pd.notnull(clean_df.iloc[i,1]) and pd.notnull(clean_df.iloc[i,4]):		
			corn_cs_spread_df.iloc[i,1] = clean_df.iloc[i,1] - clean_df.iloc[i,4]
		else:
			corn_cs_spread_df.iloc[i,1] = np.NaN

		if pd.notnull(clean_df.iloc[i,2]) and pd.notnull(clean_df.iloc[i,5]):					
			corn_cs_spread_df.iloc[i,2] = clean_df.iloc[i,2] - clean_df.iloc[i,5]
		else:
			corn_cs_spread_df.iloc[i,2] = np.NaN

		# if pd.isnull(corn_cs_spread_df.iloc[i,0]) and pd.isnull(corn_cs_spread_df.iloc[i,1]) and pd.isnull(corn_cs_spread_df.iloc[i,2]):
		# 	corn_cs_spread_df.drop(labels=[corn_cs_spread_df.index[i]], axis=0, inplace=True)
	corn_cs_spread_df.dropna(axis=0, thresh =3, inplace=True)
	return corn_cs_spread_df

def corn_cs_spread_year_cal(corn_cs_spread_orig, contract_month=1):
	corn_cs_spread_df = corn_cs_spread_orig.copy()
	corn_cs_spread_df['date'] = corn_cs_spread_df.index
	# print(corn_cs_spread_df.loc[corn_cs_spread_df.index[0],'date'])
	# print(type(corn_cs_spread_df.loc[corn_cs_spread_df.index[0],'date']))
	# print(corn_cs_spread_df.loc[corn_cs_spread_df.index[0],'date'].year)
	minyear = corn_cs_spread_df.loc[corn_cs_spread_df.index[0],'date'].year
	maxyear = corn_cs_spread_df.loc[corn_cs_spread_df.index[-1],'date'].year

	df_d = {}
	for year in range(minyear,maxyear+1,1):
		d1 = dtt.datetime(year,contract_month+1,1)
		d2 = dtt.datetime(year+1,contract_month,10)
		df_d['%s'%year] = corn_cs_spread_df[(corn_cs_spread_df['date']>=d1) & (corn_cs_spread_df['date']<d2)]

	startday = dtt.datetime(2000,contract_month+1,1)
	endday = dtt.datetime(2001,contract_month,10)
	oneday = dtt.timedelta(days=1)
	columns = ['%sc-cs%s'% (i, contract_month) for i in range(minyear+1,maxyear+2)]
	rows = pd.date_range(startday, endday-oneday)
	corn_cs_spread_year_df = pd.DataFrame(np.zeros([len(rows),len(columns)]), index=rows, columns=columns)

	d = startday
	while d<endday:
		dstr = d.strftime('%m-%d')
		l = []
		for key,value in df_d.items():
			if dstr in list(value['monthday']):
				basis = value[value['monthday']==dstr].loc[:,'corncs%s'%contract_month].values[0]
				if pd.isnull(basis):
					basis = np.NaN
			else:
				basis = np.NaN
			l.append(basis)
		if l.count(np.NaN)<len(columns):
			corn_cs_spread_year_df.loc[d,:] = l	
		else:
			corn_cs_spread_year_df.drop(labels=[d], axis=0, inplace=True)
		d += oneday
	# corn_cs_spread_year_df.fillna(method='bfill', limit=2, inplace=True)
	return corn_cs_spread_year_df

######################################################################
# gather and out Data process per week
def data_year_process(df_origin, startmonth=10,columnsNum=1):
	df = df_origin.copy()
	df.columns = ['%s' % (i+1) for i in range(df.shape[1])]
	df['date'] = df.index
	clean_df = df[['%s'%columnsNum, 'date']]
	for date in clean_df.index:
		clean_df.loc[date,'weeknum'] = date.isocalendar()[1]
	# print(clean_df)
	minyear = clean_df.index[0].year
	maxyear = clean_df.index[-1].year
	df_d = {}
	for year in range(minyear, maxyear+1):
		d1 = dtt.datetime(year,startmonth,1)
		d2 = dtt.datetime(year+1,startmonth,1)
		df_d['%s/%s'%(year,year+1)] = clean_df[(clean_df['date']>=d1) & (clean_df['date']<d2)]

	weeknumlist = list(range(40,53))+list(range(1,40))
	columnsName = ['%s/%s' % (i,i+1) for i in range(minyear,maxyear+1)]
	columnsName += ['weeknum']
	port_year_df = pd.DataFrame(np.zeros([52,len(columnsName)]), index=pd.date_range('20011007', freq='W', periods=52), columns=columnsName)
	for i in range(len(weeknumlist)):
		port_year_df.iloc[i,len(columnsName)-1] = weeknumlist[i]
		for key,value in df_d.items():
			if weeknumlist[i] in list(value['weeknum']):
				port_year_df.loc[port_year_df.index[i],key] = value[value['weeknum']==weeknumlist[i]].loc[:,'%s'%columnsNum].values[0]
			else:
				port_year_df.loc[port_year_df.index[i],key] = np.NaN

	return port_year_df


cornPrice_df = pd.read_excel(CORNPRICE_EXCEL, sheet_name='prices', skiprows=0, index_col =0)
futures_df = pd.read_excel(FUTURES_EXCEL, sheet_name='futures', skiprows=0, index_col =0)
clean_df = pd.merge(futures_df, cornPrice_df[["jinzhou"]], how="left", on="date")

basis_df = basis_cal(clean_df)
# print(basis_df)
basis1_year_df = basis_year_cal(basis_df, contract_month=1)
# print(basis1_year_df)
basis5_year_df = basis_year_cal(basis_df, contract_month=5)
# print(basis5_year_df)
basis9_year_df = basis_year_cal(basis_df, contract_month=9)
# print(basis9_year_df)
corn_cs_spread_df = corn_cs_spread_cal(futures_df)
# print(corn_cs_spread_df)
corn_cs_spread1_year = corn_cs_spread_year_cal(corn_cs_spread_df, contract_month=1)
# print(corn_cs_spread1_year)
corn_cs_spread5_year = corn_cs_spread_year_cal(corn_cs_spread_df, contract_month=5)
# print(corn_cs_spread5_year)
corn_cs_spread9_year = corn_cs_spread_year_cal(corn_cs_spread_df, contract_month=9)
# print(corn_cs_spread9_year)
c9_1_spread_year_df = month_spread_cal(futures_df, recent=9, far=1, contract='corn')
c1_5_spread_year_df = month_spread_cal(futures_df, recent=1, far=5, contract='corn')
c5_9_spread_year_df = month_spread_cal(futures_df, recent=5, far=9, contract='corn')
# print(c5_9_spread_year_df)

###########################################################################
# process port carryout data
nport_df = pd.read_excel(NORTHPORT_EXCEL, sheet_name='Sheet1', skiprows=1, index_col =0)
if nport_df.iloc[-1,[2,5,8,11]].isnull().sum() > 0:
	portclean_df = nport_df[:-1]
else:
	portclean_df = nport_df
northPort_df = portclean_df.iloc[:,:12]
northPort_df.columns = ["jzgather", "jzout", "jzcarry", "byqgather", "byqout", "byqcarry", "blgather", "blout", "blcarry", "dywgather", "dywout", "dywcarry"]
northPort_df["northGather"] = portclean_df.iloc[:,[0,3,6,9]].apply(lambda x: x.sum(), axis=1)
northPort_df["northOut"] = portclean_df.iloc[:,[1,4,7,10]].apply(lambda x: x.sum(), axis=1)
northPort_df["northCarryOutSum"] = portclean_df.iloc[:,[2,5,8,11]].apply(lambda x: x.sum(), axis=1)
northPort_df["weekchange"] = northPort_df["northCarryOutSum"] - northPort_df["northCarryOutSum"].shift(1)
northPort_df = pd.merge(northPort_df, cornPrice_df[["jinzhou"]], how="left", on="date")
# print(northPort_df)
sport_df = pd.read_excel(SORTHPORT_EXCEL, sheet_name='Sheet1', skiprows=1, index_col =0)
if sport_df.iloc[-1,[2,5]].isnull().sum() > 0:
	portclean_df = sport_df[:-1]
else:
	portclean_df = sport_df
GDPort_df = portclean_df.iloc[:,[0,1,2,3,4,5,18,19,20]]
GDPort_df.columns = ["gdnmgather", "gdnmout", "gdnmcarry", "gdjkgather", "gdjkout", "gdjkcarry", "importcorn", "importUSAsorghum", "importbarley"]
GDPort_df["GDGather"] = portclean_df.iloc[:,[0,3]].apply(lambda x: x.sum(), axis=1)
GDPort_df["GDOut"] = portclean_df.iloc[:,[1,4]].apply(lambda x: x.sum(), axis=1)
GDPort_df["GDCarryOutSum"] = portclean_df.iloc[:,[2,5]].apply(lambda x: x.sum(), axis=1)
GDPort_df["weekchange"] = GDPort_df["GDCarryOutSum"] - GDPort_df["GDCarryOutSum"].shift(1)
# import profit
GDPort_df = pd.merge(GDPort_df, cornPrice_df[["guangdong"]], how="left", on="date")
GDPort_df["importProfit"] = GDPort_df["guangdong"] - GDPort_df["importcorn"]
# northport to sorthport profit
GDPort_df = pd.merge(GDPort_df, cornPrice_df[["jinzhou"]], how="left", on="date")
GDPort_df = pd.merge(GDPort_df, cornPrice_df[["jz_sk_freight"]], how="left", on="date")
GDPort_df = pd.merge(GDPort_df, northPort_df[["northCarryOutSum"]], how="left", on="date")
GDPort_df["ntosProfit"] = GDPort_df["guangdong"] - GDPort_df["jinzhou"] - GDPort_df["jz_sk_freight"] - 90
GDPort_df["nsRatio"] = GDPort_df["GDCarryOutSum"]/GDPort_df["northCarryOutSum"]
# print(GDPort_df)

gatherN_year_df = data_year_process(northPort_df, columnsNum=13)
gatherGD_year_df = data_year_process(GDPort_df, columnsNum=10)
outN_year_df = data_year_process(northPort_df, columnsNum=14)
outGD_year_df = data_year_process(GDPort_df, columnsNum=11)
gatherBBW_year_df = data_year_process(sport_df, columnsNum=7)
outBBW_year_df = data_year_process(sport_df, columnsNum=8)
gatherZZ_year_df = data_year_process(sport_df, columnsNum=10)
outZZ_year_df = data_year_process(sport_df, columnsNum=11)

###########################################################################
# pigPrice add guangdong price
pigPrice_df = pd.read_excel(PIGPRICE_EXCEL, sheet_name='PigPrice', skiprows=0, index_col =0)
pigPrice_df = pd.merge(pigPrice_df, cornPrice_df[["guangdong"]], how="left", on="date")
###########################################################################
# salerate add 5avg
salerate_df = pd.read_excel(SALERATE_EXCEL, sheet_name='salerate', skiprows=1, index_col =0)
salerate_avg_df = salerate_df.copy()
salerate_avg_df["hlj_avg5"] = salerate_df.iloc[:,[-15,-23,-31,-39,-47]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["jl_avg5"] = salerate_df.iloc[:,[-14,-22,-30,-38,-46]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["ln_avg5"] = salerate_df.iloc[:,[-13,-21,-29,-37,-45]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["nm_avg5"] = salerate_df.iloc[:,[-12,-20,-28,-36,-44]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["hb_avg5"] = salerate_df.iloc[:,[-11,-19,-27,-35,-43]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["sd_avg5"] = salerate_df.iloc[:,[-10,-18,-26,-34,-42]].apply(lambda x: x.mean(), axis=1)
salerate_avg_df["hn_avg5"] = salerate_df.iloc[:,[-9,-17,-25,-33,-41]].apply(lambda x: x.mean(), axis=1)
###########################################################################
# 115factory inventory
factory115 = pd.read_excel(FACTORY_EXCEL, sheet_name='Sheet1', skiprows=0, index_col =0)
factory115 = pd.merge(factory115, cornPrice_df[["jinzhou"]], how="left", on="date")
factory115['weekchange'] = factory115.iloc[:,0] - factory115.iloc[:,0].shift(1)
###########################################################################
# summarize
ffi_df = pd.read_excel(ANALYSIS_EXCEL, sheet_name='feedfactoryinventory', header=None)
chanqu_df = pd.read_excel(ANALYSIS_EXCEL, sheet_name='chanqu', header=None)
summarize_df = pd.read_excel(ANALYSIS_EXCEL, sheet_name='summarize', header=None)
visual_df = ffi_df.append([chanqu_df,summarize_df])
###########################################################################
importcorn_df = pd.read_excel(IMPORTANDEXPORT_EXCEL, sheet_name='Sheet1', skiprows=1, index_col=0)
importcorn_df = importcorn_df.sort_index()

if importcorn_df.index[0].month < 11:
	startyear = importcorn_df.index[0].year
else:
	startyear = importcorn_df.index[0].year + 1
if importcorn_df.index[-1].month < 10:
	endyear = importcorn_df.index[-1].year
else:
	endyear = importcorn_df.index[-1].year + 1

importcumsum_df = pd.DataFrame()
for year in range(startyear,endyear):
	cumsum_df = importcorn_df[(importcorn_df.index<dtt.datetime(year+1,10,1)) & (importcorn_df.index>dtt.datetime(year,9,30))].iloc[:,[0,4,8,12]].cumsum()
	cumsum_df.columns = ["%scorn"%year, "%sbarley"%year, "%ssorghum"%year, "%scassava"%year]
	cumsum_df.index = ["10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"][:cumsum_df.shape[0]]
	importcumsum_df = pd.concat([importcumsum_df,cumsum_df],axis=1, sort=False)
###########################################################################
# jinzhou price seasonal
jinzhou_df = pd.read_excel(CORNPRICE_EXCEL, sheet_name='prices', skiprows=1, index_col=0).iloc[:,0]
# print(type(jinzhou_df))
def price_year_process(df):
	df = df.to_frame()
	minyear = df.index[0].year
	maxyear = df.index[-1].year
	df['date'] = df.index
	df['monday'] = df['date'].apply(lambda x: x.strftime("%m-%d"))
	# print(df)
	year_df = pd.DataFrame(index=pd.date_range("20001001","20010930"))
	year_df["date"] = year_df.index
	year_df["monday"] = year_df['date'].apply(lambda x: x.strftime("%m-%d"))

	for year in range(minyear, maxyear+1):
		thisyear_df = df[(df.index>dtt.datetime(year-1,9,30))&(df.index<dtt.datetime(year,10,1))]
		thisyear_df.columns = ["%s/%s"%(year-1, year), "date" , "monday"]
		year_df = pd.merge(year_df, thisyear_df.iloc[:,[0,2]], how="left", on="monday")
	return year_df

jinzhouprice_year_df = price_year_process(jinzhou_df)
###########################################################################
# CFTC manage money interest CFTC_EXCEL
cftc_df = pd.read_excel(CFTC_EXCEL, sheet_name='Sheet1', skiprows=1, index_col=0)
cftc_df.columns = ["corn_opi", "corn_long", "corn_short", "wsr_opi", "wsr_long", "wsr_short", "whr_opi", "whr_long", "whr_short", "soy_opi", "soy_long", "soy_short"]
cftc_df['corn_netlong'] = cftc_df.iloc[:,1] - cftc_df.iloc[:,2]
cftc_df['corn_netlong_ratio'] = round(cftc_df['corn_netlong'] / cftc_df.iloc[:,0],2)
cftc_df['wsr_netlong'] = cftc_df.iloc[:,4] - cftc_df.iloc[:,5]
cftc_df['wsr_netlong_ratio'] = round(cftc_df['wsr_netlong'] / cftc_df.iloc[:,3],2)
cftc_df['whr_netlong'] = cftc_df.iloc[:,7] - cftc_df.iloc[:,8]
cftc_df['whr_netlong_ratio'] = round(cftc_df['whr_netlong'] / cftc_df.iloc[:,6],2)
cftc_df['soy_netlong'] = cftc_df.iloc[:,10] - cftc_df.iloc[:,11]
cftc_df['soy_netlong_ratio'] = round(cftc_df['soy_netlong'] / cftc_df.iloc[:,9],2)

cbotFuturesPrice_df = pd.read_excel(CBOTFUTURES_EXCEL, sheet_name='Sheet1', skiprows=0, index_col=0)
# print(cbotFuturesPrice_df)
cftc_df = pd.merge(cftc_df, cbotFuturesPrice_df, how="left", on="date")
cftc_data_df = cftc_df[['corn_opi', 'corn_netlong_ratio', 'corn_netlong', 'cornprice', 'wsr_opi', 'wsr_netlong_ratio', 'wsr_netlong', 'wsrprice', 'whr_opi', 'whr_netlong_ratio', 'whr_netlong', 'soy_opi', 'soy_netlong_ratio', 'soy_netlong', 'soyprice']]
###########################################################################



# write to excel
writer = pd.ExcelWriter(WEBDATA_EXCEL)
# clean_df.to_excel(writer, 'cleanpricedata')
# corn basis
basis_df.to_excel(writer, 'cornbasis')
basis1_year_df.to_excel(writer, 'cornyearbasis1')
basis5_year_df.to_excel(writer, 'cornyearbasis5')
basis9_year_df.to_excel(writer, 'cornyearbasis9')
# corn cs spread
corn_cs_spread_df.to_excel(writer, 'corncsspread')
corn_cs_spread1_year.to_excel(writer, 'corncsspread1year')
corn_cs_spread5_year.to_excel(writer, 'corncsspread5year')
corn_cs_spread9_year.to_excel(writer, 'corncsspread9year')
# corn month spread
c9_1_spread_year_df.to_excel(writer, 'c9_1_year')
c1_5_spread_year_df.to_excel(writer, 'c1_5_year')
c5_9_spread_year_df.to_excel(writer, 'c5_9_year')
# port year data
northPort_df.to_excel(writer, 'northPort')
GDPort_df.to_excel(writer, 'GDPort')
gatherN_year_df.to_excel(writer, 'gatherN_year')
gatherGD_year_df.to_excel(writer, 'gatherGD_year')
outN_year_df.to_excel(writer, 'outN_year')
outGD_year_df.to_excel(writer, 'outGD_year')
gatherBBW_year_df.to_excel(writer, 'gatherBBW_year')
outBBW_year_df.to_excel(writer, 'outBBW_year')
gatherZZ_year_df.to_excel(writer, 'gatherZZ_year')
outZZ_year_df.to_excel(writer, 'outZZ_year')

pigPrice_df.to_excel(writer, 'pigPrice')
salerate_avg_df.to_excel(writer, 'salerate_avg')

factory115.to_excel(writer, 'factory115')

visual_df.to_excel(writer, 'visual')

importcumsum_df.to_excel(writer, 'importcumsum')

jinzhouprice_year_df.to_excel(writer, 'jinzhouprice_year')

cftc_data_df.to_excel(writer, 'cftc_data')


writer.save()

















