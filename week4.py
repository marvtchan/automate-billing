import pandas as pd
import numpy as np
import random 


import os
print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

# show files in current working directory

#input conference room sheets
#input data frames for each sheet

xls = pd.ExcelFile('conferencerooms.xlsx')
df1 = pd.read_excel(xls, 'Week 4')
#df2 = pd.read_excel(xls, 'Week 2')
#df3 = pd.read_excel(xls, 'Week 3')
#df4 = pd.read_excel(xls, 'Week 4')
#df5 = pd.read_excel(xls, 'Week 5')


#create variable for Mammoth

options1 = ['AesculaTech']

#single out mammoth Hours and create dataframe

rslt_df1_aescula = df1.loc[df1['Organization'].isin(options1)]

Total_Hours_aescula = sum(rslt_df1_aescula['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_aescula)

print('AesculaTech used ' + str(Total_Hours_aescula) + ' Hours in this Week')

#create variable for Alaunus

options2 = ['Alaunus']


rslt_df1_alaunus = df1.loc[df1['Organization'].isin(options2)]

Total_Hours_alaunus = sum(rslt_df1_alaunus['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_alaunus)

print('Alaunus used ' + str(Total_Hours_alaunus) + ' Hours in this Week')

#create variable for Alpine Roads

options3 = ['Alpine Roads','Alpine Roads, Inc.','Alpine Roads, Inc']


rslt_df1_alpine = df1.loc[df1['Organization'].isin(options3)]

Total_Hours_alpine = sum(rslt_df1_alpine['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_alpine)

print('Alpine Roads used ' + str(Total_Hours_alpine) + ' Hours in this Week')

#create variable for Applaud

options4 = ['Applaud Medical']

rslt_df1_applaud = df1.loc[df1['Organization'].isin(options4)]

Total_Hours_applaud = sum(rslt_df1_applaud['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_applaud)

print('Applaud used ' + str(Total_Hours_applaud) + ' Hours in this Week')

#create variable for Coral

options5 = ["Coral Genomics"]

rslt_df1_coral = df1.loc[df1['Organization'].isin(options5)]

Total_Hours_coral = sum(rslt_df1_coral['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_coral)

print('Coral Genomics used ' + str(Total_Hours_coral) + ' Hours in this Week')


#create variable for Cypre

options6 = ["Cypre, Inc."]

rslt_df1_cypre = df1.loc[df1['Organization'].isin(options6)]

Total_Hours_cypre = sum(rslt_df1_cypre['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_cypre)

print('Cypre used ' + str(Total_Hours_cypre) + ' Hours in this Week')

#create variable for Deciduous

options7 = ["Deciduous Therapeutics"]

rslt_df1_deciduous = df1.loc[df1['Organization'].isin(options7)]

Total_Hours_deciduous = sum(rslt_df1_deciduous['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_deciduous)

print('Deciduous used ' + str(Total_Hours_deciduous) + ' Hours in this Week')

#create variable for Delve

options8 = ["Delve Therapeutics"]

rslt_df1_delve = df1.loc[df1['Organization'].isin(options8)]

Total_Hours_delve = sum(rslt_df1_delve['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_delve)

print('Delve used ' + str(Total_Hours_delve) + ' Hours in this Week')

#create variable for Epiodyne

options9 = ['Epiodyne']

rslt_df1_epiodyne = df1.loc[df1['Organization'].isin(options9)]

Total_Hours_epiodyne = sum(rslt_df1_epiodyne['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_epiodyne)

print('Epiodyne used ' + str(Total_Hours_epiodyne) + ' Hours in this Week')
#create variable for Fountain

options10 = ["Fountain Therapeutics"]

rslt_df1_fountain = df1.loc[df1['Organization'].isin(options10)]

Total_Hours_fountain = sum(rslt_df1_fountain['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_fountain)

print('Fountain used ' + str(Total_Hours_fountain) + ' Hours in this Week')

#create variable for GeneTether

options11 = ['GeneTether, Inc']

rslt_df1_genetether = df1.loc[df1['Organization'].isin(options11)]

Total_Hours_genetether = sum(rslt_df1_genetether['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_genetether)

print('GeneTether used ' + str(Total_Hours_genetether) + ' Hours in this Week')

#create variable for Gordian

options12 = ['Gordian Biotechnology']

rslt_df1_gordian = df1.loc[df1['Organization'].isin(options12)]

Total_Hours_gordian = sum(rslt_df1_gordian['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_gordian)

print('Gordian used ' + str(Total_Hours_gordian) + ' Hours in this Week')

#create variable for Graphwear

options13 = ['GraphWear Technologies Inc.']

rslt_df1_graphwear = df1.loc[df1['Organization'].isin(options13)]

Total_Hours_graphwear = sum(rslt_df1_graphwear['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_graphwear)

print('GraphWear used ' + str(Total_Hours_graphwear) + ' Hours in this Week')

#create variable for LogicInk

options14 = ["Logic.Ink", "LogicInk"]

rslt_df1_logicink = df1.loc[df1['Organization'].isin(options14)]

Total_Hours_logicink = sum(rslt_df1_logicink['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_logicink)

print('LogicInk used ' + str(Total_Hours_logicink) + ' Hours in this Week')


#create variable for Mammoth

options15 = ["Mammoth Diagnostics","Mammoth BioSciences","Mammoth Biosciences"]

rslt_df1_mammoth = df1.loc[df1['Organization'].isin(options15)]

Total_Hours_mammoth = sum(rslt_df1_mammoth['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_mammoth)

print('Mammoth used ' + str(Total_Hours_mammoth) + ' Hours in this Week')


#create variable for Mitokinin

options16 = ['Mitokinin INC']

rslt_df1_mitokinin = df1.loc[df1['Organization'].isin(options16)]

Total_Hours_mitokinin = sum(rslt_df1_mitokinin['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_mitokinin)

print('Mitokinin used ' + str(Total_Hours_mitokinin) + ' Hours in his Week')

#create variable for Mojo

options17 = ['Mojo Health Inc']

rslt_df1_mojo = df1.loc[df1['Organization'].isin(options17)]

Total_Hours_mojo = sum(rslt_df1_mojo['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_mojo)

print('Mojo used ' + str(Total_Hours_mojo) + ' Hours in this Week')

#create variable for Naked Biome

options18 = ['Naked Biome']

rslt_df1_naked = df1.loc[df1['Organization'].isin(options18)]

Total_Hours_naked = sum(rslt_df1_naked['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_naked)

print('Naked Biome used ' + str(Total_Hours_naked) + ' Hours in this Week')

#create variable for Nitrome Biosciences

options19 = ['Nitrome Biosciences']

rslt_df1_nitrome = df1.loc[df1['Organization'].isin(options19)]

Total_Hours_nitrome = sum(rslt_df1_nitrome['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_nitrome)

print('Nitrome used ' + str(Total_Hours_nitrome) + ' Hours in Week 3')


#create variable for OneSkin

options20 = ['OneSkin']

rslt_df1_oneskin = df1.loc[df1['Organization'].isin(options20)]

Total_Hours_oneskin = sum(rslt_df1_oneskin['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_oneskin)

print('OneSkin used ' + str(Total_Hours_oneskin) + ' Hours in this Week')

#create variable for Prellis

options21 = ['Prellis Biologics']

rslt_df1_prellis = df1.loc[df1['Organization'].isin(options21)]

Total_Hours_prellis = sum(rslt_df1_prellis['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_prellis)

print('Prellis used ' + str(Total_Hours_prellis) + ' Hours in this Week')

#create variable for Provenance

options22 = ['Provenance Biofabrics, Inc','Provenance','Provenance Biofabrics Inc.']

rslt_df1_provenance = df1.loc[df1['Organization'].isin(options22)]

Total_Hours_provenance = sum(rslt_df1_provenance['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_provenance)

print('Provenance used ' + str(Total_Hours_provenance) + ' Hours in this Week')

#create variable for Quartz

options23 = ["Quartz Therapeutics"]

rslt_df1_quartz = df1.loc[df1['Organization'].isin(options23)]

Total_Hours_quartz = sum(rslt_df1_quartz['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_quartz)

print('Quartz used ' + str(Total_Hours_quartz) + ' Hours in this Week')

#create variable for Rumi

options24 = ['Rumi Scientific']

rslt_df1_rumi = df1.loc[df1['Organization'].isin(options24)]

Total_Hours_rumi = sum(rslt_df1_rumi['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_rumi)

print('Rumi used ' + str(Total_Hours_rumi) + ' Hours in this Week')


#create variable for Scaled BioLabs

options25 = ['Scaled Biolabs','Scaled Biolabs Inc.']

rslt_df1_scaled = df1.loc[df1['Organization'].isin(options25)]

Total_Hours_scaled = sum(rslt_df1_scaled['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_scaled)

print('Scaled used ' + str(Total_Hours_scaled) + ' Hours in this Week')


#create variable for Scribe

options26 = ["Scribe Biosciences"]

rslt_df1_scribe = df1.loc[df1['Organization'].isin(options26)]

Total_Hours_scribe = sum(rslt_df1_scribe['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_scribe)

print('Scribe used ' + str(Total_Hours_scribe) + ' Hours in this Week')

#create variable for Siolta

options27 = ['Siolta','Siolta Therapeutics']

rslt_df1_siolta = df1.loc[df1['Organization'].isin(options27)]

Total_Hours_siolta = sum(rslt_df1_siolta['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_siolta)

print('Siolta used ' + str(Total_Hours_siolta) + ' Hours in this Week')

#create variable for Soteria

options28 = ["Soteria Biotherapeutics",'Soteria Biotherapeutics, Inc.']

rslt_df1_soteria = df1.loc[df1['Organization'].isin(options28)]

Total_Hours_soteria = sum(rslt_df1_soteria['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_soteria)

print('Soteria used ' + str(Total_Hours_soteria) + ' Hours in this Week')

#create variable for SyntheX

options29 = ["SyntheX, Inc."]

rslt_df1_synthex= df1.loc[df1['Organization'].isin(options29)]

Total_Hours_synthex = sum(rslt_df1_synthex['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_synthex)

print('SyntheX used ' + str(Total_Hours_synthex) + ' Hours in this Week')


#create variable for Telo

options30 = ['Telo Therapeutics, Inc.']

rslt_df1_telo= df1.loc[df1['Organization'].isin(options30)]

Total_Hours_telo = sum(rslt_df1_telo['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_telo)

print('Telo used ' + str(Total_Hours_telo) + ' Hours in this Week')

#create variable for Wild Type

options31 = ['The Wild Type']

rslt_df1_wildtype= df1.loc[df1['Organization'].isin(options31)]

Total_Hours_wildtype = sum(rslt_df1_wildtype['Hours'])

print('\nResult dataframe1 :\n', rslt_df1_wildtype)

print('Wild Type used ' + str(Total_Hours_wildtype) + ' Hours in this Week')

frames = [Total_Hours_aescula,Total_Hours_alaunus,Total_Hours_alpine,Total_Hours_applaud,
Total_Hours_coral,Total_Hours_cypre,Total_Hours_deciduous,Total_Hours_delve,Total_Hours_epiodyne,
Total_Hours_fountain,Total_Hours_genetether,Total_Hours_gordian,Total_Hours_graphwear,Total_Hours_logicink,
Total_Hours_mammoth,Total_Hours_mitokinin,Total_Hours_mojo,Total_Hours_naked,Total_Hours_nitrome,
Total_Hours_oneskin,Total_Hours_prellis,Total_Hours_provenance,Total_Hours_quartz,Total_Hours_rumi,
Total_Hours_scaled,Total_Hours_scribe,Total_Hours_siolta,Total_Hours_soteria,Total_Hours_synthex,
Total_Hours_telo,Total_Hours_wildtype]

rslt_df1 = pd.DataFrame(frames, columns = ['Hours'])

companiesdf1 = pd.DataFrame({'Companies': ['AesculaTech','Alaunus','Alpine Roads','Applaud',
	'Coral Genomics','Cypre','Deciduous','Delve','Epiodyne','Fountain','GeneTether','Gordian','GraphWear',
	'LogicInk','Mammoth','Mitokinin','Mojo','Naked Biome','Nitrome','OneSkin','Prellis','Provenance',
	'Quartz','Rumi','Scaled BioLabs','Scribe','Siolta','Soteria','SyntheX','Telo','The Wild Type']})

rsltcompaniesdf1 = companiesdf1.join(rslt_df1)


print('\nResult Week4 :\n', rsltcompaniesdf1)



rsltcompaniesdf1.to_excel('output4.xlsx')

output4 = pd.read_excel('output4.xlsx', index_col=None, 
	na_values=['NA'], usecols = 'B:C')

m1 = output4.iloc[:,1].ge(4)
m2 = output4.iloc[:,1].le(4)

output4.loc[m1,'Hours'] -= 4
output4.loc[m2,'Hours'] = 0


print('\nResult Week4 Billable :\n', output4)

output4.to_excel('week4bill.xlsx')