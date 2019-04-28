import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter
import openpyxl

#display all rows in datafram
pd.options.display.max_rows = 999
pd.options.display.max_columns = 999

#import any excel file for parsing
xls = pd.ExcelFile('Book1.xlsx')
df1 = pd.read_excel(xls, 'Sheet1')

df = df1.iloc[:, [0,3,4,5,6,7]]
df.columns = ['Month','Company','Bins','Wash','Type','Shared']

df['Company'] = df['Company'].str.replace("WILTYPE", "WILDTYPE", case = False)
df['Shared'] = df['Shared'].str.replace("shared", "Yes", case = False)
df['Shared'].fillna("No", inplace = True) 


autoclavedf = df.sort_values(by=['Company','Wash'])


print(autoclavedf)

# to find monthly subscription full service autclave
def autoclave_gravity():
		bins = df.groupby(['Company','Wash']).sum()
		bins = bins.reset_index(level=['Company','Wash'])
		binsmonthly = bins[bins['Wash'].str.contains('Grav.2')]
		binsmonthly.loc[binsmonthly.Bins < 1, 'Bill'] = '$0.00'
		binsmonthly.loc[binsmonthly.Bins >= 1, 'Bill'] = '$65.00'
		binsmonthly.loc[binsmonthly.Bins > 2, 'Bill'] = '$130.00'
		binsmonthly.loc[binsmonthly.Bins > 3, 'Bill'] = 'Count'
		
		return (binsmonthly)


binsmonthly = autoclave_gravity()

print(binsmonthly)

# to find count of gravity wash not in subscription(>3 bins)
def autoclave_gravity_count():
    companies_count = binsmonthly[binsmonthly['Bill'].str.contains('Count')]
    grav = companies_count['Company'].tolist()
    grav_companies = autoclavedf[autoclavedf['Wash'].str.contains('Grav.2')]
    gdf = grav_companies.iloc[:, [1,5]]
    cgdf = gdf.groupby(['Company', 'Shared']).size().reset_index(name='counts')
    gravity_counts = cgdf[cgdf['Company'].isin(grav)]
    gravity_counts['Shared'] = gravity_counts['Shared'].str.replace("Yes", "$25.00", case = False)
    gravity_counts['Shared'] = gravity_counts['Shared'].str.replace("No", "$50.00", case = False)
    return (gravity_counts)

    

gravity = autoclave_gravity_count()
print(gravity)

#to find price of each 
def autoclave_liquid():
		liquid = autoclavedf[autoclavedf['Wash'].str.contains('Liq')]
		liquid.loc[liquid.Type == 'TOP LOAD', 'Bill'] = '$35.00'
		liquid.loc[liquid.Type == 'FULL', 'Bill'] = '$50.00'
		liquid.loc[liquid.Shared == 'Yes', 'Bill'] = '$25.00'
		liquidcount = liquid.groupby(['Company', 'Bill']).size().reset_index(name='counts')

		return(liquidcount)

autoclave_liquid()

liquid = autoclave_liquid()


print(liquid)

writer = pd.ExcelWriter('autoclaveglasswash.xlsx', engine='openpyxl')

binsmonthly.to_excel(writer, sheet_name='Monthly', index=False)

gravity.to_excel(writer, sheet_name='Gravity', index=False)

liquid.to_excel(writer, sheet_name='Liquid', index=False)

writer.save()



