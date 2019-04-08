import os
import pandas as pd

print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

path = os.getcwd()

files = os.listdir(path)

files_xls = [f for f in files if f[-9:] == 'bill.xlsx']

print(files_xls)

df = pd.DataFrame()

for f in files_xls:
	data = pd.read_excel(f, 'Sheet1')
	df = df.append(data)

df['Total'] = df.groupby(['Companies'])['Hours'].transform('sum')

new_df = df.drop_duplicates(subset=['Companies'])

cols = [1,3]

billdf = new_df[new_df.columns[cols]]


newbilldf = billdf.set_index('Companies')


print(newbilldf)

newbilldf.to_excel('conferenceroomsmarch.xlsx')