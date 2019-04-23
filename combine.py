import os
import pandas as pd

print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

path = os.getcwd()

#list all files in current working directory
files = os.listdir(path)

#function for looping through all files ending in bill.xlsx
files_xls = [f for f in files if f[-9:] == 'bill.xlsx']

print(files_xls)

df = pd.DataFrame()

#loop through sheet 1 in each file
for f in files_xls:
	data = pd.read_excel(f, 'Sheet1')
	df = df.append(data)

#add companies and sum
df['Total'] = df.groupby(['Companies'])['Hours'].transform('sum')

new_df = df.drop_duplicates(subset=['Companies'])

cols = [1,3]

billdf = new_df[new_df.columns[cols]]


newbilldf = billdf.set_index('Companies')


print(newbilldf)

newbilldf.to_excel('conferenceroomsmarch.xlsx') 
  
