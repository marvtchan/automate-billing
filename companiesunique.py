import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter

#display all rows in dataframe
pd.options.display.max_rows = 999

#import any excel file for parsing
xls = pd.ExcelFile('conferencerooms.xls')
df1 = pd.read_excel(xls, 'Sheet 2')

#create variable for sum of hours billable
global bill_hours
bill_hours = []

#create variable for email domains that will be extracted
global emails
emails = []

#function for extracting email domain by uniqueness
def companies():
		users = df1['email'].str.split("@", expand= True)
		users.columns = ['user','domain']
		emails = users.domain.unique()
		return emails

#intiating function to set emails variable
companies()
emails = companies()

#parse files for billable hours
def weekly_run():
	for options in emails:
		print(options)
		results = df1[df1['email'].str.contains(options)]
		companyhours = float(sum(results['Hours']))
		bill_hours.append(companyhours)

weekly_run()
#double check
print(bill_hours)

#final dataframe for unique domains and billable hours
companiesdf1 = pd.DataFrame({'Companies': emails, 'Hours' : bill_hours})

print(companiesdf1)
