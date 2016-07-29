"""
Created on Tue Jun  7 10:56:00 2016

@author: Nick

Purpose: Transforms the Pipeline information exported from JIRA
            into a table of probability weighted monthly revenue 
            to be visualized in Power BI.

"""

import pandas as pd
import numpy as np
import datetime
import openpyxl as pyxl
import string
import xlsxwriter

from dateutil.relativedelta import relativedelta

def date_calc(start, duration):
    #Select Start Date
    x = start
    #Transform Start Date
    str(x)
    day = 1
    year = int(x[0:4])
    month = int(x[5:7])
    int(day)
    strt_dt = datetime.date(year,month,day)
    
    months = (12*duration)
    
    dates = []
    count = 0
    while count < months:
        m_track = strt_dt + relativedelta(months=+count)
        dates.append(m_track)
        count = count + 1
        
    
    return dates

def date_parse(d):
    x = d
    str(x)
    y = x.split('-')
    year = y[0]
    month = y[1]
    day = y[2]
    
    tempx = (year,month,day)    
    
    return tempx
    
def quarter_calc(month):
    if (1 <= month <=3):
        quarter = 1
    elif (4<=month<=6):
        quarter = 2
    elif (7<=month<=9):
        quarter = 3
    else:
        quarter = 4  
    return quarter
    
def combo(df,dictionary,column_list):
    
    w = 0
    while w < len(column_list):
        temp_list = []
        for index,row in df.iterrows():
            print('still going')
            
            tempa = index #take index of each row
            
            value = dictionary[tempa][:-1]
            
            temp_list.append(value[w])
            #name = col_list[w]
            #str(name)
            #new_df[name] = temp_list
        name = column_list[w]
        df[name] = temp_list
        w = w + 1
    
    return df

#Selects which Excel File will be used
 #Make sure it is the latest JIRA export        
file_name = 'JIRA Today 7.8.xlsx'

#Creates a Pandas Dataframe from imported Excel File
xl_file = pd.ExcelFile(file_name)
tempdf = xl_file.parse('Sheet1')
tempdf.to_csv('Excel to CSV.csv')

#Imports only Necessary columns from JIRA export file
fields = ['Summary','Status','Priority','Labels','Start Date','End Date','PoC', \
        'First Year Contract Value', 'p(win)','Government or Commercial', \
        'Product, Service, Both','Escalation Factor', 'Contract Duration']

df = pd.read_csv('Excel to CSV.csv', index_col = 0, usecols = fields, encoding='latin-1')

print('CSV File Read Successfully')

#Creates List of Unique Prime Customers
prime_list = []
for index, row in df.iterrows():
    temp = index
    prime_list.append(temp)

cust_list = list(set(prime_list))

#Creates Dictionary with Total Revenue for Each Contract
###Possibly delete with JIRA format
contract_value_dict = {}
for key, row in df.iterrows():
    name = key
    rev = row['First Year Contract Value']
    int(rev)
    if name in contract_value_dict.keys():
        x = contract_value_dict[name]
        int(x)
        values = x + rev
        contract_value_dict[name] = values
    else:
        contract_value_dict[name] = rev
            

#Reads Beginning and End Dates of Contract and
# creates a dictionary containg all of the months 
# that the contract will be active
total_dict = {}
date_dict = {}
for index, row in df.iterrows():
    start = row['Start Date']
    duration = row['Contract Duration']
    str(start)
    int(duration)
    
    date_list = date_calc(start,duration)
    date_dict[index]=date_list
    total_dict[index] = list(row)
    total_dict[index].append(date_list)
    
    
#Transforms Dates in Readable Format
new_date = {}
num_dict = {}    

for key, values in total_dict.items():
    months = values[-1]    
    new = []
    for this in months:
        x = this.strftime("%Y-%m-%d")
        new.append(x)
    months = new
    new_date[key] = months
    # Creates a Dictionary of the Length of each Contract
    num_dict[key] = int(len(months))
    
    
    
#Tracks which Month of Contract
temp_dict = new_date

outside_dict = {}

for key,values in temp_dict.items():
    inside_dict = {}
    temp = values

    r = 1
    while r < (len(temp)+1):
        temp1 = temp[r-1]
        inside_dict[temp1] = r
        r = r + 1
    outside_dict[key] = inside_dict
    

#Creates New DataFrame for Dates of Active Contracts
date_df = pd.DataFrame.from_dict(new_date, orient = "index")

date_df[0].fillna(np.Inf, inplace=True)
df2 = pd.concat([date_df[col] for col in date_df], axis=0)
df2.dropna(inplace=True)
df2[df2 == np.Inf] = np.NaN

#Creates Column for Active Contract Quarters
new_df = pd.DataFrame(df2, columns = ['Z_Date Active'])

year_list = []
month_list = []
day_list = []
q_list = []

for row in new_df['Z_Date Active']:
    str(row)
    temp = date_parse(row)
    year = temp[0]
    month = temp [1]
    day = temp[2]
    year_list.append(year)
    month_list.append(month)
    day_list.append(day)
    q_month = int(month)
    q = quarter_calc(q_month)
    form = 'Q'
    xyz = form + str(q)
    q_list.append(xyz)   
    
new_df['Z_Year_New'] = year_list
new_df['Z_Month_New'] = month_list
new_df['Z_Day_New'] = day_list
new_df['Z_Quarter_New'] = q_list


print('Dataframe Expanded into Monthly values')

# Adds Column for Contract Length in Months to New_DF
num_list = []
for index, row in new_df.iterrows():
    if index in num_dict:
        value = num_dict[index]
        num_list.append(value)
    else:
        num_list.append('Error')
new_df['Z_Contract Length'] = num_list

#Adds Column that Tracks Contract Months Number
month_tracker = []
for index, row in new_df.iterrows():
    if index in outside_dict:
        values = outside_dict[index]
        comp1 = row['Z_Date Active']
        if comp1 in values:
            new_val = outside_dict[index][comp1]
            month_tracker.append(new_val)
new_df['Z_Contract Month Number'] = month_tracker


col_list = list(df.columns.values)

combined_df = combo(new_df,total_dict,col_list)   
print('DataFrames Combined')

factor_value = []
for index, row in combined_df.iterrows():
    bv = float(row['First Year Contract Value'])
    prob = float(row['p(win)'])
    q_rev = bv*prob
    month_rev = q_rev/12
    factor_value.append(month_rev)
combined_df['Factored Revenue'] = factor_value


#Adds Column that Calculates Actual Monthly Revenue using
#   escalation factors assuming Annual
actual_rev_list = []
for index,row in combined_df.iterrows():
    a = index
    b = float(row['Escalation Factor'])
    c = int(row['Z_Contract Month Number'])
    d = float(row['Factored Revenue'])
    e = int(row['Z_Contract Length'])
    
    if b == 0:
        value = d
    else:
        interval = 12.01
            
        check = c//interval
        if check == 0:
            value = d
        else:
            value = (d*((1+b)**(check)))
    
    actual_rev_list.append(value)

combined_df['Actual Month Revenue'] = actual_rev_list


#Adds a Column for classifying p(win) segments

prob_seg = []
for index,row in combined_df.iterrows():
    a = float(row['p(win)'])
    if a < 0.5:
        seg = 'Low'
    elif 0.5<a<0.8:
        seg = 'Medium'
    elif 0.8<a<1:
        seg = 'High'
    else:
        seg = 'Won'
    prob_seg.append(seg)
combined_df['Confidence - p(win)'] = prob_seg


#Creates Column that indexes the Fiscal Quarter in which the contract is active
qqq = []
for index,row in combined_df.iterrows():
    a = str(row['Z_Quarter_New'])
    b = str(row['Z_Year_New'])
    zzz = b+' '+a
    qqq.append(zzz)
combined_df['Quarter'] = qqq

combined_df.index.name = 'Prime Customer'

final_df = combined_df[['Z_Date Active','Z_Year_New','Z_Month_New','Quarter', \
    'Z_Contract Length','Z_Contract Month Number','Status','p(win)', \
    'Government or Commercial','Product, Service, Both','Escalation Factor',\
    'Actual Month Revenue','Confidence - p(win)']]


#Writes Dataframe to Excel File
file_name = "JIRA Demo 2.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

#Converts the dataframe to an XlsxWriter Excel object.
final_df.to_excel(writer, sheet_name='Sheet1')

print('DataFrame written to Excel')
#Adds Revenue Number Formatting
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

format1 = workbook.add_format({'num_format': '$#,##0.00'})

worksheet.set_column('M:M', None, format1)

#Formats Data into a Table in Excel
worksheet_table_header = writer.sheets['Sheet1']

end_row = len(final_df.index)
end_column = len(final_df.columns)
cell_range = xlsxwriter.utility.xl_range(0, 0, end_row, end_column)

final_df.reset_index(inplace=True)
header = [{'header': di} for di in final_df.columns.tolist()]
worksheet_table_header.add_table(cell_range,{'header_row': True,'first_column': True,'columns':header})


writer.save()

print('Done')
print('File called: ',file_name)
