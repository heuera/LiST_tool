import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
import xlsxwriter
import os

# Read in sheets from dataset
output = load_workbook("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx")
births = pd.read_excel("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx", sheet_name='1. Total Births', skiprows=0, header=1)
outcomes = pd.read_excel("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx", sheet_name='2. Birth Outcomes (percent)', skiprows=0, header=1)
cov = pd.read_excel("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx", sheet_name='1. Coverage of pregnancy interv', skiprows=0)
mn = pd.read_excel("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx", sheet_name='2. Maternal nutrition', skiprows=0)

# Organize data
outcomes.columns = ["File name", "Country",	"ISO 3166-1 alpha-3", "Subnational region",	"Module",	"Indicator",	"Configuration", "2017", "2018", "2019",	"2020",	"2021",	"2022",	"2023",	"2024",	"2025",	"2026",	"2027",	"2028",	"2029",	"2030"]
outcomes = outcomes.sort_values(by='Configuration') # order: LBW, PT-AGA, PT-SGA, Term-AGA, Term-SGA
outcomes = outcomes.reset_index(drop=True) # resets index to match the sorted dataframe
labs = ['Baseline', 'LiST 1', 'LiST 2']
years = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030']

# Subset LBW
lbw = outcomes.iloc[0:3, 7:]
pt_aga = outcomes.iloc[3:6, 7:]
pt_sga = outcomes.iloc[6:9, 7:]
term_aga = outcomes.iloc[9:12, 7:]
term_sga = outcomes.iloc[12:15, 7:]
bo = [lbw, pt_aga, pt_sga, term_aga, term_sga]

# Subset total births
t_births = births.iloc[:3, 7:]
t_births.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
t_births = t_births.set_index('Projection')
t_births.index.name = "Total number of births"

tables = list()
for outcome in bo:
    outcome = pd.DataFrame(outcome)
    outcome.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    outcome = outcome.set_index('Projection')
    outcome.index.name = "% " + str(outcome) + " births"
    #print(outcome)
    number = pd.DataFrame((outcome.values/100)*t_births.values)
    number.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    number = number.set_index('Projection')
    number.columns = years
    number.index.name = "Number of " + str(outcome) + " births"
    #print(number)
    change = pd.DataFrame(number.values - number.iloc[0].values, columns=number.columns)
    change.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    change = change.set_index('Projection')
    change.columns = years
    change.index.name = "Change in number of " + str(outcome) + " births"
    #print(change)
    tables.append(outcome)
    tables.append(number)
    tables.append(change)

print(tables)


'''
writer = pd.ExcelWriter("/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx", engine='xlsxwriter')
workbook = writer.book
t_births.to_excel(writer, sheet_name='Summary', startrow=0, startcol=0)
s_row = [0, 6, 12, 18, 24, 30, 6, 12, 18, 24, 30, 6, 12, 18, 24, 30]
s_col = [0, 0, 0, 0, 0, 0, 16, 16, 16, 16, 16, 32, 32, 32, 32, 32]
for b_outcome in tables:
    for r in s_row:
        for c in s_col:
            b_outcome.to_excel(writer, sheet_name='Summary', startrow=r, startcol=c)
cov.to_excel(writer, sheet_name='Intervention Coverage')
mn.to_excel(writer, sheet_name='Interventions AF & efficacy')
writer.save()
'''

'''
writer = pd.ExcelWriter(f, engine='xlsxwriter')
workbook = writer.book
t_births.to_excel(writer, sheet_name='Summary', startrow=0, startcol=0)
num_pt_sga.to_excel(writer, sheet_name='Summary', startrow=6, startcol=0)
num_pt_aga.to_excel(writer, sheet_name='Summary', startrow=12, startcol=0)
num_term_sga.to_excel(writer, sheet_name='Summary', startrow=18, startcol=0)
num_term_aga.to_excel(writer, sheet_name='Summary', startrow=24, startcol=0)
num_lbw.to_excel(writer, sheet_name='Summary', startrow=30, startcol=0)
c_num_pt_sga.to_excel(writer, sheet_name='Summary', startrow=6, startcol=16)
c_num_pt_aga.to_excel(writer, sheet_name='Summary', startrow=12, startcol=16)
c_num_term_sga.to_excel(writer, sheet_name='Summary', startrow=18, startcol=16)
c_num_term_aga.to_excel(writer, sheet_name='Summary', startrow=24, startcol=16)
c_num_lbw.to_excel(writer, sheet_name='Summary', startrow=30, startcol=16)
pt_sga.to_excel(writer, sheet_name='Summary', startrow=6, startcol=32)
pt_aga.to_excel(writer, sheet_name='Summary', startrow=12, startcol=32)
term_sga.to_excel(writer, sheet_name='Summary', startrow=18, startcol=32)
term_aga.to_excel(writer, sheet_name='Summary', startrow=24, startcol=32)
lbw.to_excel(writer, sheet_name='Summary', startrow=30, startcol=32)
cov.to_excel(writer, sheet_name='Intervention Coverage')
mn.to_excel(writer, sheet_name='Interventions AF & efficacy')

writer.save()
'''