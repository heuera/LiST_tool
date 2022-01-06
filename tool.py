import numpy as np
import pandas as pd
import openpyxl

# Read in sheets from dataset
births = pd.read_excel('/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx', sheet_name = '1. Total Births', skiprows=0, header=1)
outcomes = pd.read_excel('/Users/austinheuer/Desktop/LiST/ethiopia_out.xlsx', sheet_name = '2. Birth Outcomes (percent)', skiprows=0, header=1)

# Organize data
outcomes.columns = ["File name", "Country",	"ISO 3166-1 alpha-3", "Subnational region",	"Module",	"Indicator",	"Configuration", "2017", "2018", "2019",	"2020",	"2021",	"2022",	"2023",	"2024",	"2025",	"2026",	"2027",	"2028",	"2029",	"2030"]
outcomes = outcomes.sort_values(by='Configuration') # order: LBW, PT-AGA, PT-SGA, Term-AGA, Term-SGA
outcomes = outcomes.reset_index(drop=True) # resets index to match the sorted dataframe
labs = ['Baseline', 'LiST 1', 'LiST 2']

# Subset total births
t_births = births.iloc[:3, 7:]
t_births.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(t_births)

# Subset LBW
lbw = outcomes.iloc[0:3, 7:]
lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(lbw)

# Subset PT-AGA
pt_aga = outcomes.iloc[3:6, 7:]
pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(pt_aga)

# Subset PT-SGA
pt_sga = outcomes.iloc[6:9, 7:]
pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(pt_sga)

# Subset term-AGA
term_aga = outcomes.iloc[9:12, 7:]
term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(term_aga)

# Subset term-SGA
term_sga = outcomes.iloc[12:15, 7:]
term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
print(term_sga)

