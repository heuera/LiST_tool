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
years = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030']

# Subset total births
t_births = births.iloc[:3, 7:]
t_births.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
t_births = t_births.set_index('Projection')
print(t_births)

# Subset LBW
lbw = outcomes.iloc[0:3, 7:]
lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
lbw = lbw.set_index('Projection')
num_lbw = pd.DataFrame((lbw.values/100)*t_births.values)
num_lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
num_lbw = num_lbw.set_index('Projection')
num_lbw.columns = years
c_num_lbw = pd.DataFrame(num_lbw.values - num_lbw.iloc[0].values, index=num_lbw.index, columns=num_lbw.columns)
print(c_num_lbw)

# Subset PT-AGA
pt_aga = outcomes.iloc[3:6, 7:]
pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
pt_aga = pt_aga.set_index('Projection')
num_pt_aga = pd.DataFrame((pt_aga.values/100)*t_births.values)
num_pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
num_pt_aga = num_pt_aga.set_index('Projection')
num_pt_aga.columns = years
c_num_pt_aga = pd.DataFrame(num_pt_aga.values - num_pt_aga.iloc[0].values, index=num_pt_aga.index, columns=num_pt_aga.columns)
print(c_num_pt_aga)

# Subset PT-SGA
pt_sga = outcomes.iloc[6:9, 7:]
pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
pt_sga = pt_sga.set_index('Projection')
num_pt_sga = pd.DataFrame((pt_sga.values/100)*t_births.values)
num_pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
num_pt_sga = num_pt_sga.set_index('Projection')
num_pt_sga.columns = years
c_num_pt_sga = pd.DataFrame(num_pt_sga.values - num_pt_sga.iloc[0].values, index=num_pt_sga.index, columns=num_pt_sga.columns)
print(c_num_pt_sga)

# Subset term-AGA
term_aga = outcomes.iloc[9:12, 7:]
term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
term_aga = term_aga.set_index('Projection')
num_term_aga = pd.DataFrame((term_aga.values/100)*t_births.values)
num_term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
num_term_aga = num_term_aga.set_index('Projection')
num_term_aga.columns = years
c_num_term_aga = pd.DataFrame(num_term_aga.values - num_term_aga.iloc[0].values, index=num_term_aga.index, columns=num_term_aga.columns)
print(c_num_term_aga)

# Subset term-SGA
term_sga = outcomes.iloc[12:15, 7:]
term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
term_sga = term_sga.set_index('Projection')
num_term_sga = pd.DataFrame((term_sga.values/100)*t_births.values)
num_term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
num_term_sga = num_term_sga.set_index('Projection')
num_term_sga.columns = years
c_num_term_sga = pd.DataFrame(num_term_sga.values - num_term_sga.iloc[0].values, index=num_term_sga.index, columns=num_term_sga.columns)
print(c_num_term_sga)
