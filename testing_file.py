import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
import xlsxwriter
import os
'''
# Assign directory
dir = '/Users/austinheuer/Desktop/LiST/LiST_tool/Folder'

# Iterate over files in folder
for filename in os.listdir(dir):
    if not filename.startswith('.'):
        filename = filename.lower()
        f = os.path.join(dir, filename)
        print(f)
'''

# Read in sheets from dataset
output = load_workbook("/Users/austinheuer/Desktop/LiST/LiST_tool/Folder/test1_output.xlsx")
births = pd.read_excel("/Users/austinheuer/Desktop/LiST/LiST_tool/Folder/test1_output.xlsx", sheet_name='1. Total Births', skiprows=0, header=1)
outcomes = pd.read_excel("/Users/austinheuer/Desktop/LiST/LiST_tool/Folder/test1_output.xlsx", sheet_name='2. Birth Outcomes (percent)', skiprows=0, header=1)
cov = pd.read_excel("/Users/austinheuer/Desktop/LiST/LiST_tool/Folder/test1_output.xlsx", sheet_name='1. Coverage of pregnancy interv', skiprows=0)
mn = pd.read_excel("/Users/austinheuer/Desktop/LiST/LiST_tool/Folder/test1_output.xlsx", sheet_name='2. Maternal nutrition', skiprows=0)


# Organize data
outcomes.columns = ["File name", "Country",	"ISO 3166-1 alpha-3", "Subnational region",	"Module",	"Indicator",	"Configuration", "2017", "2018", "2019",	"2020",	"2021",	"2022",	"2023",	"2024",	"2025",	"2026",	"2027",	"2028",	"2029",	"2030"]
outcomes = outcomes.sort_values(['Country', 'Configuration']) # order: LBW, PT-AGA, PT-SGA, Term-AGA, Term-SGA
outcomes = outcomes.reset_index(drop=True) # resets index to match the sorted dataframe
labs = ['Baseline', 'LiST 1', 'LiST 2']
years = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030']
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Subset total births
tb_start = np.linspace(0, 3, 2).tolist()
tb_start = list(map(int, tb_start))
tb_end = np.linspace(3, 6, 2).tolist()
tb_end = list(map(int, tb_end))
birth_dfs = {}
for start, end in zip(tb_start, tb_end):
    country_name = births.iloc[start, 1]
    t_births = births.iloc[start:end, 7:]
    t_births.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    t_births = t_births.set_index('Projection')
    t_births.index.name = "Total number of births"
    birth_dfs[country_name] = t_births

# Subset LBW
lbw_start = np.linspace(0, 15, 2).tolist()
lbw_start = list(map(int, lbw_start))
lbw_end = np.linspace(3, 18, 2).tolist()
lbw_end = list(map(int, lbw_start))
lowbw = {}
numb_lbw = {}
ch_num_lbw = {}
for start, end in zip(lbw_start, lbw_end):
    country_name = outcomes.iloc[start, 1]
    lbw = outcomes.iloc[start:end, 7:]
    lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    lbw = lbw.set_index('Projection')
    lbw.index.name = "% LBW births"
    lowbw[country_name] = lbw
    num_lbw = pd.DataFrame((lbw.values/100)*birth_dfs[country_name].values)
    num_lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    num_lbw = num_lbw.set_index('Projection')
    num_lbw.columns = years
    num_lbw.index.name = "Number of LBW births"
    numb_lbw[country_name] = num_lbw
    c_num_lbw = pd.DataFrame(num_lbw.values - num_lbw.iloc[0].values, columns=num_lbw.columns)
    c_num_lbw.index.name = "Change in number of LBW births"
    ch_num_lbw[country_name] = c_num_lbw

# Subset PT-AGA
ptaga_start = np.linspace(3, 18, 2).tolist()
ptaga_start = list(map(int, ptaga_start))
ptaga_end = np.linspace(6, 21, 2).tolist()
ptaga_end = list(map(int, ptaga_end))
pret_aga = {}
numb_pret_aga = {}
ch_num_pret_aga = {}
for start, end in zip(ptaga_start, ptaga_end):
    country_name = outcomes.iloc[start, 1]
    pt_aga = outcomes.iloc[start:end, 7:]
    pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    pt_aga = pt_aga.set_index('Projection')
    pt_aga.index.name = "% Pre-term AGA births"
    pret_aga[country_name] = pt_aga
    num_pt_aga = pd.DataFrame((pt_aga.values/100)*birth_dfs[country_name].values)
    num_pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    num_pt_aga = num_pt_aga.set_index('Projection')
    num_pt_aga.columns = years
    num_pt_aga.index.name = "Number of Pre-term AGA births"
    numb_pret_aga[country_name] = num_pt_aga
    c_num_pt_aga = pd.DataFrame(num_pt_aga.values - num_pt_aga.iloc[0].values, columns=num_pt_aga.columns)
    c_num_pt_aga.index.name = "Change in number of Pre-term AGA births"
    ch_num_pret_aga[country_name] = c_num_pt_aga

# Subset PT-SGA
ptsga_start = np.linspace(3, 18, 2).tolist()
ptsga_start = list(map(int, ptsga_start))
ptsga_end = np.linspace(6, 21, 2).tolist()
ptsga_end = list(map(int, ptsga_end))
pret_sga = {}
numb_pret_sga = {}
ch_num_pret_sga = {}
for start, end in zip(ptsga_start, ptsga_end):
    country_name = outcomes.iloc[start, 1]
    pt_sga = outcomes.iloc[start:end, 7:]
    pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    pt_sga = pt_sga.set_index('Projection')
    pt_sga.index.name = "% Pre-term SGA births"
    pret_sga[country_name] = pt_sga
    num_pt_sga = pd.DataFrame((pt_sga.values/100)*birth_dfs[country_name].values)
    num_pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    num_pt_sga = num_pt_sga.set_index('Projection')
    num_pt_sga.columns = years
    numb_pret_sga[country_name] = num_pt_sga
    num_pt_sga.index.name = "Number of Pre-term SGA births"
    c_num_pt_sga = pd.DataFrame(num_pt_sga.values - num_pt_sga.iloc[0].values, columns=num_pt_sga.columns)
    c_num_pt_sga.index.name = "Change in number of Pre-term SGA births"
    ch_num_pret_sga[country_name] = c_num_pt_sga

# Subset term-AGA
taga_start = np.linspace(9, 24, 2).tolist()
taga_start = list(map(int, taga_start))
taga_end = np.linspace(12, 27, 2).tolist()
taga_end = list(map(int, taga_end))
t_aga = {}
numb_t_aga = {}
ch_num_t_aga = {}
for start, end in zip(taga_start, taga_end):
    country_name = outcomes.iloc[start, 1]
    term_aga = outcomes.iloc[start:end, 7:]
    term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    term_aga = term_aga.set_index('Projection')
    term_aga.index.name = "% Term AGA births"
    t_aga[country_name] = term_aga
    num_term_aga = pd.DataFrame((term_aga.values/100)*birth_dfs[country_name].values)
    num_term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    num_term_aga = num_term_aga.set_index('Projection')
    num_term_aga.columns = years
    num_term_aga.index.name = "Number of Term AGA births"
    numb_t_aga[country_name] = num_term_aga
    c_num_term_aga = pd.DataFrame(num_term_aga.values - num_term_aga.iloc[0].values, columns=num_term_aga.columns)
    c_num_term_aga.index.name = "Change in number of Term AGA births"
    ch_num_t_aga[country_name] = c_num_term_aga

# Subset term-SGA
tsga_start = np.linspace(12, 27, 2).tolist()
tsga_start =list(map(int, tsga_start))
tsga_end = np.linspace(15, 30, 2).tolist()
tsga_end = list(map(int, tsga_end))
t_sga = {}
numb_t_sga = {}
ch_num_t_sga = {}
for start, end in zip(tsga_start, tsga_end):
    country_name = outcomes.iloc[start, 1]
    term_sga = outcomes.iloc[start:end, 7:]
    term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    term_sga = term_sga.set_index('Projection')
    term_sga.index.name = "% Term SGA births"
    t_sga[country_name] = term_sga
    num_term_sga = pd.DataFrame((term_sga.values/100)*birth_dfs[country_name].values)
    num_term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
    num_term_sga = num_term_sga.set_index('Projection')
    num_term_sga.columns = years
    num_term_sga.index.name = "Number of Term SGA births"
    numb_t_sga[country_name] = num_term_sga
    c_num_term_sga = pd.DataFrame(num_term_sga.values - num_term_sga.iloc[0].values, columns=num_term_sga.columns)
    c_num_term_sga.index.name = "Change in number of Term SGA births"
    ch_num_t_sga[country_name] = c_num_term_sga

'''
# Put each dataframe into correct position on Summary sheet
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
cov.to_excel(writer, sheet_name='Intervention Coverage', index=False)
mn.to_excel(writer, sheet_name='Interventions AF & efficacy', index=False)

writer.save()

'''