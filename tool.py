import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
import xlsxwriter
import os

# Assign directory
dir = '/Users/austinheuer/Desktop/LiST/LiST_tool/Folder'

# Iterate over files in folder
for filename in os.listdir(dir):
    if not filename.startswith('.'):
        filename = filename.lower()
        f = os.path.join(dir, filename)
        print(f)
        # Read in sheets from dataset
        output = load_workbook(f)
        births = pd.read_excel(f, sheet_name='1. Total Births', skiprows=0, header=1)
        outcomes = pd.read_excel(f, sheet_name='2. Birth Outcomes (percent)', skiprows=0, header=1)
        cov = pd.read_excel(f, sheet_name='1. Coverage of pregnancy interv', skiprows=0)
        mn = pd.read_excel(f, sheet_name='2. Maternal nutrition', skiprows=0)

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
        t_births.index.name = "Total number of births"

        # Subset LBW
        lbw = outcomes.iloc[0:3, 7:]
        lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        lbw = lbw.set_index('Projection')
        lbw.index.name = "% LBW births"
        num_lbw = pd.DataFrame((lbw.values/100)*t_births.values)
        num_lbw.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        num_lbw = num_lbw.set_index('Projection')
        num_lbw.columns = years
        num_lbw.index.name = "Number of LBW births"
        c_num_lbw = pd.DataFrame(num_lbw.values - num_lbw.iloc[0].values, columns=num_lbw.columns)
        c_num_lbw.index.name = "Change in number of LBW births"

        # Subset PT-AGA
        pt_aga = outcomes.iloc[3:6, 7:]
        pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        pt_aga = pt_aga.set_index('Projection')
        pt_aga.index.name = "% Pre-term AGA births"
        num_pt_aga = pd.DataFrame((pt_aga.values/100)*t_births.values)
        num_pt_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        num_pt_aga = num_pt_aga.set_index('Projection')
        num_pt_aga.columns = years
        num_pt_aga.index.name = "Number of Pre-term AGA births"
        c_num_pt_aga = pd.DataFrame(num_pt_aga.values - num_pt_aga.iloc[0].values, columns=num_pt_aga.columns)
        c_num_pt_aga.index.name = "Change in number of Pre-term AGA births"

        # Subset PT-SGA
        pt_sga = outcomes.iloc[6:9, 7:]
        pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        pt_sga = pt_sga.set_index('Projection')
        pt_sga.index.name = "% Pre-term SGA births"
        num_pt_sga = pd.DataFrame((pt_sga.values/100)*t_births.values)
        num_pt_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        num_pt_sga = num_pt_sga.set_index('Projection')
        num_pt_sga.columns = years
        num_pt_sga.index.name = "Number of Pre-term SGA births"
        c_num_pt_sga = pd.DataFrame(num_pt_sga.values - num_pt_sga.iloc[0].values, columns=num_pt_sga.columns)
        c_num_pt_sga.index.name = "Change in number of Pre-term SGA births"

        # Subset term-AGA
        term_aga = outcomes.iloc[9:12, 7:]
        term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        term_aga = term_aga.set_index('Projection')
        term_aga.index.name = "% Term AGA births"
        num_term_aga = pd.DataFrame((term_aga.values/100)*t_births.values)
        num_term_aga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        num_term_aga = num_term_aga.set_index('Projection')
        num_term_aga.columns = years
        num_term_aga.index.name = "Number of Term AGA births"
        c_num_term_aga = pd.DataFrame(num_term_aga.values - num_term_aga.iloc[0].values, columns=num_term_aga.columns)
        c_num_term_aga.index.name = "Change in number of Term AGA births"

        # Subset term-SGA
        term_sga = outcomes.iloc[12:15, 7:]
        term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        term_sga = term_sga.set_index('Projection')
        term_sga.index.name = "% Term SGA births"
        num_term_sga = pd.DataFrame((term_sga.values/100)*t_births.values)
        num_term_sga.insert(loc=0, column='Projection', value=labs, allow_duplicates=True)
        num_term_sga = num_term_sga.set_index('Projection')
        num_term_sga.columns = years
        num_term_sga.index.name = "Number of Term SGA births"
        c_num_term_sga = pd.DataFrame(num_term_sga.values - num_term_sga.iloc[0].values, columns=num_term_sga.columns)
        c_num_term_sga.index.name = "Change in number of Term SGA births"

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


