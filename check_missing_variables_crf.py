import pandas as pd
import openpyxl
import re

df_crf_parameters = pd.read_excel('20240819_Euroblood CRFs_ALL_PROJECTS(6).xlsx', sheet_name='Parameters')
df_dataquality_rules = pd.read_excel('20240819_Euroblood CRFs_ALL_PROJECTS(6).xlsx', sheet_name='Data Quality Rules')

# check if all the variables of column FIELD NAME_TECH REDCap of df_crfs_parameters are in column FIELD NAME_TECH of df_dataquality_rules
for index, row in df_crf_parameters.iterrows():
    field_name_tech = row['FIELD NAME_TECH REDCap']
    matching_row = df_dataquality_rules[df_dataquality_rules['FIELD NAME_TECH'] == field_name_tech]
    
    if matching_row.empty:
        # if nan, skip
        if pd.isna(field_name_tech):
            continue
        else:
            print(f"Warning: {field_name_tech} not found in df_dataquality_rules")

# store a csv with the missing variables without spaces between rows
df_missing_variables = pd.DataFrame(columns=['FIELD NAME_TECH REDCap'])
for index, row in df_crf_parameters.iterrows():
    field_name_tech = row['FIELD NAME_TECH REDCap']
    matching_row = df_dataquality_rules[df_dataquality_rules['FIELD NAME_TECH'] == field_name_tech]
    
    if matching_row.empty:
        # if nan, skip
        if pd.isna(field_name_tech):
            continue
        else:
            df_missing_variables = df_missing_variables._append({'FIELD NAME_TECH REDCap': field_name_tech}, ignore_index=True)

df_missing_variables.to_csv('missing_variables.csv', index=False, header=False)