import pandas as pd
import openpyxl
import re

# Load the data
df_iecs_file = pd.read_excel('radeep_custom_rules_iecs.xlsx')
df_dictionary = pd.read_csv('RADeep2024New_DataDictionary_2024-10-08.csv')

# Check if values of columns 'min value' and 'max value' of df_iecs_file are in columns 
# 'Text Validation Min' and 'Text Validation Max' of df_dictionary
for index, row in df_iecs_file.iterrows():
    min_value = row['min value']
    max_value = row['max value']
    variable_name = row['Variable / Field Name']
    
    # Check and compare min values
    if pd.notna(min_value):
        min_value = float(min_value)  # Cast to float
        matching_row = df_dictionary[df_dictionary['Variable / Field Name'] == variable_name]
        
        if not matching_row.empty:
            text_validation_min = matching_row['Text Validation Min'].values[0]
            
            if pd.notna(text_validation_min):
                text_validation_min = float(text_validation_min)  # Cast to float if not NaN
            
            if text_validation_min != min_value:
                print(f"Error: {variable_name} min value {min_value} is not in Text Validation Min {text_validation_min} of df_dictionary")
            else:
                print(f"Success: {variable_name} min value {min_value} is in Text Validation Min of df_dictionary")
        else:
            print(f"Warning: {variable_name} not found in df_dictionary")

    # Check and compare max values
    if pd.notna(max_value):
        max_value = float(max_value)  # Cast to float
        matching_row = df_dictionary[df_dictionary['Variable / Field Name'] == variable_name]
        
        if not matching_row.empty:
            text_validation_max = matching_row['Text Validation Max'].values[0]
            
            if pd.notna(text_validation_max):
                text_validation_max = float(text_validation_max)  # Cast to float if not NaN
            
            if text_validation_max != max_value:
                print(f"Error: {variable_name} max value {max_value} is not in Text Validation Max {text_validation_max} of df_dictionary")
            else:
                print(f"Success: {variable_name} max value {max_value} is in Text Validation Max of df_dictionary")
        else:
            print(f"Warning: {variable_name} not found in df_dictionary")


# Substitute the values of columns 'Text Validation Min' and 'Text Validation Max' of df_dictionary 
# with the values of columns 'min value', 'max value' of df_iecs_file
for index, row in df_iecs_file.iterrows():
    min_value = row['min value']
    max_value = row['max value']
    variable_name = row['Variable / Field Name']
    
    if pd.notna(min_value):
        df_dictionary.loc[df_dictionary['Variable / Field Name'] == variable_name, 'Text Validation Min'] = min_value

    if pd.notna(max_value):
        df_dictionary.loc[df_dictionary['Variable / Field Name'] == variable_name, 'Text Validation Max'] = max_value

# DICTIONARY SHOULD BE UPDATED HERE (Save the updated dictionary)
# df_dictionary.to_csv('Updated_RADeep2024New_DataDictionary.csv', index=False)