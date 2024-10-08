import pandas as pd
import openpyxl
import re

# Load the data
df_iecs_file = pd.read_excel('radgenint_data_quality_check.xlsx')
df_custom_rules = df_iecs_file[['Variable / Field Name', 'Form Name', 'Field Label', 'Branching Logic (Show field only if...)', 'Custom data quality']]

# Initialize new columns
df_custom_rules['min date'] = ''
df_custom_rules['max date'] = ''
df_custom_rules['min alert'] = ''
df_custom_rules['max alert'] = ''
df_custom_rules['min value'] = ''
df_custom_rules['max value'] = ''

# Iterate through the rows and extract the values
for index, row in df_custom_rules.iterrows():
    custom_data_quality = row['Custom data quality']

    if pd.notna(custom_data_quality):
        # Split the custom_data_quality at '/' to separate alerts and value ranges
        parts = custom_data_quality.split(' / ')

        # Initialize flags to ensure we don't overwrite between value and alert conditions
        value_processed = False
        alert_processed = False

        # Iterate over the parts to extract value and alert ranges
        for part in parts:
            lower_alert_match = None
            higher_alert_match = None
            lower_value_match = None
            higher_value_match = None

            # Handle "alert" conditions
            if 'alert' in part or 'Alert' in part and not alert_processed:
                lower_alert_match = re.search(r'lower than (-?\d+\.?\d*)', part)  # Allow negative numbers
                higher_alert_match = re.search(r'higher than (-?\d+\.?\d*)', part)
                if lower_alert_match:
                    df_custom_rules.loc[index, 'min alert'] = lower_alert_match.group(1)
                if higher_alert_match:
                    df_custom_rules.loc[index, 'max alert'] = higher_alert_match.group(1)
                alert_processed = True  # Mark that alert has been processed

            # Handle "value range" conditions
            if ('value range should' in part or 'Value range' in part) and not value_processed:
                lower_value_match = re.search(r'lower than (-?\d+\.?\d*)', part)  # Allow negative numbers
                higher_value_match = re.search(r'higher than (-?\d+\.?\d*)', part)
                if higher_value_match:
                    df_custom_rules.loc[index, 'min value'] = higher_value_match.group(1)
                if lower_value_match:
                    df_custom_rules.loc[index, 'max value'] = lower_value_match.group(1)
                value_processed = True  # Mark that value range has been processed

    # Handle "Date range"
    if pd.notna(custom_data_quality) and 'Date range' in custom_data_quality:
        if 'Date range must be between [birth_date] and [death_date] or [episode_date]' in custom_data_quality:
            df_custom_rules.loc[index, 'min date'] = '[date_of_birth]'
            df_custom_rules.loc[index, 'max date'] = '[timestamp]'
    
    # Handle 'If not empty, date range'
    if pd.notna(custom_data_quality) and 'If not empty, date range' in custom_data_quality:
        if  'If not empty, date range must be between [birth_date] and [death_date] or [episode_date]' in custom_data_quality:
            df_custom_rules.loc[index, 'min date'] = '[date_of_birth]'
            df_custom_rules.loc[index, 'max date'] = '[timestamp]'

    # handle date_of_birth
    if pd.notna(custom_data_quality) and 'date_of_birth' in row['Variable / Field Name']:
        df_custom_rules.loc[index, 'min date'] = '[current_date] - 100 years'
        df_custom_rules.loc[index, 'max date'] = '[timestamp]'
    
    # handle death_date
    if pd.notna(custom_data_quality) and 'death_date' in row['Variable / Field Name']:
        df_custom_rules.loc[index, 'min date'] = '[date_of_birth]'
        df_custom_rules.loc[index, 'max date'] = '[timestamp]'
    
    if pd.notna(custom_data_quality) and 'year_immigration' in row['Variable / Field Name']:
        df_custom_rules.loc[index, 'min date'] = 'year([date_of_birth])'
        df_custom_rules.loc[index, 'max date'] = 'year([timestamp])'

# Save the updated DataFrame to a new Excel file
df_custom_rules.to_excel('radgenint_custom_rules_iecs.xlsx', index=False)


    
