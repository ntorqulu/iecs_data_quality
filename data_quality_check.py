import pandas as pd
import openpyxl
import re

# read xlsx file, sheet CRF
df_crf = pd.read_excel('20240819_Euroblood CRFs_ALL_PROJECTS(6).xlsx', sheet_name='Data Quality Rules')
# get rows where column INTEGRA is 'x' or GenoMed4ALL is 'x' or RADeep is 'x'
# df_crf = df_crf[(df_crf['INTEGRA'] == 'x') | (df_crf['GenoMed4ALL'] == 'x') | (df_crf['RADeep'] == 'x')]
df_crf = df_crf[df_crf['GenoMed4ALL'] == 'x']
df_rules = pd.read_csv('RADeepGenomed4ALLINTEGRADataCo_DataQualityRules_2024-10-02.csv')
df_dictionary_redcap = pd.read_csv('Genomed4ALL2024New_DataDictionary_2024-10-08.csv')

# Get only the columns FIELD NAME_TECH, RULE for DATA QUALITY, STATUS, and Required
df_crf = df_crf[['FIELD NAME_TECH', 'UPDATED RULE for DATA QUALITY', 'STATUS']]

# create column REQUIRED to empty
df_crf['Required'] = ''
df_crf['REDCAP data quality'] = ''

# Split the 'FIELD NAME_TECH' column by '/' and expand into multiple rows
df_crf['FIELD NAME_TECH'] = df_crf['FIELD NAME_TECH'].str.split('/')  # Split by '/'
df_crf = df_crf.explode('FIELD NAME_TECH')  # Expand into multiple rows

# Remove leading/trailing spaces from FIELD NAME_TECH
df_crf['FIELD NAME_TECH'] = df_crf['FIELD NAME_TECH'].str.strip()

df_crf.loc[df_crf['STATUS'].str.startswith('Mandatory', na=False), 'Required'] = 'x'

df_merged = pd.merge(df_crf, 
                     df_dictionary_redcap[['Variable / Field Name', 'Form Name', 'Field Label', 'Required Field?', 'Field Type', 'Text Validation Type OR Show Slider Number', 'Branching Logic (Show field only if...)']], 
                     left_on='FIELD NAME_TECH', 
                     right_on='Variable / Field Name', 
                     how='left')

# check if all rows with 'Required Field?' == 'y' have 'Required' == 'x'
required_mismatch = df_merged[(df_merged['Required Field?'] == 'y') & (df_merged['Required'] != 'x')]
if not required_mismatch.empty:
    print("Warning: There are rows where 'Required Field?' is 'y' but 'Required' is not 'x'")
    print(required_mismatch)
else:
    print("All rows with 'Required Field?' == 'y' have 'Required' == 'x'")

# add variables from Variable / Field Name column from df_dictionary_redcap that are not in df_merged
df_merged = pd.concat([df_merged, df_dictionary_redcap[~df_dictionary_redcap['Variable / Field Name'].isin(df_merged['FIELD NAME_TECH'])]], ignore_index=True)

# Reorder the columns to match the expected output
df_merged = df_merged[['Variable / Field Name', 'Form Name', 'Field Label', 'Required Field?', 'Field Type', 'Branching Logic (Show field only if...)', 'Text Validation Type OR Show Slider Number', 'UPDATED RULE for DATA QUALITY', 'REDCAP data quality']]

# delete rows with empty 'Variable / Field Name'
df_merged = df_merged.dropna(subset=['Variable / Field Name'])

df_merged['Custom data quality'] = ''

df_rules_filtered = df_rules[df_rules['rule_name'].str.startswith('[', na=False)]
for index, row in df_rules_filtered.iterrows():
    # Extract the variable name from the rule_name (removing the brackets)
    match = re.match(r'\[(.*?)\]', row['rule_name'])  # Extracts the item within the brackets
    if match:
        variable_name = match.group(1)  # The item name from the rule_name
        rule_logic = row['rule_logic']
        
        # Step 4: Check if the variable is in FIELD NAME_TECH column of df_check
        mask = df_merged['Variable / Field Name'] == variable_name
        
        # Step 5: If variable is found and REDCAP data quality is empty, update with rule_logic
        df_merged.loc[mask, 'REDCAP data quality'] = rule_logic
# Extract variable names from rule_name and check for duplicates
df_rules_filtered['variable_name'] = df_rules_filtered['rule_name'].apply(lambda x: re.match(r'\[(.*?)\]', x).group(1) if re.match(r'\[(.*?)\]', x) else None)

# Step 7: Check for duplicates in the variable_name column
duplicate_rules = df_rules_filtered[df_rules_filtered.duplicated('variable_name', keep=False)]
print(duplicate_rules)

# Populate 'Custom data quality' column with the required rules, appending new rules if needed

for index, row in df_merged.iterrows():
    custom_quality_rules = []
    created_rule = []
    
    # Rule 1: If 'Required Field?' == 'y' and 'Branching Logic (Show field only if...)' is empty
    if row['Required Field?'] == 'y' and pd.isna(row['Branching Logic (Show field only if...)']):
        custom_quality_rules.append('Cannot be empty')
    
    # Rule 2: If 'Required Field?' == 'y' and 'Branching Logic (Show field only if...)' is not empty
    if row['Required Field?'] == 'y' and pd.notna(row['Branching Logic (Show field only if...)']):
        custom_quality_rules.append('Cannot be empty if ' + row['Branching Logic (Show field only if...)'])

     # Rule 3: If 'Required Field?' is empty and 'Branching Logic (Show field only if...)' is not empty
    if row['Required Field?'] != 'y' and pd.notna(row['Branching Logic (Show field only if...)']):
        custom_quality_rules.append('If not empty, branching must be ' + row['Branching Logic (Show field only if...)'])
    
    # Rule 4: If 'Text validation type OR Show Slider Number' is 'date_dmy' and 'Field Type' is not 'birth_date'
    if row['Text Validation Type OR Show Slider Number'] == 'date_dmy' and row['Variable / Field Name'] != 'birth_date' and row['Variable / Field Name'] != 'death_date' and row['Required Field?'] == 'y':
        custom_quality_rules.append('Date range must be between [birth_date] and [death_date] or [episode_date]')
    
    if row['Text Validation Type OR Show Slider Number'] == 'date_dmy' and row['Variable / Field Name'] != 'birth_date' and row['Required Field?'] != 'y':
        custom_quality_rules.append('If not empty, date range must be between [birth_date] and [death_date] or [episode_date]')
    
    # Rule 4: If 'Field Type' is 'birth_date'
    if row['Variable / Field Name'] == 'birth_date':
        custom_quality_rules.append('Date range must be between [current_date] and [current_date] - 100 years')
    
    if row['Variable / Field Name'] == 'death_date':
        custom_quality_rules.append('Date range must be between [birth_date] and [current_date]')
    
    # Rule 7: If 'Text Validation Type OR Show Slider Number' is 'integer'
    if row['Text Validation Type OR Show Slider Number'] == 'integer' and row['Required Field?'] == 'y' and row['Variable / Field Name'] != 'number_of_pregnancies':
        custom_quality_rules.append('Year must be between years [birth_date] and [death_date] or [episode_date]')
    
    if row['Text Validation Type OR Show Slider Number'] == 'integer' and row['Required Field?'] != 'y' and row['Variable / Field Name'] != 'number_of_pregnancies':
        custom_quality_rules.append('If year not empty, year must be between years [birth_date] and [death_date] or [episode_date]')
    
    if 'Alert' in str(row['UPDATED RULE for DATA QUALITY']) and row['Required Field?'] == 'y':
        alert_rule = row['UPDATED RULE for DATA QUALITY']
        
        # Use regex to extract value range before 'Alert'
        # This captures "higher than X" and "lower than Y" before the 'Alert'
        value_higher_than = re.search(r'higher than (-?\d*\.?\d+)', alert_rule.split('Alert')[0].lower())
        value_lower_than = re.search(r'lower than (-?\d*\.?\d+)', alert_rule.split('Alert')[0].lower())
        
        # Check if the value range has both higher and lower limits
        if value_higher_than and value_lower_than:
            value_range = f"Value range should be higher than {value_higher_than.group(1)} and lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_higher_than and not value_lower_than:
            value_range = f"Value range should be higher than {value_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_lower_than and not value_higher_than:
            value_range = f"Value range should be lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        # Now handle alert thresholds after the 'Alert' keyword
        alert_part = alert_rule.split('Alert')[-1].lower()  # Extract everything after 'Alert'
        
        # Extract "lower than X" and "higher than Y" after 'Alert'
        alert_lower_than = re.search(r'lower than (-?\d*\.?\d+)', alert_part)
        alert_higher_than = re.search(r'higher than (-?\d*\.?\d+)', alert_part)
        
        # Append alert thresholds
        if alert_higher_than and alert_lower_than:
            value_range = f"Alert shown if lower than {alert_lower_than.group(1)} or higher than {alert_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif alert_higher_than and not value_lower_than:
            value_range = f"Alert shown if higher than {alert_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_lower_than and not value_higher_than:
            value_range = f"Alert shown if lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
    
    
    if 'alert' in str(row['UPDATED RULE for DATA QUALITY']) and row['Required Field?'] != 'y':
        alert_rule = row['UPDATED RULE for DATA QUALITY']
        
        # Use regex to extract value range before 'Alert'
        # This captures "higher than X" and "lower than Y" before the 'Alert'
        value_higher_than = re.search(r'higher than (-?\d*\.?\d+)', alert_rule.split('alert')[0].lower())
        value_lower_than = re.search(r'lower than (-?\d*\.?\d+)', alert_rule.split('alert')[0].lower())
        
        # Check if the value range has both higher and lower limits
        if value_higher_than and value_lower_than:
            value_range = f"If not empty, value range should be higher than {value_higher_than.group(1)} and lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_higher_than and not value_lower_than:
            value_range = f"If not empty, value range should be higher than {value_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_lower_than and not value_higher_than:
            value_range = f"If not empty, value range should be lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        # Now handle alert thresholds after the 'Alert' keyword
        alert_part = alert_rule.split('alert')[-1].lower()  # Extract everything after 'Alert'
        
        # Extract "lower than X" and "higher than Y" after 'Alert'
        alert_lower_than = re.search(r'lower than (-?\d*\.?\d+)', alert_part)
        alert_higher_than = re.search(r'higher than (-?\d*\.?\d+)', alert_part)
        
        # Append alert thresholds
        if alert_higher_than and alert_lower_than:
            value_range = f"If not empty, alert shown if lower than {alert_lower_than.group(1)} or higher than {alert_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif alert_higher_than and not value_lower_than:
            value_range = f"If not empty, alert shown if higher than {alert_higher_than.group(1)}"
            custom_quality_rules.append(value_range)
        
        elif value_lower_than and not value_higher_than:
            value_range = f"If not empty, alert shown if lower than {value_lower_than.group(1)}"
            custom_quality_rules.append(value_range)
    
    
    # If there are custom quality rules to add
    if custom_quality_rules:
        # If 'Custom data quality' is not empty, append the new rules with '/'
        if pd.notna(row['Custom data quality']) and row['Custom data quality'] != '':
            df_merged.at[index, 'Custom data quality'] = row['Custom data quality'] + ' / ' + ' / '.join(custom_quality_rules)
        else:
            # Otherwise, just add the new rules
            df_merged.at[index, 'Custom data quality'] = ' / '.join(custom_quality_rules)

# Save the final dataframe to an Excel file
df_merged.to_excel('genomed_data_quality_check.xlsx', index=False)
