import pandas as pd
import re

# Load the data
df_data_quality_rules = pd.read_excel('integra_data_quality_check.xlsx')

# Create a new DataFrame for custom rules
df_custom_rules = pd.DataFrame(columns=['rule_name', 'rule_logic', 'real_time_execution'])

# Iterate through the rows of df_data_quality_rules
for index, row in df_data_quality_rules.iterrows():
    custom_data_quality = row['Custom data quality']
    
    # 'Cannot be empty'
    if pd.notna(custom_data_quality) and 'Cannot be empty' in custom_data_quality and 'Cannot be empty if' not in custom_data_quality:
        variable_field_name = row['Variable / Field Name']
        field_label = row['Field Label']
        form_name = row['Form Name']

        # Construct the rule_name and rule_logic
        rule_name = f"[{variable_field_name}] ({field_label}) should have a value but is missing."
        rule_logic = f"[{variable_field_name}] = '' and [{form_name}_complete] = '2'"
        real_time_execution = 'y'

        # Append the new rule to df_custom_rules
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': real_time_execution
        }, ignore_index=True)
    
    # 'Cannot be empty if'
    if pd.notna(custom_data_quality) and 'Cannot be empty if' in custom_data_quality:
        variable_field_name = row['Variable / Field Name']
        field_label = row['Field Label']
        form_name = row['Form Name']
        
        # Extract the condition after "if"
        condition = re.search(r'Cannot be empty if (.*?)(?: /|$)', custom_data_quality).group(1).strip()
        
        # Construct the rule_logic
        rule_logic = f"[{variable_field_name}] = '' and {condition} and [{form_name}_complete] = '2'"
        
        # Construct the rule_name for this case
        rule_name = f"[{variable_field_name}] ({field_label}) should have a value but is missing based on condition."
        
        # Append the new rule to df_custom_rules
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    
    # 'If not empty, alert shown if'
    if pd.notna(custom_data_quality) and 'If not empty, alert shown if' in custom_data_quality:
        lower_match = re.search(r'lower than (\d+)', custom_data_quality)
        higher_match = re.search(r'higher than (\d+)', custom_data_quality)
        rule_name = f"[{variable_field_name}] ({field_label}) is out of expected range. Please check the values."
        # 1. Handle "If not empty, alert shown if" (both lower and higher)
        if lower_match and higher_match:
            lower_value = lower_match.group(1)
            higher_value = higher_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] < '{lower_value}' or [{variable_field_name}] > '{higher_value}') and [{form_name}_complete] = '2'"
            
            # Append the new rule to df_custom_rules
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)

        # 2. Handle "If not empty, alert shown if" (only lower)
        elif lower_match:
            lower_value = lower_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and [{variable_field_name}] < '{lower_value}' and [{form_name}_complete] = '2'"
            
            # Append the new rule
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)

        # 3. Handle "If not empty, alert shown if" (only higher)
        elif higher_match:
            higher_value = higher_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and [{variable_field_name}] > '{higher_value}' and [{form_name}_complete] = '2'"
            
            # Append the new rule
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)

    # 'Alert if value'
    if pd.notna(custom_data_quality) and 'Alert if value' in custom_data_quality:
        lower_match = re.search(r'lower than (\d+)', custom_data_quality)
        higher_match = re.search(r'higher than (\d+)', custom_data_quality)
        rule_name = f"[{variable_field_name}] ({field_label}) is out of expected range. Please check the values."
        # 1. Handle "If not empty, alert shown if" (both lower and higher)
        if lower_match and higher_match:
            lower_value = lower_match.group(1)
            higher_value = higher_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] < '{lower_value}' or [{variable_field_name}] > '{higher_value}') and [{form_name}_complete] = '2'"
            
            # Append the new rule to df_custom_rules
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)

        # 2. Handle "If not empty, alert shown if" (only lower)
        elif lower_match:
            lower_value = lower_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and [{variable_field_name}] < '{lower_value}' and [{form_name}_complete] = '2'"
            
            # Append the new rule
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)

        # 3. Handle "If not empty, alert shown if" (only higher)
        elif higher_match:
            higher_value = higher_match.group(1)
            rule_logic = f"[{variable_field_name}] <> '' and [{variable_field_name}] > '{higher_value}' and [{form_name}_complete] = '2'"
            
            # Append the new rule
            df_custom_rules = df_custom_rules._append({
                'rule_name': rule_name,
                'rule_logic': rule_logic,
                'real_time_execution': 'y'
            }, ignore_index=True)
    
    # If not empty, date range
    if pd.notna(custom_data_quality) and 'If not empty, date range must be between' in custom_data_quality and row['Variable / Field Name'] != 'date_of_birth' and row['Variable / Field Name'] != 'death_date':
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date] or [episode_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

    # Date range
    if pd.notna(custom_data_quality) and 'If not empty, date range must be between' in custom_data_quality and row['Variable / Field Name'] != 'date_of_birth' and row['Variable / Field Name'] != 'death_date':
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date] or [episode_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

    # date_of_birth -> not sure this will work, do it by hand
    if pd.notna(custom_data_quality) and 'If not empty, date range must be between' in custom_data_quality and row['Variable / Field Name'] == 'date_of_birth':
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] > [death_date] or [episode_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    
    # death_date -> not sure this will work, do it by hand
    if pd.notna(custom_data_quality) and 'If not empty, date range must be between' in custom_data_quality and row['Variable / Field Name'] == 'death_date':
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] < [date_of_birth] or [episode_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    
    # year
    # do it by hand


    

# Save the new DataFrame to a new Excel file
df_custom_rules.to_excel('custom_data_quality_rules.xlsx', index=False)
    
