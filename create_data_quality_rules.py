import pandas as pd
import re

# Load the data
df_data_quality_rules = pd.read_excel('radeep_data_quality_check.xlsx')

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
    
    
    # DATE QR creation

    # 1. Dob
    if pd.notna(custom_data_quality) and row['Variable / Field Name'] == 'date_of_birth':
        # alive
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] <> '2' ([{variable_field_name}] > 'today' or [{variable_field_name}] < '1924-01-01') and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # dead
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] = '2' ([{variable_field_name}] > [death_date] or [{variable_field_name}] < '1924-01-01') and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    

    # 2. Death date
    elif pd.notna(custom_data_quality) and row['Variable / Field Name'] == 'death_date':
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and ([{variable_field_name}] > 'today' or [{variable_field_name}] < [date_of_birth]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    
    # 3. Year immmigration
    elif pd.notna(custom_data_quality) and row['Variable / Field Name'] == 'year_immigration':
        # Alive
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] <> '2' and ([{variable_field_name}] < year([date_of_birth]) or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] = '2' and ([{variable_field_name}] < year([date_of_birth]) or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    
    # 4. Date range must be between [birth_date] and [death_date] or [episode_date]
    elif pd.notna(custom_data_quality) and 'Date range must be between [birth_date] and [death_date] or [episode_date]' in custom_data_quality:
        # Alive
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] <> '2' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [patient_status] = '2' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

    # 5. Dates in repeated events
    elif pd.notna(custom_data_quality) and "If [current-instance] = '1', date range must be between [birth_date] and [death_date] or [episode_date]. Else, date range must be between [episode_date] - 1 year and [death_date] or [episode_date]" in custom_data_quality:
        
        # First instance
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        # Alive
        rule_logic = f"[{variable_field_name}] <> '' and [current-instance] = '1' and [patient_status] <> '2' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
        # Dead
        rule_logic = f"[{variable_field_name}] <> '' and [current-instance] = '1' and [patient_status] = '2' and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)


        # Other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        # Alive
        rule_logic = f"[{variable_field_name}] <> '' and [current-instance] > '1' and [patient_status] <> '2' and ([{variable_field_name}] < [timestamp] - 1 year or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
        # Dead
        rule_logic = f"[{variable_field_name}] <> '' and [current-instance] > '1' and [patient_status] = '2' and ([{variable_field_name}] < [death_date] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

    # 6. date_acute_r1 to r4
    elif pd.notna(custom_data_quality) and 'date_acute_r' in row['Variable / Field Name']:
        number = re.search(r'\d+', row['Variable / Field Name']).group()

        # 1. for variable acute_id_rxx
        variable_of_branching_associated = f'acute_id_r{number}'
        # Alive and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] = '9' or [{variable_of_branching_associated}] = '18' or [{variable_of_branching_associated}] = '20' or [{variable_of_branching_associated}] = '22') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] = '9' or [{variable_of_branching_associated}] = '18' or [{variable_of_branching_associated}] = '20' or [{variable_of_branching_associated}] = '22') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] <> '2' and ([{variable_field_name}] < [timestamp] - 1 year or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] = '2' and ([{variable_field_name}] < [death_date] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] <> '9' and [{variable_of_branching_associated}] <> '18' and [{variable_of_branching_associated}] <> '20' and [{variable_of_branching_associated}] <> '22') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] <> '9' and [{variable_of_branching_associated}] <> '18' and [{variable_of_branching_associated}] <> '20' and [{variable_of_branching_associated}] <> '22') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # 2. for variable acute_id2_rxx
        variable_of_branching_associated = f'acute_id2_r{number}'
        # Alive and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] = '9' or [{variable_of_branching_associated}] = '18' or [{variable_of_branching_associated}] = '20' or [{variable_of_branching_associated}] = '22') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] = '9' or [{variable_of_branching_associated}] = '18' or [{variable_of_branching_associated}] = '20' or [{variable_of_branching_associated}] = '22') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] <> '2' and ([{variable_field_name}] < [timestamp] - 1 year or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] = '2' and ([{variable_field_name}] < [death_date] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] <> '9' and [{variable_of_branching_associated}] <> '18' and [{variable_of_branching_associated}] <> '20' and [{variable_of_branching_associated}] <> '22') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] <> '9' and [{variable_of_branching_associated}] <> '18' and [{variable_of_branching_associated}] <> '20' and [{variable_of_branching_associated}] <> '22') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)
    

    # 7. date_acute_scd_r1 to r16
    elif pd.notna(custom_data_quality) and 'date_acute_scd_r' in row['Variable / Field Name']:
        number = re.search(r'\d+', row['Variable / Field Name']).group()

        # 1. for variable acute_id_scd_rxx
        variable_of_branching_associated = f'acute_id_scd_r{number}'
        # Alive and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] = '25' or [{variable_of_branching_associated}] = '27' or [{variable_of_branching_associated}] = '35' or [{variable_of_branching_associated}] = '42' or [{variable_of_branching_associated}] = '44' or [{variable_of_branching_associated}] = '45' or [{variable_of_branching_associated}] = '46') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] = '25' or [{variable_of_branching_associated}] = '27' or [{variable_of_branching_associated}] = '35' or [{variable_of_branching_associated}] = '42' or [{variable_of_branching_associated}] = '44' or [{variable_of_branching_associated}] = '45' or [{variable_of_branching_associated}] = '46') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] <> '2' and ([{variable_field_name}] < [timestamp] - 1 year or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] = '2' and ([{variable_field_name}] < [death_date] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] <> '25' and [{variable_of_branching_associated}] <> '27' and [{variable_of_branching_associated}] <> '35' and [{variable_of_branching_associated}] <> '42' and [{variable_of_branching_associated}] <> '44' and [{variable_of_branching_associated}] <> '45' and [{variable_of_branching_associated}] <> '46') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] <> '25' and [{variable_of_branching_associated}] <> '27' and [{variable_of_branching_associated}] <> '35' and [{variable_of_branching_associated}] <> '42' and [{variable_of_branching_associated}] <> '44' and [{variable_of_branching_associated}] <> '45' and [{variable_of_branching_associated}] <> '46') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # 2. for variable acute_id2_scd_rxx
        variable_of_branching_associated = f'acute_id2_scd_r{number}'
        # Alive and first instance and selected conditions
        # Alive and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] = '25' or [{variable_of_branching_associated}] = '27' or [{variable_of_branching_associated}] = '35' or [{variable_of_branching_associated}] = '42' or [{variable_of_branching_associated}] = '44' or [{variable_of_branching_associated}] = '45' or [{variable_of_branching_associated}] = '46') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] = '25' or [{variable_of_branching_associated}] = '27' or [{variable_of_branching_associated}] = '35' or [{variable_of_branching_associated}] = '42' or [{variable_of_branching_associated}] = '44' or [{variable_of_branching_associated}] = '45' or [{variable_of_branching_associated}] = '46') and ([{variable_field_name}] < [timestamp] - 2 years or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] <> '2' and ([{variable_field_name}] < [timestamp] - 1 year or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and other instances
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] > 1 and [patient_status] = '2' and ([{variable_field_name}] < [death_date] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Alive and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] <> '2' and ([{variable_of_branching_associated}] <> '25' and [{variable_of_branching_associated}] <> '27' and [{variable_of_branching_associated}] <> '35' and [{variable_of_branching_associated}] <> '42' and [{variable_of_branching_associated}] <> '44' and [{variable_of_branching_associated}] <> '45' and [{variable_of_branching_associated}] <> '46') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [timestamp]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)

        # Dead and first instance and other selected conditions
        rule_name = f"[{variable_field_name}] ({field_label}) must fall between the specified date range."
        rule_logic = f"[{variable_field_name}] <> '' and [current-event] = 1 and [patient_status] = '2' and ([{variable_of_branching_associated}] <> '25' and [{variable_of_branching_associated}] <> '27' and [{variable_of_branching_associated}] <> '35' and [{variable_of_branching_associated}] <> '42' and [{variable_of_branching_associated}] <> '44' and [{variable_of_branching_associated}] <> '45' and [{variable_of_branching_associated}] <> '46') and ([{variable_field_name}] < [date_of_birth] or [{variable_field_name}] > [death_date]) and [{form_name}_complete] = '2'"
        df_custom_rules = df_custom_rules._append({
            'rule_name': rule_name,
            'rule_logic': rule_logic,
            'real_time_execution': 'y'
        }, ignore_index=True)


# Save the new DataFrame to a new Excel file
df_custom_rules.to_excel('radeep_custom_data_quality_rules.xlsx', index=False)
    
