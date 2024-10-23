import pandas as pd
from config import project_name

# read files radgenint_custom_rules_comments.xlsx and radgenint_custom_rules.xlsx
df_modified = pd.read_excel(f'{project_name}/{project_name}_custom_rules_iecs_comments.xlsx')

# iecs_comments.xlsx
df_iecs = pd.read_excel('iecs_comments.xlsx')

df_modified['changed_data_quality_rule'] = ''
df_modified['changed_date'] = ''
df_modified['changed_value'] = ''
df_modified['changed_value_qr'] = ''

# check variables where Custom data quality has changed from df_modified and df_iecs
for index, row in df_modified.iterrows():
    variable_name = row['Variable / Field Name']
    custom_data_quality = row['Custom data quality']
    custom_data_quality_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['Custom data quality'].values
    # get first element of the list
    custom_data_quality_iecs = custom_data_quality_iecs[0] if len(custom_data_quality_iecs) > 0 else ''

    custom_data_quality_changed = not (pd.isna(custom_data_quality) and pd.isna(custom_data_quality_iecs)) and (custom_data_quality != custom_data_quality_iecs)

    if custom_data_quality_changed:
        df_modified.at[index, 'changed_data_quality_rule'] = 'y'

# check variables where min date and max date has changed
for index, row in df_modified.iterrows():
    variable_name = row['Variable / Field Name']
    min_date = row['min date']
    max_date = row['max date']
    min_date_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['min date'].values
    max_date_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['max date'].values
    # get first element of the list
    min_date_iecs = min_date_iecs[0] if len(min_date_iecs) > 0 else ''
    max_date_iecs = max_date_iecs[0] if len(max_date_iecs) > 0 else ''

    # Compare both NaN and non-NaN values
    min_date_changed = not (pd.isna(min_date) and pd.isna(min_date_iecs)) and (min_date != min_date_iecs)
    max_date_changed = not (pd.isna(max_date) and pd.isna(max_date_iecs)) and (max_date != max_date_iecs)
    
    # If either date has changed
    if min_date_changed or max_date_changed:
        df_modified.at[index, 'changed_date'] = 'y'

# check variables where min value and max value has changed
for index, row in df_modified.iterrows():
    variable_name = row['Variable / Field Name']
    min_value = row['min value']
    max_value = row['max value']
    min_value_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['min value'].values
    max_value_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['max value'].values
    # get first element of the list
    min_value_iecs = min_value_iecs[0] if len(min_value_iecs) > 0 else ''
    max_value_iecs = max_value_iecs[0] if len(max_value_iecs) > 0 else ''

    # Compare both NaN and non-NaN values
    min_value_changed = not (pd.isna(min_value) and pd.isna(min_value_iecs)) and (min_value != min_value_iecs)
    max_value_changed = not (pd.isna(max_value) and pd.isna(max_value_iecs)) and (max_value != max_value_iecs)

    if min_value_changed or max_value_changed:
        df_modified.at[index, 'changed_value'] = 'y'

# check variables where min alert and max alert from df_iecs has changed compared to df_modified min value qr and max value qr
for index, row in df_modified.iterrows():
    variable_name = row['Variable / Field Name']
    min_value_qr = row['min value qr']
    max_value_qr = row['max value qr']
    min_value_qr_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['min alert'].values
    max_value_qr_iecs = df_iecs[df_iecs['Variable / Field Name'] == variable_name]['max alert'].values
    # get first element of the list
    min_value_qr_iecs = min_value_qr_iecs[0] if len(min_value_qr_iecs) > 0 else ''
    max_value_qr_iecs = max_value_qr_iecs[0] if len(max_value_qr_iecs) > 0 else ''

    min_value_qr_changed = not (pd.isna(min_value_qr) and pd.isna(min_value_qr_iecs)) and (min_value_qr != min_value_qr_iecs)
    max_value_qr_changed = not (pd.isna(max_value_qr) and pd.isna(max_value_qr_iecs)) and (max_value_qr != max_value_qr_iecs)

    if min_value_qr_changed or max_value_qr_changed:
        df_modified.at[index, 'changed_value_qr'] = 'y'

# save the modified dataframe to an Excel file
df_modified.to_excel(f'{project_name}/{project_name}_custom_rules_modified_columns.xlsx', index=False)


