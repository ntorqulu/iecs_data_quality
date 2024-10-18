# Data quality rules generator

Code to automatically generate the data quality rules to set up the REDCap projects.

Message for Kalia and Stella (in process)

1. First, I rewrote the data quality rules of the CRF into a more standarized format in order to be able to extract regex patterns with code. If you look in the shared CRF, in the sheed Data Quality you will see two columns for each variable, named 'Rule for Data Quality' and 'Updated Rule for Data Quality'. The second one should contain the same information as the first one, but in a more structured manner.

2. I modified the data quality rules for the variables ‘date_acute_r1' and ‘date_acute_scd_r1' according to the comments in sheet Parameters in order to update the ranges for this dates depending on the follow-up instance and the selected options in acute_complications and acute_complications_scd.

3. The format for the data quality rules in the crf is as follows:
	- If the variable is a required field: Cannot be empty
	- If the variable is a required field with associated branching logic: Cannot be empty if [variable description] is [value]
	- If the variable is an optional field with associated branching logic: If not empty, [variable description] should be [value]
	- If the variable is a required date: Date range should be higher than [variable description] and lower than [variable description]
	- If the variable is an optional date: If not empty, date range should be higher than [variable description] and lower than [variable description]
	- If the variable is a required field with range values: Value range should be higher than [value] and lower than [value]
	- If the variable is an optional field with range values: If not empty, value range should be higher than [value] and lower than [value]
	- If the variable is a required field with alert values: Alert shown if value is lower than [value] or higher than [value]
	- If the variable is an optional field with alert values: If not empty, alert shown if value is lower than [value] or higher than [value]
	- date_acute_r1 follows a different format
	- date_acute_scd_r1 follows a different format

4. Now the CRF is missing the data quality rules for variables that are dates and are located in forms that can be repeated. I’ll add more data quality rules for this dates so that:
	- If the variable is a required date in a repeated event: if baseline registration, date range should be . If follow-up registration, date range should be
	-  If the variable is an optional date in a repeated event: If not empty and baseline registration, date range should be. If not empty and follow-up registration, date range should be 
This date changes are updated in the provided .xlsx files but not in the CRF, I’ll do it next week.

5. Now regarding the files, radeep_custom_rules_iecs.xlsx has the following columns:
	- Variable / Field Name: name of the variable in REDCap. In this column, it appears all the variables from the REDCap project, not just the ones from the CRF.
	- Form Name: name of the from this variable belongs in REDCap
	- Field Label: description associated with the variable
	- Branching Logic: branching logic associated with the variable, used to construct the data quality rule after
	- Custom data quality: data quality rule written in machine readable format, but it is not code yet
