import pandas as pd
from config import project_name

# load file
print(project_name)
df_custom_rules = pd.read_excel(f'{project_name}/{project_name}_custom_rules_iecs.xlsx')
# load file iecs_comments.xlsx
df_comments = pd.read_excel('iecs_comments.xlsx')

# create column Comment in df_custom_rules
df_custom_rules['comments'] = ''
df_custom_rules['answers'] = ''

# list all the comments
comments = df_comments['comments'].unique()

# dictionary with answers for each comment
dict_answers = {
    'a la fecha de nacimiento hay que agregarle la referencia al evento / timestamp incluye h:m:s por lo que no se puede usar para validar dmy, se usó la variable timestamp_label que está en dmy': 
    'añadida referencia a evento en fecha de nacimiento / timestamp_label usado en lugar de timestamp',
    'es una calculada oculta, no tiene sentido min y max': 'eliminadas condiciones min y max',
    'No se puede aplicar la función year en las propiedades min o max. Se podría resolver con una variable calculada adicional y un campo campo descriptivo de alerta. Si se podría usar en reglas':
    'eliminadas funciones year en min y max, implementar solo como regla',
    'se agregó el cero como mínimo de alerta': 'no agregar cero como mínimo de alerta',
    'se copiaron los valores de value a alert. Indicar si no correspondía alert o si fue un error': 'no copiar valores de value a alert',
    'tenía los valores min=20 y max=3.4 aunque en la nota del campo indica [2.0, 3.4]. Como están vacías en esta planilla entonces quedan vacías.': 'correcto, valores vacíos',
    'tenía los valores min=209 y max=26.4 aunque en la nota del campo indica [20.9, 26.9]. Como están vacías en esta planilla entonces quedan vacías.': 'correcto, valores vacíos',
    'timestamp no existe en este formulario, se cambia por today': 'cambiado por today',
    'timestamp no existe en este formulario, se cambió por today': 'cambiado por today',
    'timestamp no existe en este formulario, se cambió por today / no se puden calcular menos cien años en la propiedad de min / se podría hacer en reglas o una alerta con un campo auxiliar de edad calculada': 'cambiado por today / no calcular menos cien años en min / hacer en reglas la edad calculada',
    'a la fecha de nacimiento hay que agregarle la referencia al evento / se usa timestamp_lab porque no se puede refeenciar a timestamp porque no se sabe a que instancia corresponde (en integra)':
    'añadida referencia a evento en fecha de nacimiento / timestamp_lab usado en lugar de timestamp'
}

# find in rows that match column Variable / Field Name in both files
for index, row in df_custom_rules.iterrows():
    variable_field_name = row['Variable / Field Name']
    comment = df_comments[df_comments['Variable / Field Name'] == variable_field_name]['comments']
    if not comment.empty and comment.values[0] in comments:
        df_custom_rules.at[index, 'comments'] = comment.values[0]
        # assign answer to the comment
        df_custom_rules.at[index, 'answers'] = dict_answers[comment.values[0]]
        print(dict_answers[comment.values[0]])

# save the file
df_custom_rules.to_excel(f'{project_name}/{project_name}_custom_rules_iecs_comments.xlsx', index=False)