from automation import get_metrics_data, get_data_manual, update_db_columns

tu_email = 'nehuenv620@gmail.com'
tu_contraseña = 'wrivyyyajqwjttrt'

#ejecución del programa:

get_metrics_data(tu_email, tu_contraseña)

#Ejecución de la función manual:

Día_incial = '2023-01-10' #AAAA-MM-DD

Día_final = '2023-01-15' 

#get_data_manual(Día_incial, Día_final, tu_email, tu_contraseña)



# Para agregar nuevas métricas:

nueva_métrica = "EvolokEvents"              #Nombre de la métrica exactamente como aparace en la página de la API de GA.

tipo = "numero"                       # "texto" o "numero"

#update_db_columns(new_column= nueva_métrica, type = tipo)