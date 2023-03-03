from automation import get_metrics_data, get_data_manual

tu_email = 'nehuenv620@gmail.com'
tu_contraseña = 'vzkfvhnhhiooiiuk'
rango_de_dias = 'semanal'

#ejecución del programa:

#get_metrics_data(rango_de_dias, tu_email, tu_contraseña)

#Ejecución de la función manual:

Día_incial = '2023-01-10' #AAAA-MM-DD

Día_final = '2023-01-15'

get_data_manual(Día_incial, Día_final, tu_email, tu_contraseña)