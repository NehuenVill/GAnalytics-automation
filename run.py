from automation import get_metrics_data

tu_email = 'nehuenv620@gmail.com'
tu_contraseña = 'vzkfvhnhhiooiiuk'
rango_de_dias = 'semanal'  # Puede ser: 'semanal', 'mensual' o una cantidad de días, como 1, 2, 15, 20, etc.
#ejecución del programa:

get_metrics_data(rango_de_dias, tu_email, tu_contraseña)