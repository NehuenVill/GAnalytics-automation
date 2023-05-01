

def create(path : str, start_number : int, end_number : int) -> None:

    for i in range(start_number, end_number + 1):

        with open(f'{path}({i}).json', 'x') as f:

            pass
        
if __name__ == '__main__':

    create('C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/source_medium_traffic_ellas',
            1, 10)
            