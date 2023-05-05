from shutil import ExecError
from time import sleep
from get_historic_data import get_historics_faster, save_historic_to_excel
import threading
from datetime import date



start_date = date(day=1,month=1,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/a6899125w13265737p13945804'

file_number = {'Channel' : 1,
            'Source' : 60,
            'Page' : 1,
            'Device' : 0,
            'Segmetation': 1,
            }

op_files_prensa = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/prensa_historico/channel_traffic_prensa',
            'Source' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/prensa_historico/source_medium_traffic_prensa',
            'Page' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/prensa_historico/page_traffic_prensa',
            'Device' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/prensa_historico/device_traffic_prensa',
            'Segmetation': '',
            }

op_files_martes = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/channel_traffic_midiario',
            'Source' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/source_medium_traffic_midiario',
            'Page' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/page_traffic_midiario',
            'Device' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/device_traffic_midiario',
            'Segmetation': '',
            }

op_files_midiario = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/channel_traffic_midiario',
            'Source' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/source_medium_traffic_midiario',
            'Page' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/page_traffic_midiario',
            'Device' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/midiario_historico/device_traffic_midiario',
            'Segmetation': '',
            }

op_files_ellas = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/channel_traffic_ellas',
            'Source' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/source_medium_traffic_ellas',
            'Page' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/page_traffic_ellas',
            'Device' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/device_traffic_ellas',
            'Segmetation': '',
            }

btns_sequence = {'Channel' : ('Acquisition', 'All Traffic', 'Channels'),
                'Source' : ('Acquisition', 'All Traffic', 'Source/Medium'),
                'Page' : ('Behavior', 'Site Content', 'All Pages'),
                'Device' : ('Audience', 'Mobile', 'Overview'),
                'Segmetation': (),
                }

def save_all(*metrics):
    

    for metric in metrics:

        try:

            save_historic_to_excel(op_files_prensa[metric])

        except Exception as err:

            print(f'\nThe error was: {err.args}\n')


def run(*metrics):

    for metric in metrics:

        btns = btns_sequence[metric]

        get_historics_faster(start_date, end_date, url, op_files_midiario[metric], file_number[metric], btns[0], btns[1], btns[2])

if __name__ == "__main__":

    #run('Page')

    save_all('Channel', 'Device', 'Source')