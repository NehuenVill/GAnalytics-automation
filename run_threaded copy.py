from time import sleep
from get_historic_data import get_historics_nv
import threading
from datetime import date



start_date = date(day=1,month=1,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/a6898250w13264172p13944111'

file_number = {'Channel' : 1,
            'Source' : 1,
            'Page' : 1,
            'Device' : 1,
            'Segmetation': 1,
            }

op_files = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/ellas_historico/channel_traffic_ellas',
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


def run(*metrics):

    for metric in metrics:

        btns = btns_sequence[metric]

        x = threading.Thread(target=get_historics_nv, args=(start_date, end_date, url, op_files[metric], file_number[metric], btns[0], btns[1], btns[2]))
        x.start()

        sleep(60)

if __name__ == "__main__":

    run('Device', 'Channel', 'Source')