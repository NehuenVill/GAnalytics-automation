from get_historic_data import get_historics_nv
import threading
from datetime import date



start_date = date(day=1,month=1,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/report-home/a6897643w13263073p13942924'

op_files = {'Channel' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/channel_traffic_prensa',
            'Source' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/source_medium_traffic_prensa',
            'Page' : 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/page_traffic_prensa',
            'Device' : ('Audience', 'Mobile', 'Overview'),
            'Segmetation': (),
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

        x = threading.Thread(target=get_historics_nv(), args=(start_date, end_date, url, tvd_file, btns[0], btns[1], btns[2],))
        x.start()


if __name__ == "__main__":

    run('Source', 'Page')