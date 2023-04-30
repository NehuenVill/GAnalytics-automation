from get_historic_data import get_historics_nv
import threading
from datetime import date



start_date = date(day=1,month=1,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/report-home/a6897643w13263073p13942924'
tvd_file = 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/channel_traffic_prensa'


if __name__ == '__main__':

    get_historics_nv(start_date, end_date, url, tvd_file, 'Acquisition', 'All Traffic', 'Channels')

def thread_function(start_date: date, end_date: date, url: str, op_file: str, *btns: Any):
    


if __name__ == "__main__":

    x = threading.Thread(target=thread_function, args=(1,))
    x.start()