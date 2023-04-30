from get_historic_data import get_traffic_by_device, get_traffic_by_page
from datetime import date

start_date = date(day=4,month=3,year=2009)
end_date = date(day=24,month=4,year=2023)
url = 'https://analytics.google.com/analytics/web/?authuser=2#/report-home/a6897643w13263073p212900937'
tvd_file = 'C:/Users/nehue/Documents/programas_de_python/Upwork_tasks/Google_Analytics_automation/Prensa_historico/device_traffic_midiario.json'

if __name__ == "__main__":

    get_traffic_by_device(start_date, end_date, url, tvd_file)
