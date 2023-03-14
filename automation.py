from os import environ, remove
import ssl
from tkinter import E
import pandas as pd
import smtplib
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (DateRange,
                                                Metric,
                                                Dimension,
                                                RunReportRequest,
                                                FilterExpression,
                                                Filter
                                                )
from google.api_core.exceptions import InvalidArgument
import openpyxl
import sqlite3
from sqlite3 import IntegrityError
from time import sleep
 
property_id = '347990037'
environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Proyecto API-8b9749be5c38.json'

client = BetaAnalyticsDataClient()

metrics_list = ['Date',
                'activeUsers',
                'averageSessionDuration',
                'bounceRate',
                'crashFreeUsersRate',
                'dauPerMau',
                'dauPerWau',
                'engagedSessions',
                'engagementRate',
                'eventCount',
                'eventCountPerUser',
                'eventsPerSession',
                'newUsers',
                'screenPageViews',
                'screenPageViewsPerSession',
                'sessions',
                'sessionsPerUser',
                'totalUsers',
                'userEngagementDuration',
                'wauPerMau',
                'ReturningUsers',
                'EvolokEvents'
                ]


# Change insert_to_db function so that it doesn't depend on a fixed amount
# of metrics_list items.

def insert_to_db(f_data):

    print(f'Inserting data: {f_data} into the database')
    
    connection = sqlite3.connect("Prensa_base_de_datos.db")
    cursor = connection.cursor()

    try:

        cursor.execute(f"""INSERT INTO Analytics_data VALUES ('{str(f_data['Date'])}',{f_data['activeUsers']},{f_data['averageSessionDuration']},{f_data['bounceRate']},
        {f_data['crashFreeUsersRate']},{f_data['dauPerMau']},{f_data['dauPerWau']},{f_data['engagedSessions']},
        {f_data['engagementRate']}, {f_data['eventCount']}, {f_data['eventCountPerUser']}, {f_data['eventsPerSession']}, {f_data['newUsers']},
        {f_data['screenPageViews']}, {f_data['screenPageViewsPerSession']}, {f_data['sessions']}, {f_data['sessionsPerUser']},
        {f_data['totalUsers']}, {f_data['userEngagementDuration']}, {f_data['wauPerMau']})""")

    except IntegrityError:

        print('Already exists')

    connection.commit()
    connection.close()

def set_dates(date_range):

    end_date = date.today()

    if date_range == 'mensual':

        start_date = end_date - relativedelta(months=+1)

    elif date_range == 'semanal':

        start_date = end_date - relativedelta(weeks=+1)

    elif type(date_range) == int:

        start_date = end_date - relativedelta(days=+date_range)

    else:

        print('Wrong date range!')

    end_date = end_date.strftime('%Y-%m-%d')

    start_date = start_date.strftime('%Y-%m-%d')

    return [end_date, start_date]

# Change get_data_manual function so that it doesn't depend on a fixed amount
# of metrics_list items.

# Change the get_metrics_data to use the get data manual with fixed dates.

def get_data_manual(start_date, end_date, email, password):

    start = date(year=int(start_date.split('-')[0]), month=int(start_date.split('-')[1]),day=int(start_date.split('-')[2]))

    end = date(year=int(end_date.split('-')[0]),month=int(end_date.split('-')[1]),day=int(end_date.split('-')[2]))

    data = []

    day_num = 0

    while True:

        data.append({'Date': start.strftime('%m/%d/%Y').replace('2023', '23').replace('2024', '24')})

        print(f"Getting info from date: {data[day_num]['Date']}")

        st_date_range = start.strftime('%Y-%m-%d')

        for j in range(3,4):

            print('-'*100)

            print(f'\nGetting metrics')

            try:

                if j == 0:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[1]), Metric(name=metrics_list[2]), Metric(name=metrics_list[3]), Metric(name=metrics_list[4]), 
                        Metric(name=metrics_list[5]), Metric(name=metrics_list[6]), Metric(name=metrics_list[7]), Metric(name=metrics_list[8]), Metric(name=metrics_list[9]), 
                        Metric(name=metrics_list[10]) ],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)],
                    )
                

                    full_response = client.run_report(request)

                    try:

                        for n, metric in enumerate(full_response.rows[0].metric_values):

                            data[day_num][full_response.metric_headers[n].name] = metric.value

                            print(f'Extracted metric: {full_response.metric_headers[n].name} ---- value = {data[day_num][full_response.metric_headers[n].name]}')

                    except IndexError:

                        print(f'metric: {metric} not available.')

                        data[day_num][metric] = 'Not available'

                elif j == 1:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[11]), Metric(name=metrics_list[12]), Metric(name=metrics_list[13]), Metric(name=metrics_list[14]), 
                        Metric(name=metrics_list[15]), Metric(name=metrics_list[16]), Metric(name=metrics_list[17]), Metric(name=metrics_list[18]), Metric(name=metrics_list[19])],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)]
                    )

                    full_response = client.run_report(request)

                    try:

                        for n, metric in enumerate(full_response.rows[0].metric_values):

                            data[day_num][full_response.metric_headers[n].name] = metric.value

                            print(f'Extracted metric: {full_response.metric_headers[n].name} ---- value = {data[day_num][full_response.metric_headers[n].name]}')

                    except IndexError:

                        print(f'metric: {metric} not available.')

                        data[day_num][metric] = 'Not available'

                elif j == 2:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[17])],
                        dimensions=[Dimension(name='newVsReturning')],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)]
                    )

                    full_response = client.run_report(request)

                    try:

                        new = full_response.rows[0].metric_values[0].value

                        returning = full_response.rows[1].metric_values[0].value

                        data[day_num]['NewUsers'] = new

                        data[day_num]['ReturningUsers'] = returning

                        print(f"Extracted metric: NewUsers ---- value = {data[day_num]['NewUsers']}")
                        print(f"Extracted metric: ReturningUsers ---- value = {data[day_num]['ReturningUsers']}")

                    except IndexError:

                        print(f'metric: {metric} not available.')

                        data[day_num][metric] = 'Not available'

                elif j == 3:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[9])],
                        dimensions=[Dimension(name='customEvent:')],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)],
                    )
                    full_response = client.run_report(request)

                    print(full_response)

                    try:

                        evolok = full_response.rows[7].metric_values[0].value

                        data[day_num]['EvolokEvents'] = evolok

                        print(f"Extracted metric: EvolokEvents ---- value = {data[day_num]['EvolokEvents']}")

                    except IndexError:

                        print(f'metric: {metric} not available.')

                        data[day_num][metric] = 'Not available'

            except Exception as e:

                print(e)

                print('Could not get info from metrics metrics requested')

        start = start + timedelta(days=1)

        day_num += 1

        if start > end:

            break

    save_to_excel(data, start_date, end_date)

    send_mail(email, password, start_date, end_date)

def get_metrics_data(d_range, email, password) -> dict:

    """Function to get data from the GA API.
        
    """

    date_range = set_dates(d_range)

    end_date = date_range[0]
    start_date = date_range[1]

    today = date.today()

    data = [{}, {}, {}, {}, {}, {}, {}]

    for i in range(7,0,-1):

        data[i-1]['Date'] = (today - timedelta(days=i+1)).strftime('%m/%d/%Y').replace('2023', '23').replace('2024', '24')

        print(f"Getting info from date: {data[i-1]['Date']}")

        for j in range(0,2):

            print('-'*100)

            print(f'\nGetting metrics')

            try:

                if j == 0:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[1]), Metric(name=metrics_list[2]), Metric(name=metrics_list[3]), Metric(name=metrics_list[4]), 
                        Metric(name=metrics_list[5]), Metric(name=metrics_list[6]), Metric(name=metrics_list[7]), Metric(name=metrics_list[8]), Metric(name=metrics_list[9]), 
                        Metric(name=metrics_list[10]) ],
                        date_ranges=[DateRange(start_date=f'{i+1}daysAgo', end_date=f'{i+1}daysAgo')]
                    )

                    full_response = client.run_report(request)

                elif j == 1:

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metrics_list[11]), Metric(name=metrics_list[12]), Metric(name=metrics_list[13]), Metric(name=metrics_list[14]), 
                        Metric(name=metrics_list[15]), Metric(name=metrics_list[16]), Metric(name=metrics_list[17]), Metric(name=metrics_list[18]), Metric(name=metrics_list[19])],
                        date_ranges=[DateRange(start_date=f'{i+1}daysAgo', end_date=f'{i+1}daysAgo')]
                    )

                    full_response = client.run_report(request)

            except InvalidArgument:

                print('Could not get info from metrics metrics requested')

            try:

                for n, metric in enumerate(full_response.rows[0].metric_values):

                    data[i-1][full_response.metric_headers[n].name] = metric.value

                    print(f'Extracted metric: {full_response.metric_headers[n].name} ---- value = {data[i-1][full_response.metric_headers[n].name]}')

            except IndexError:

                print(f'metric: {metric} not available.')

                sleep(50)

                data[i-1][metric] = 'Not available'

    data.reverse()

    save_to_excel(data, start_date, end_date)

    for row in data:

        insert_to_db(row)

    send_mail(email, password, start_date, end_date)


def save_to_excel(metrics_data, start_date, end_date):

    print('Saving to excel...')

    df = pd.DataFrame(metrics_data, index =pd.RangeIndex(0,len(metrics_data)) ,columns = metrics_list)
    
    df.to_excel(f"Metricas_desde_{start_date}_a_{end_date}.xlsx", index=True, columns=metrics_list)


def send_mail(email, password, start_date, end_date):

    print('Sending email...')

    excel_file = f'Metricas_desde_{start_date}_a_{end_date}.xlsx'
    
    mail = MIMEMultipart()
    mail['From'] = email
    mail['To'] = email
    mail['Subject'] = f"Reporte desde {start_date} a {end_date}"

    with open(excel_file, "rb") as f:

        part = MIMEApplication(
            f.read(),
            Name=basename(excel_file)
        )

    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(excel_file)
    mail.attach(part)

    port = 465

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        
        server.login(email, password)
        
        server.sendmail(email, email, mail.as_string())

    remove(basename(excel_file))

    print('email sent successfully.')

def export_hist_data_to_db():

    workbook = openpyxl.load_workbook('Datos_históricos_Google_Analytics.xlsx', data_only=True)

    worksheet = workbook['Hoja1']

    for row in worksheet.iter_rows(min_row=2, max_row=1418):

        data = {}

        for i, metric in enumerate(metrics_list):

            if row[i].value == 'Not available':

                data[metric] = 'Null'

            elif '-' in str(row[i].value):

                data[metric] = str(row[i].value).replace('-', '/').replace(' 00:00:00', '').replace('2010','10').replace('2011','11').replace('2012','12')

            elif ':' in str(row[i].value):

                print(str(row[i].value))

                minutes = int(str(row[i].value).split(':')[0])
                seconds = int(str(row[i].value).split(':')[1])

                final_time = float(minutes * 60 + seconds)

                data[metric] = final_time

            else:

                data[metric] = row[i].value

        print('-'*100)

        print(f'Exporting record: {data}')

        print('-'*100)

        insert_to_db(data)

if __name__ == '__main__':

    pass