import math
from os import environ, remove
import ssl
from tkinter import E
from numpy import full
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

METRICS_LIST = [
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
                'screenPageViews',
                'screenPageViewsPerSession',
                'sessions',
                'sessionsPerUser',
                'totalUsers',
                'userEngagementDuration',
                'wauPerMau',
                ]

SPECIAL_METRICS = [                             #Metrics that require a dimension
                'totalUsers',
                'eventCount'
]


# EXCEL_HEADS = ['Date',
#                 'activeUsers',
#                 'averageSessionDuration',
#                 'bounceRate',
#                 'crashFreeUsersRate',
#                 'dauPerMau',
#                 'dauPerWau',
#                 'engagedSessions',
#                 'engagementRate',
#                 'eventCount',
#                 'eventCountPerUser',
#                 'eventsPerSession',
#                 'screenPageViews',
#                 'screenPageViewsPerSession',
#                 'sessions',
#                 'sessionsPerUser',
#                 'totalUsers',
#                 'userEngagementDuration',
#                 'wauPerMau',
#                 'NewUsers',
#                 'ReturningUsers',
#                 'Login',
#                 'Entitlement',
#                 'Login/Entitlement',
#                 ]

EXCEL_HEADS = ['Date',
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
                'screenPageViews',
                'screenPageViewsPerSession',
                'sessions',
                'sessionsPerUser',
                'totalUsers',
                'userEngagementDuration',
                'wauPerMau',
                'NewUsers',
                'ReturningUsers',
                'EvolokEvents'
                ]


def insert_to_db(f_data):

    print(f'Inserting data: {f_data} into the database')
    
    connection = sqlite3.connect("Prensa_base_de_datos.db")
    cursor = connection.cursor()

    insert_values = ''

    for value in EXCEL_HEADS:

        if value == 'Date':

            insert_values += f"'{str(f_data[value])}',"

        else:

            insert_values += f'{f_data[value]},'

    try:

        # cursor.execute(f"""INSERT INTO Analytics_data VALUES ('{str(f_data['Date'])}',{f_data['activeUsers']},{f_data['averageSessionDuration']},{f_data['bounceRate']},
        # {f_data['crashFreeUsersRate']},{f_data['dauPerMau']},{f_data['dauPerWau']},{f_data['engagedSessions']},
        # {f_data['engagementRate']}, {f_data['eventCount']}, {f_data['eventCountPerUser']}, {f_data['eventsPerSession']}, {f_data['newUsers']},
        # {f_data['screenPageViews']}, {f_data['screenPageViewsPerSession']}, {f_data['sessions']}, {f_data['sessionsPerUser']},
        # {f_data['totalUsers']}, {f_data['userEngagementDuration']}, {f_data['wauPerMau']})""")

        cursor.execute(f"""INSERT INTO Analytics_historic_data VALUES ({insert_values.strip(',')})""")    

    except IntegrityError:

        print('Already exists')

    connection.commit()
    connection.close()

def update_db_columns(new_column, type):

    print(f'Adding column: {new_column} to the data base.')
    
    connection = sqlite3.connect("Prensa_base_de_datos.db")
    cursor = connection.cursor()

    if type == 'numero':

        cursor.execute(f"""ALTER TABLE Analytics_data_2 ADD {new_column} FLOAT""")    

        print('Column was added succesfully')

    elif type == 'texto':

        cursor.execute(f"""ALTER TABLE Analytics_data_2 ADD {new_column} VARCHAR(200)""")    

        print('Column was added succesfully')

    else:

        print('Incorrect type, unable to add new column')


def get_data_manual(start_date, end_date, email, password):

    start = date(year=int(start_date.split('-')[0]), month=int(start_date.split('-')[1]),day=int(start_date.split('-')[2]))

    end = date(year=int(end_date.split('-')[0]),month=int(end_date.split('-')[1]),day=int(end_date.split('-')[2]))

    data = []

    day_num = 0

    while True:

        data.append({'Date': start.strftime('%m/%d/%Y').replace('2023', '23').replace('2024', '24')})

        print(f"Getting info from date: {data[day_num]['Date']}")

        st_date_range = start.strftime('%Y-%m-%d')

        print(f'\nGetting metrics')

        try:

            metrics = [Metric(name=m) for m in METRICS_LIST]

            loops = math.floor(len(METRICS_LIST) / 10)

            last_loop_metrics = len(METRICS_LIST) % 10

            metric_count = 0

            for i in range(0, loops):

                request = RunReportRequest(
                    property=f"properties/{property_id}",
                    metrics=metrics[metric_count:metric_count+10],
                    date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)],
                )
            
                full_response = client.run_report(request)

                metric_count += 10

                try:

                    for n, metric in enumerate(full_response.rows[0].metric_values):

                        data[day_num][full_response.metric_headers[n].name] = metric.value

                        print(f'Extracted metric: {full_response.metric_headers[n].name} ---- value = {data[day_num][full_response.metric_headers[n].name]}')

                except IndexError:

                    print(f'metric: {metric} not available.')

                    data[day_num][metric] = 'Not available'


            if last_loop_metrics > 0:

                request = RunReportRequest(
                    property=f"properties/{property_id}",
                    metrics=metrics[metric_count:metric_count+last_loop_metrics],
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

            for metric in SPECIAL_METRICS:

                if metric == 'totalUsers':

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metric)],
                        dimensions=[Dimension(name='newVsReturning')],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)]
                    )

                    full_response = client.run_report(request)

                    try:

                        new = full_response.rows[0].metric_values[0].value

                        returning = full_response.rows[1].metric_values[0].value

                        data[day_num]['NewUsers'] = new

                        data[day_num]['ReturningUsers'] = returning

                        print(f"Extracted metric: NewUsers = ---- value = {data[day_num]['NewUsers']}")
                        print(f"Extracted metric: ReturningUsers ---- value = {data[day_num]['ReturningUsers']}")

                    except IndexError:

                        print(f'metric: {metric} not available.')

                        data[day_num][metric] = 'Not available'

                elif metric == 'eventCount':

                    request = RunReportRequest(
                        property=f"properties/{property_id}",
                        metrics=[Metric(name=metric)],
                        dimensions=[Dimension(name='eventName')],
                        date_ranges=[DateRange(start_date=st_date_range, end_date=st_date_range)],
                    )
                    full_response = client.run_report(request)

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

    return data

def get_metrics_data(email, password):

    """
    Function to get data from the GA API.
    """

    today = date.today()

    start = today - timedelta(days=8)
    end = today - timedelta(days=2)

    data = get_data_manual(start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d'), email, password)

    for row in data:

        insert_to_db(row)


def save_to_excel(metrics_data, start_date, end_date):

    print('Saving to excel...')

    df = pd.DataFrame(metrics_data, index =pd.RangeIndex(0,len(metrics_data)) ,columns = EXCEL_HEADS)
    
    df.to_excel(f"Metricas_desde_{start_date}_a_{end_date}.xlsx", index=True, columns=EXCEL_HEADS)


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

    workbook = openpyxl.load_workbook('Historico - GA3.xlsx', data_only=True)

    worksheet = workbook['Históricos']

    for row in worksheet.iter_rows(min_row=2, max_row=5114):

        data = {}

        for i, metric in enumerate(EXCEL_HEADS):

            if row[i].value == 'Not available':

                data[metric] = 'Null'

            elif '-' in str(row[i].value):

                print(str(row[i].value))

                try:
                
                    day = str(row[i].value).split('-')[2].replace(' 00:00:00', '')
                    month = str(row[i].value).split('-')[1] 
                    year = str(row[i].value).split('-')[0].replace('20','')

                    data[metric] = f"{month}/{day}/{year}"

                except IndexError:

                    data[metric] = str(row[i].value)

            elif ':' in str(row[i].value):

                minutes = int(str(row[i].value).split(':')[0])
                seconds = int(str(row[i].value).split(':')[1])

                final_time = float(minutes * 60 + seconds)

                data[metric] = final_time

            else:

                data[metric] = row[i].value

        print('-'*100)

        print(f'Exporting record {data}')

        print('-'*100)

        insert_to_db(data)

if __name__ == '__main__':

    export_hist_data_to_db()