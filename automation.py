from os import environ, remove
import ssl
import pandas as pd
import smtplib
from datetime import date
from dateutil.relativedelta import relativedelta
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange
from google.analytics.data_v1beta.types import Dimension
from google.analytics.data_v1beta.types import Metric
from google.analytics.data_v1beta.types import RunReportRequest
from google.analytics.data_v1beta.types import OrderBy
from google.api_core.exceptions import InvalidArgument

property_id = '347990037'
environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Proyecto API-8b9749be5c38.json'

client = BetaAnalyticsDataClient()

metrics_list = ['activeUsers',
                'adUnitExposure',
                'addToCarts',
                'advertiserAdClicks',
                'advertiserAdCost',
                'advertiserAdCostPerClick',
                'advertiserAdCostPerConversion',
                'advertiserAdImpressions',
                'averagePurchaseRevenue',
                'averagePurchaseRevenuePerPayingUser',
                'averagePurchaseRevenuePerUser',
                'averageRevenuePerUser',
                'averageSessionDuration',
                'bounceRate',
                'cartToViewRate',
                'checkouts',
                'cohortActiveUsers',
                'cohortTotalUsers',
                'conversions',
                'crashAffectedUsers',
                'crashFreeUsersRate',
                'dauPerMau',
                'dauPerWau',
                'ecommercePurchases',
                'engagedSessions',
                'engagementRate',
                'eventCount',
                'eventCountPerUser',
                'eventValue',
                'eventsPerSession',
                'firstTimePurchaserConversionRate',
                'firstTimePurchasers',
                'firstTimePurchasersPerNewUser',
                'itemListClickEvents',
                'itemListClickThroughRate',
                'itemListViewEvents',
                'itemPromotionClickThroughRate',
                'itemRevenue',
                'itemViewEvents',
                'itemsAddedToCart',
                'itemsCheckedOut',
                'itemsClickedInList',
                'itemsClickedInPromotion',
                'itemsPurchased',
                'itemsViewed',
                'itemsViewedInList',
                'itemsViewedInPromotion',
                'newUsers',
                'organicGoogleSearchAveragePosition',
                'organicGoogleSearchClickThroughRate',
                'organicGoogleSearchClicks',
                'organicGoogleSearchImpressions',
                'promotionClicks',
                'promotionViews',
                'publisherAdClicks',
                'publisherAdImpressions',
                'purchaseRevenue',
                'purchaseToViewRate',
                'purchaserConversionRate',
                'returnOnAdSpend',
                'screenPageViews',
                'screenPageViewsPerSession',
                'sessionConversionRate',
                'sessions',
                'sessionsPerUser',
                'shippingAmount',
                'taxAmount',
                'totalAdRevenue',
                'totalPurchasers',
                'totalRevenue',
                'totalUsers',
                'transactions',
                'transactionsPerPurchaser',
                'userConversionRate',
                'userEngagementDuration',
                'wauPerMau'
                ]

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

def get_metrics_data(d_range, email, password) -> dict:

    """Function to get data from the GA API.
        
    """

    date_range = set_dates(d_range)

    end_date = date_range[0]
    start_date = date_range[1]

    data = {}

    for metric in metrics_list:

        print('-'*100)

        print(f'\nGetting metric: {metric}')

        try:

            request = RunReportRequest(
                property=f"properties/{property_id}",
                metrics=[Metric(name=metric)],
                date_ranges=[DateRange(start_date=start_date, end_date=end_date)]
            )

            full_response = client.run_report(request)

        except InvalidArgument:
        

            print(f'Could not get info from the metric: {metric}')

            data[metric] = 'Not available'

            continue

        

        try:

            data[metric] = full_response.rows[0].metric_values[0].value

            print(f'Extracted metric: {metric} ---- value = {data[metric]}')

        except IndexError:

            print(f'metric: {metric} not available.')

            data[metric] = 'Not available'

    save_to_excel(data, start_date, end_date)

    send_mail(email, password, start_date, end_date)

def save_to_excel(metrics_data, start_date, end_date):

    print('Saving to excel...')

    df = pd.DataFrame(metrics_data, index=pd.RangeIndex(0,100) ,columns= metrics_data)

    df.to_excel(f"Metricas_desde_{start_date}_a_{end_date}.xlsx", index=True, columns=metrics_data)

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

if __name__ == '__main__':

    tu_email = 'nehuenv620@gmail.com'
    tu_contraseña = 'vzkfvhnhhiooiiuk'
    fecha_inicial = '2023-01-10'
    fecha_final = '2023-01-15'

    send_mail(tu_email, tu_contraseña, fecha_inicial, fecha_final)