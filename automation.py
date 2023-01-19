from os import environ
import ssl
import pandas as pd
import smtplib
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

def get_metrics_data(start_date, end_date, email, password) -> dict:

    """Function to get data from the GA API.
        
    """

    data = {}

    for metric in metrics_list:

        print(f'\nGetting metric: {metric}')

        request = RunReportRequest(
        property=f"properties/{property_id}",
        metrics=[Metric(name=metric)],
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)]
        )

        full_response = client.run_report(request)

        data[metric] = full_response.rows[0].metric_values[0].value

        print(f'Extracted metric: {metric} ---- value = {data[metric]}')

    save_to_excel(data, start_date, end_date)

    send_mail(email, password, start_date, end_date)

def save_to_excel(metrics_data, start_date, end_date):

    df = pd.DataFrame(metrics_data, columns=[metrics_data])

    df.to_excel(f"Metricas_desde_{start_date}_a_{end_date}.xlsx", index=False, columns=metrics_data)

def send_mail(email, password, start_date, end_date):
    
    mail = MIMEMultipart()
    mail['From'] = email
    mail['To'] = email
    mail['Subject'] = f"Reporte desde {start_date} a {end_date}"

    with open(f'Metricas_desde_{start_date}_a_{end_date}.xlsx', "rb") as f:
        part = MIMEApplication(
            f.read(),
            Name=basename(f)
        )

    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
    mail.attach(part)

    port = 465

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        
        server.login(email, password)
        
        server.sendmail(email, email, mail.as_string())



