from os import environ
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange
from google.analytics.data_v1beta.types import Dimension
from google.analytics.data_v1beta.types import Metric
from google.analytics.data_v1beta.types import RunReportRequest
from google.analytics.data_v1beta.types import OrderBy

property_id = '347990037'
environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Proyecto API-8b9749be5c38.json'

client = BetaAnalyticsDataClient()

def test(request):

    response = client.run_report(request)

    return response

request = RunReportRequest(
    property=f"properties/{property_id}",
    metrics=[Metric(name="activeUsers")],
    date_ranges=[DateRange(start_date="today", end_date="today")]
)

print(test(request))