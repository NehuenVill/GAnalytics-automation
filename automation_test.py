from google.oauth2 import service_account
from googleapiclient.discovery import build
from os import environ

SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Proyecto API-8b9749be5c38.json'
VIEW_ID = '347990037'

def get_data():

    credentials = service_account.Credentials.from_service_account_file(
        environ['GOOGLE_APPLICATION_CREDENTIALS'], scopes=SCOPES
    )

    analytics = build('analyticsreporting', 'v4', credentials=credentials)

    query = {
        'viewId' : VIEW_ID,
        'dateRanges': [{'startDate':'2023-01-01', 'endDate' : '2023-03-01'}],
        'metrics' : [{'expression':'ga:totalEvents'}],
        'dimensions' : [{'name' : 'ga:eventLabel'}],
        'dimensionFilterClauses' : [{
            
            'filters': [{

                'dimensionName' : 'ga:eventAction',
                'operator' : 'EXACT',
                'expressions' : ['EVOLOK']
                
                }]
        }]
    }

    response = analytics.reports().batchGet(body={'reportRequest' : [query]}).execute()

    print(response)


if __name__ == '__main__':

    get_data()