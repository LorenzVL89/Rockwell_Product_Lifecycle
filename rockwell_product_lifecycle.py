import requests
import pandas as pd
import easygui
from datetime import datetime


def get_status(ra_part_number):
    """
    Make an API request to retrieve information based on the RA part number.
    :param ra_part_number: RA part number for the search query.
    :return: JSON response from the API.
    """
    # Define the base URL and parameters for the API request
    base_url = "https://es-be-ux-search.cloudhub.io/api/ux/v2/search"
    params = {
        'queryText': ra_part_number,
        'role': 'rockwell-search',
        'spellingCorrect': True,
        'spellcheckPremium': 10,
        'segments': 'Productsv4',
        'startIndex': 0,
        'numResults': 20,
        'facets': '',
        'languages': 'en',
        'locales': 'en-GB,en_GLOBAL',
        'sort': 'bma',
        'collections': 'Literature,Web,Sample_Code',
        'site': 'RA'
    }

    # Define the headers required for the API request
    headers = {
        'client_id': 'fb000cbbe476420b9e70be741abd7a63',
        'client_secret': 'Db420ae8BAdD47ADA4E12cE90Fb1b747',
        'correlation_id': '1eaaba80-16be-6c60-9dd6-1378852fe624'
    }

    # Make the API request
    response = requests.get(base_url, params=params, headers=headers)

    # Raise an HTTPError for non-200 status codes
    response.raise_for_status()

    return response.json()


def process_data(row):
    part_number = row['Part No.']
    data = get_status(part_number)

    if 'response' in data and 'docs' in data['response']:
        for doc in data['response']['docs']:
            if part_number == doc['catalogNumber']:
                row['Description'] = str(doc.get('technicalDescription', ''))
                row['Lifecycle status'] = str(doc.get('lifecycleStatus', ''))
                if 'discontinuedDate' in doc:
                    discontinued_date_str = doc['discontinuedDate']
                    discontinued_date = datetime.fromisoformat(discontinued_date_str.replace('Z', '+00:00'))
                    row['Discontinued date'] = discontinued_date.strftime('%Y-%m-%d')
                if 'replacementText' in doc and 'replacementCategory' in doc:
                    row['Replacement Part No.'] = str(doc.get('replacementText', ''))
                    row['Replacement category'] = str(doc.get('replacementCategory', ''))
    return row


if __name__ == "__main__":
    # Define the location of the Excel document
    excel_file = easygui.fileopenbox()
    # Get the part numbers from an Excel Spreadsheet (first Column)
    df = pd.read_excel(excel_file)

    # Apply the process_data function to each row using apply
    df = df.apply(process_data, axis=1)

    # Save the updated dataframe to a new Excel file with a timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    df.to_excel(excel_file.replace('.xlsx', f'_{timestamp}.xlsx'), index=False)
