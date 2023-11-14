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

    # Check if the response status code is 200 (OK)
    if response.status_code == 200:
        # Parse and return the response data as JSON
        return response.json()
    else:
        # Return an error message if the status code is not 200
        raise requests.HTTPError(f"Request failed with status code {response.status_code}")


if __name__ == "__main__":
    # Get the user input for the part number
    # part_number = input("Enter the RA Part Number: ").upper()

    # Define the location of the Excel document
    excel_file = easygui.fileopenbox()
    # Get the part numbers from an Excel Spreadsheet (first Column)
    df = pd.read_excel(excel_file)

    # Explicitly set the dtype for columns that will be updated
    df['Description'] = df['Description'].astype(str)
    df['Lifecycle status'] = df['Lifecycle status'].astype(str)
    df['Discontinued date'] = df['Discontinued date'].astype(object)  # or adjust to the appropriate dtype
    df['Replacement Part No.'] = df['Replacement Part No.'].astype(str)
    df['Replacement category'] = df['Replacement category'].astype(str)

    for i, row in df.iterrows():
        part_number = row['Part No.']
        # print(part_number)
        data = get_status(part_number)
        if 'response' in data and 'docs' in data['response']:
            for y, doc in enumerate(data['response']['docs']):
                if part_number == doc['catalogNumber']:
                    df.at[i, 'Description'] = str(doc['technicalDescription'])
                    # print(row['Description'])
                    df.at[i, 'Lifecycle status'] = str(doc['lifecycleStatus'])
                    # print(row['Lifecycle status'])
                    if 'discontinuedDate' in doc:
                        discontinued_date_str = doc['discontinuedDate']
                        discontinued_date = datetime.fromisoformat(discontinued_date_str.replace('Z', '+00:00'))
                        df.at[i, 'Discontinued date'] = discontinued_date.strftime('%Y-%m-%d')
                    if 'replacementText' in doc and 'replacementCategory' in doc:
                        df.at[i, 'Replacement Part No.'] = str(doc['replacementText'])
                        df.at[i, 'Replacement category'] = str(doc['replacementCategory'])

    # Save the updated dataframe to a new Excel file with a timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    df.to_excel(excel_file.replace('.xlsx', f'_{timestamp}.xlsx'), index=False)
