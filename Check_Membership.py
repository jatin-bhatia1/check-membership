import requests, json, xlwt, openpyxl, pandas as pd
from Import_files.writeExcel import GenerateExcelFile

TOKEN_URL = 'https://api.helloasso.com/oauth2/token'

client_id = 'put client-id here'
client_secret = 'put client-secret here'
grant_type = 'client_credentials'

COURSE_SLUG = "cours-anglais-1er-semestre-2022-23"
COURSE_SLUG_API_URL = "https://api.helloasso.com/v5/organizations/lyon4water/forms/Event/{0}/items".format(COURSE_SLUG)

MEMBERSHIP_SLUG = "adhesion-annee-2023-2024"
MEMBERSHIP_SLUG_API_URL = "https://api.helloasso.com/v5/organizations/lyon4water/forms/Membership/{0}/items".format(MEMBERSHIP_SLUG)

ACCESS_TOKEN = ""
PARTICIPANTS = {}
MEMBERS = {}
COURSE_REGISTRATION_DATA = []

def recover_token():
    payload = {
        'client_id': client_id,
        'client_secret': client_secret,
        'grant_type': grant_type
    }
    response = requests.post(TOKEN_URL, data=payload)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        print(f"Error {response.status_code} occurred. Unable to retrieve the authorized token from the Api")
        return -1

def recover_data(API_URL):
    all_data = []
    PageSize = 100
    totalPages = 2
    PageIndex = 1

    headers = {
        "Authorization": "Bearer {0}".format(ACCESS_TOKEN)
    }

    while PageIndex <= totalPages:

        # format the api url with pageSize and PageIndex
        API_URL_formated = ("{0}?pageSize={1}&pageIndex={2}").format(API_URL, PageSize, PageIndex)
        
        # fetch the needed data
        response = requests.get(API_URL_formated, headers=headers)
        if response.status_code == 200:
            res = response.json()

            # Recover and set the totalPages number from the response 
            totalPages = res['pagination'] ['totalPages']

            # Append data from the current page to the list
            for element in res['data']:
                all_data.append(element)

            # Increase the pageIndex to get the next page
            PageIndex +=1
        else:
            print(f"Error {response.status_code} occurred.")
            return -1

    return all_data

def Load_Json(json_file_path):
    with open(json_file_path, 'r') as f:
        return json.load(f)

def Prepare_Course_Registration_Data():

    print("-------- Precessing Course Participants Data ---------------")

    for element in PARTICIPANTS:
       participant = {
            'Nom' : element['user']['lastName'] if("user" in element) else element['payer']['lastName'],
            'Prénom' : element['user']['firstName'] if("user" in element) else element['payer']['firstName'],
            'Email Achateur' : element['payer']['email'].strip(),
            'Date' : element['order']['date'].split('T')[0],
            'Montant' : int(element['amount']/100),
            'Type de tarif' : element['name'] if("user" in element) else "N/A"
        }
       COURSE_REGISTRATION_DATA.append(participant)

def Check_Membership():

    print("-------- Checking Participants Membership Data ---------------")

    for participant in COURSE_REGISTRATION_DATA:

        for member in MEMBERS:
            if member['payer']['email'].strip() == participant['Email Achateur'] :
                participant['Adhésion'] = "Payé"
                participant["Date adhésion"] = member['order']['date'].split('T')[0]
                break

            # HelloAsso user data is not consistent so we need to check first if the user data exists then we can do the further comparison
            elif ("user" in member):
                if(member['user']['firstName'] == participant['Prénom'] and member['user']['lastName'].strip() == participant['Nom']):
                    participant['Adhésion'] = "Payé"
                    participant["Date adhésion"] = member['order']['date'].split('T')[0]
                    break

            else:
                participant['Adhésion'] = "Non Payé"
                participant["Date adhésion"] = "N/A"
                #print(participant['Email Achateur'])

if __name__ == "__main__":

    # Recovering the authorization token first
    ACCESS_TOKEN = recover_token()

    # Retrieve participants data
    PARTICIPANTS = recover_data(COURSE_SLUG_API_URL)

    # Retrieve membership data
    MEMBERS = recover_data(MEMBERSHIP_SLUG_API_URL)

    # Preprae the course participants data correctly
    Prepare_Course_Registration_Data()

    # Check for every paricipants if they've the membership or not
    Check_Membership()

    # Recovers the data and Generates the excel file
    Excel_file_name = GenerateExcelFile(COURSE_REGISTRATION_DATA, COURSE_SLUG)

    