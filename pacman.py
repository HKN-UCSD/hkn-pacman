# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START sheets_quickstart]
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import date
from dotenv import load_dotenv

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# source sheet id and range
load_dotenv()
    # event log
SOURCE_EVENT_SHEET_ID = os.getenv('SOURCE_EVENT_SHEET_ID')
SOURCE_EVENT_SHEET_RANGE = os.getenv('SOURCE_EVENT_SHEET_RANGE')
    # mentor log
SOURCE_MENTOR_SHEET_ID = os.getenv('SOURCE_MENTOR_SHEET_ID')
SOURCE_MENTOR_SHEET_RANGE = os.getenv('SOURCE_MENTOR_SHEET_RANGE')

# destination sheet id and range
    # total inductee point
DESTINATION_TOTAL_SHEET_ID = os.getenv('DESTINATION_TOTAL_SHEET_ID')
DESTINATION_TOTAL_SHEET_RANGE_HEADER = os.getenv('DESTINATION_TOTAL_SHEET_RANGE_HEADER')
DESTINATION_TOTAL_SHEET_RANGE_BODY = os.getenv('DESTINATION_TOTAL_SHEET_RANGE_BODY')
    # inductee finished
DESTINATION_DONE_SHEET_ID = os.getenv('DESTINATION_DONE_SHEET_ID')
DESTINATION_DONE_SHEET_RANGE_BODY = os.getenv('DESTINATION_DONE_SHEET_RANGE_BODY')
    # total mentor point
DESTINATION_MENTOR_TOTAL_SHEET_ID = os.getenv('DESTINATION_MENTOR_TOTAL_SHEET_ID')
DESTINATION_MENTOR_TOTAL_SHEET_RANGE_HEADER = os.getenv('DESTINATION_MENTOR_TOTAL_SHEET_RANGE_HEADER')
DESTINATION_MENTOR_TOTAL_SHEET_RANGE_BODY = os.getenv('DESTINATION_MENTOR_TOTAL_SHEET_RANGE_BODY')

# enum for updating data (specific to google sheet api)
from enum import Enum
class Dimension(Enum):
    DIMENSION_UNSPECIFIED = 0
    ROWS = 1
    COLUMNS = 2
# class to store inductee points
class InducteeValues:
    def __init__(self):
        self.points = 0.0
        self.has_mentor = 0
        self.has_professional = 0
        self.officer_list = set()
        self.event_list = []

# get google sheet service 
def get_service():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

# get google sheet
def get_sheet(service, sheet_id, sheet_range):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=sheet_range).execute()
    return result.get('values', [])

# update google sheet
def update_sheet(service, sheet_id, sheet_range, value_range_body):
    request = service.spreadsheets().values().update(spreadsheetId=sheet_id, range=sheet_range, valueInputOption='RAW', body=value_range_body)
    request.execute()
    #response = request.execute()

def get_event_data(values):
    # get total event points per email, store in dictionary
    # index based on column number of event source sheet
    source_event_data = {}
    source_event_email_index = 1
    source_event_event_index = 7
    source_event_point_index = 8
    source_event_officer_index = 9
    
    # for each event log
    for row in values:
        # if not in dictionary, make a new key-value pair
        if source_event_data.get(row[source_event_email_index].lower()) == None:
            source_event_data[row[source_event_email_index].lower()] = InducteeValues()

        # add points
        source_event_data[row[source_event_email_index].lower()].points += float(row[source_event_point_index])
        
        # check professional
        if "resume" in row[source_event_event_index].lower():
            source_event_data[row[source_event_email_index].lower()].has_professional = 1
        elif "interview" in row[source_event_event_index].lower():
            source_event_data[row[source_event_email_index].lower()].has_professional = 1
        elif "pitch" in row[source_event_event_index].lower():
            source_event_data[row[source_event_email_index].lower()].has_professional = 1

        # add events
        source_event_data[row[source_event_email_index].lower()].event_list.append(row[source_event_event_index])
        
        # add officer
        source_event_data[row[source_event_email_index].lower()].officer_list.add(row[source_event_officer_index])

    return source_event_data

def append_mentor_data(values, source_event_data):
    # index based on mentor source sheet column
    source_mentor_email_index = 2

    # for each mentor log
    for row in values:
        # if not in dictionary, make a new key-value pair
        if source_event_data.get(row[source_mentor_email_index].lower()) == None:
            source_event_data[row[source_mentor_email_index].lower()] = InducteeValues()
       
        # add points
        source_event_data[row[source_mentor_email_index].lower()].points += 1
        
        # add mentor
        source_event_data[row[source_mentor_email_index].lower()].has_mentor = 1

        # add event
        source_event_data[row[source_mentor_email_index].lower()].event_list.append("Mentor 1:1")

    return source_event_data

def update_total_log_data(service, source_event_data):
    # Overwrite to total log points header
    value_range_body = {
        "majorDimension": Dimension.ROWS.value,
        "values": [ 
            ["Inductee Email", "Total Points ("+str(date.today())+")", "Mentor 1:1", "Professional", "Event List", "Officer List"]
        ]
    }
    update_sheet(service, DESTINATION_TOTAL_SHEET_ID, DESTINATION_TOTAL_SHEET_RANGE_HEADER, value_range_body)

    # Overwrite to total log points body
    source_event_data_values = source_event_data.values()
    value_range_body = {
        "majorDimension": Dimension.COLUMNS.value,
        "values": [ 
            source_event_data.keys(),                                           # email
            [val.points for val in source_event_data_values],                   # total points
            [val.has_mentor for val in source_event_data_values],               # has mentor 1:1
            [val.has_professional for val in source_event_data_values],         # has professional event: resume/interview/pitch
            [",".join(val.event_list) for val in source_event_data_values],     # event list
            [",".join(val.officer_list) for val in source_event_data_values]    # officer list
        ]
    }
    update_sheet(service, DESTINATION_TOTAL_SHEET_ID, DESTINATION_TOTAL_SHEET_RANGE_BODY, value_range_body)

def update_mentor_log_data(service, source_event_data):
    # Overwrite to total log points header
    value_range_body = {
        "majorDimension": Dimension.ROWS.value,
        "values": [ 
            ["Mentor Email", "Total Points ("+str(date.today())+")"]
        ]
    }
    update_sheet(service, DESTINATION_MENTOR_TOTAL_SHEET_ID, DESTINATION_MENTOR_TOTAL_SHEET_RANGE_HEADER, value_range_body)

    # index based on mentor source sheet column
    source_mentor_email_index = 5

    # mentor dictionary count
    mentor_point_dict = {}

    # for each mentor log
    for row in values:
        mentor_point_dict[row[source_mentor_email_index]] = mentor_point_dict.get(row[source_mentor_email_index].lower(), 0) + 1 
    
    # Overwrite to total log points body
    mentor_point_list = sorted(mentor_point_dict.iteritems(), key = lambda x: x[1])
    value_range_body = {
        "majorDimension": Dimension.COLUMNS.value,
        "values": [ 
            [val[0] for val in mentor_point_list],                              # email
            [val[1] for val in mentor_point_list]                               # total points
        ]
    }
    update_sheet(service, DESTINATION_MENTOR_TOTAL_SHEET_ID, DESTINATION_MENTOR_TOTAL_SHEET_RANGE_BODY, value_range_body)

def update_inductee_list(service, source_event_data):
    # update inductee list
    value_range_body = {
        "majorDimension": Dimension.COLUMNS.value,
        "values": [
            # if inductee has requirements of 10 or more points, one mentor, and one professional event
            [val for val in source_event_data.keys() if source_event_data[val].points >= 10 and source_event_data[val].has_mentor and source_event_data[val].has_professional ] 
        ]
    }
    update_sheet(service, DESTINATION_DONE_SHEET_ID, DESTINATION_DONE_RANGE_BODY, value_range_body)

def main():
    # get the google sheet service
    service = get_service()

    # get the sheet values of event log
    values = get_sheet(service, SOURCE_EVENT_SHEET_ID, SOURCE_EVENT_SHEET_RANGE)
    
    # get the data of event logs
    source_event_data = get_event_data(values)

    # get the sheet values of mentor log
    values = get_sheet(service, SOURCE_MENTOR_SHEET_ID, SOURCE_MENTOR_SHEET_RANGE)
    
    # update mentor log data
    update_mentor_total_log_data(service, values)

    # add mentor data into event data
    source_event_data = append_mentor_data(values, source_event_data)

    # update total log data
    update_total_log_data(service, source_event_data)
    

if __name__ == '__main__':
    main()
# [END sheets_quickstart]
