# -*- coding: utf-8 -*-
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import pandas as pd
import numpy as np
import pickle
import os.path
import dateutil.parser
from datetime import datetime


# =============================================================================
# Initialize Google API Console Credentials
# =============================================================================

def setup_oath():

    # Set up credentials
    scopes = ['https://www.googleapis.com/auth/calendar']

    # Token generated after first time code is run.
    if os.path.exists('../secret/token.pkl'):
        with open('../secret/token.pkl', 'rb') as token:
            credentials = pickle.load(token)
    else:
        credentials = object()

    # If there are no (valid) credentials available, log in and enter authorization code manually
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
           credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("../secret/client_secret.json", scopes=scopes)
            credentials = flow.run_local_server(port=0)

       # Save the credentials for the next run
        with open('../secret/token.pkl', 'wb') as token:
            pickle.dump(credentials, token)

    service = build("calendar", "v3", credentials=credentials)

    return service

# =============================================================================
# Get event information and implement
# =============================================================================


def preprocess_file(file, header_row=1):
    """Tidy excel sheet containing participants' particulars.

    If there are multiple header rows, denote the header row to be selected in `header_row` (defaults to 1). For example, `header_row=2` specifies the second row as the main column names of interest, which will be needed in `extract_info()`.
    """

    participants = pd.read_excel(file)

    if header_row > 1:
        participants.columns = participants.iloc[header_row-2]
        participants = participants.reindex(participants.index.drop(0)).reset_index(drop=True)

    return participants


def extract_info(participants, date_col, time_col, location_col, starttime_col=None, endtime_col=None,
                 filter_column=None, select=None):
    """Extract date, time, and location of event based on header column names in the
    participants file (`date_col`, `time_col`, `location_col` respectively).

    It is assumed that `time_col` contains both the start and end time e.g., '10.00am-12.00pm'.
    If not, specify these details respectively in `starttime_col` and `endtime_col`, in HHMM format.
    Include 'AM' or 'PM' if not reported in 24HR format.

    Defaults to adding all existing participants' scheduled slots in
    the excel sheet into google calendar. To select only some participants,
    activate `filter_column` as the header column to filter by, and `select`
    as the cell entries in this column which determine the selected participants.

    Returns dates, start times, end times, locations and dataframe of filtered participants whose
    details are to be added into google calendar.
    """

    if filter_column is not None and select is not None:
        to_add = participants[participants[filter_column] == select]
    else:
        to_add = participants

    date = np.array(to_add[date_col])
    location = np.array(to_add[location_col])

    # Format time
    start_points = np.array([])
    end_points = np.array([])

    if not starttime_col and not endtime_col:
        timing = np.array(to_add[time_col])
        timings = np.array([])
        for t in timing:
            if '.' in t:
                t_new = t.replace('.', ':')  # can only detect time with colon
                timings = np.append(timings, t_new)

        for i in timings:
        # split time entry into start and end time based on '-'
            start = dateutil.parser.parse(i.split('-')[0]).time()
            end = dateutil.parser.parse(i.split('-')[1]).time()
            start_points = np.append(start_points, start)
            end_points = np.append(end_points, end)

    else:
        starttime_col = [start.replace('.', ':') for start in starttime_col]
        start_points = np.append(start_points, starttime_col)
        endtime_col = [end.replace('.', ':') for end in endtime_col]
        end_points = np.append(end_points, endtime_col)


    return date, start_points, end_points, location, to_add


def create_event(event_name, description, date, start, end, location, timezone, creator_email,
                 calendar_id='primary'):
    """Create event in terms of Google Calendar API.

    See also https://developers.google.com/calendar/v3/reference/events

    """

    event = {
      'summary': event_name,
      'colorId': '9',  #blue
      'kind': 'calendar#event',
      'location': location,
      'description': description,
      'start': {
        'dateTime': datetime.combine(date, start).strftime("%Y-%m-%dT%H:%M:%S"),
        'timeZone': timezone,
      },
      'end': {
        'dateTime': datetime.combine(date, end).strftime("%Y-%m-%dT%H:%M:%S"),
        'timeZone': timezone,
      },
      'reminders': {
        'useDefault': False,
        'overrides': [
#          {'method': 'email', 'minutes': 24 * 60},  # unhash for email notifications
          {'method': 'popup', 'minutes': 10},
        ],
      },
    }

    return event, calendar_id


# =============================================================================
# Execution
# =============================================================================


def add_event(dates, start_points, end_points, locations, to_add,
              event_name='Experiment', description='', timezone='Asia/Singapore',
              creator_email, calendar_id='primary', silent=False,
              name_col=None, date_col=None, time_col=None, starttime_col=None, endtime_col=None,
              date_col=None):
    """Execute adding of event into google calendar.

    If silent is set to True, print feedback of information that is added, columns to be denoted by
    `*_col` (Otherwise set to None).

    """

    if len(to_add) > 1:  # If more than one event to add
        events = []
        for date, start, end, location in zip(dates, start_points, end_points, locations):
            event, calendar_id = create_event(event_name=event_name, description=description,
                                              date=date, start=start, end=end, location=location,
                                              timezone=timezone,
                                              creator_email=creator_email,
                                              calendar_id=calendar_id)
            events.append(event)

    elif len(to_add) == 1:  # If only one event to add
        event, calendar_id = create_event(event_name=event_name, description=description,
                                          date=dates, start=start_points, end=end_points,
                                          location=locations,
                                          timezone=timezone,
                                          creator_email=creator_email,
                                          calendar_id=calendar_id)
        events.append(event)

    # If 'calendar_id' in `create_event` is set to primary, then use primary calendar, if not
    # input the string of the calendar name that you intend to use.
    result = service.calendarList().list().execute()
    if calendar_id != 'primary':
        choose = []
        for dictionary in result['items']:
            if dictionary['summary'] == calendar_id:
                choose.append(dictionary['id'])
        calendar_id = choose[0]

    # Execute
    if len(to_add) > 1:
        for i in events:
            service.events().insert(calendarId=calendar_id, body=i).execute()
    else:
        service.events().insert(calendarId=calendar_id, body=event).execute()

    # Print output
    if not silent:
        for name in to_add[name_col]:
            info_date = to_add[date_col][to_add[name_col] == name].iloc[0].date().strftime("%d-%m-%Y")
            if not starttime_col and not endtime_col:
                info_time = to_add[time_col][to_add[name_col] == name].iloc[0]
            else:
                info_time = to_add[starttime_col][to_add[name_col] == name].iloc[0] + ' - ' + to_add[endtime_col][to_add[name_col] == name].iloc[0]
            info_location = to_add[location_col][to_add[name_col] == name].iloc[0]

            print('Adding calendar event for ' + f'{name} ' + 'at ' + f'{info_date}, '
                  + f'{info_time}, ' + f'{info_location} ')
