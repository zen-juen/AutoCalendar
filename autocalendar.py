# -*- coding: utf-8 -*-
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow

import pandas as pd
import numpy as np
import pickle
import dateutil.parser
from datetime import datetime


def preprocess_file(file):
    """Tidy excel sheet.

    This function assumes that there are 2 header rows and it is selecting the second row as
    the main column names of interest, which will be needed in `extract_info`."""

    participants = pd.read_excel(file)
    participants.columns = participants.iloc[0]
    participants = participants.reindex(participants.index.drop(0)).reset_index(drop=True)

    return participants


def extract_info(participants, date_col, time_col, location_col,
                 filter_column=None, select=None):
    '''Extract date, time, and location of event based on header column names in the
    participants file (`date_col`, `time_col`, `location_col` respectively).

    Defaults to adding all existing participants' scheduled slots in
    the excel sheet into google calendar. To select only some participants,
    activate `filter_column` as the header column to filter by, and `select`
    as the cell entries in this column which determine the selected participants.

    '''

    if filter_column is not None and select is not None:
        to_add = participants[participants[filter_column] == select]
    else:
        to_add = participants

    date = np.array(to_add[date_col])
    location = np.array(to_add[location_col])

    # Format time
    timing = np.array(to_add[time_col])
    timings = np.array([])
    for t in timing:
        if '.' in t:
            t_new = t.replace('.', ':')  # can only detect time with colon
            timings = np.append(timings, t_new)


    start_points = np.array([])
    end_points = np.array([])
    for i in timings:
    # split time entry into start and end time based on '-'
        start = dateutil.parser.parse(i.split('-')[0]).time()
        end = dateutil.parser.parse(i.split('-')[1]).time()
        start_points = np.append(start_points, start)
        end_points = np.append(end_points, end)

    return date, start_points, end_points, location


def create_event(event_name, description, date, start, end, location, timezone, creator_email,
                 calendar_id='primary'):
    """Create event in terms of Google Calendar API.

    See also https://developers.google.com/calendar/v3/reference/events

    """

    event = {
      'summary': event_name,
      'colorId': '9', #blue
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
# Workflow and execution
# =============================================================================


def main():

    # Set up credentials
    scopes = ['https://www.googleapis.com/auth/calendar']
    flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", scopes=scopes)
    credentials = flow.run_console()  # need to manually enter authorization code here

    pickle.dump(credentials, open("token.pkl", "wb"))
    credentials = pickle.load(open("token.pkl", "rb"))
    service = build("calendar", "v3", credentials=credentials)

    # Get details
    participants = preprocess_file(file='C:/Users/Zen Juen/Dropbox/Deception_MockCrime/Deception_MockCrime/Participants Scheduling/Master_Participant_List.xlsx')

    # Modify column name entries as per your input file
    dates, start_points, end_points, locations = extract_info(participants,
                                                              date_col='Date_Session1',
                                                              time_col='Timeslot_Session1',
                                                              location_col='Location_Session1',
                                                              filter_column='Calendar_Event',
                                                              select='No')
        # select as No because it represents that the 'Calendar_Event' has not been added yet.

    # Create event
    if len(dates) > 1: # if more than one participant to add:
        events = []
        calendar_ids = []
        for date, start, end, location in zip(dates, start_points, end_points, locations):
            event, calendar_id = create_event(event_name='fMRI study Session 1', description='',
                                      date=date, start=start, end=end, location=location,
                                      timezone='Asia/Singapore',
                                      creator_email='decisiontask.study@gmail.com',
                                      calendar_id='Lab Use (NTU)')
            events.append(event)
    else:
        event, calendar_id = create_event(event_name='fMRI study Session 1', description='',
                                          date=date, start=start_points, end=end_points,
                                          location=locations,
                                          timezone='Asia/Singapore',
                                          creator_email='decisiontask.study@gmail.com',
                                          calendar_id='Lab Use (NTU)')

    # Execute
    ## If 'calendar_id' in `create_event` is set to primary, then use primary calendar, if not
    ## input the string of the calendar name that you intend to use.
    result = service.calendarList().list().execute()
    if calendar_id != 'primary':
        choose = []
        for dictionary in result['items']:  # for name, age in dictionary.iteritems():  (for Python 2.x)
            if dictionary['summary'] == calendar_id:
                choose.append(dictionary['id'])
        calendar_id = choose[0]
    else:
        calendar_id='primary'

    if len(dates) > 1:
        for i in events:
            service.events().insert(calendarId=calendar_id, body=i).execute()
    else:
        service.events().insert(calendarId=calendar_id, body=event).execute()



if __name__ == "__main__":
    main()
