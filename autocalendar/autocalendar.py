# -*- coding: utf-8 -*-
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import pandas as pd
import numpy as np
import pickle
import os.path
import dateutil.parser
import calendar
from datetime import datetime

# =============================================================================
# Scheduling tool
# =============================================================================
def autoallocate(file, transpose=True, filename='allocations', export_to='xlsx'):
    """Read and parse a downloaded doode poll (in '.xls' or '.xlsx') where participants are
    able to choose as many timeslots as possible. Automatically allocate participants to a
    slot based on their chosen availabilities. Returns dataframe containing the participants'
    allocated slots.

    Parameters
    ----------
    file : str
        Path containing doodle poll file
    transpose : bool
        Exported dataframe will be in long format if True.
    filename : str
        Name of the file containing the participants' allocations.
    export_to : str
        Exported file type. Can be 'xlsx' or 'csv'.
    """

    # Read and parse doodle poll
    poll = pd.read_excel(file)
    poll.loc[2:4,].fillna(method='ffill', axis=1, inplace=True)
    poll = poll.set_index(poll.columns[0])

    # Extract all possible datetimes
    datetimes = []
    for month, date, time in zip(poll.iloc[2], poll.iloc[3], poll.iloc[4]):
        exact_date = month + ' ' + date
        parsed_date = dateutil.parser.parse(exact_date)
        day = calendar.day_name[parsed_date.weekday()]
        dt = parsed_date.strftime("%d %B") + ', ' + day + ', ' + time
        datetimes.append(dt)

    poll.columns = datetimes
    poll = poll.iloc[5:]

    # Create empty df for appending assigned dates
    empty_df = pd.DataFrame(index=['Participant'], columns=datetimes)

    # Allocate slots
    assignments = []

    for assigned_date in poll.columns:
        # Number of subjects who chose the slot
        n_selections = poll[assigned_date].astype(str).str.contains('OK').sum()

        if n_selections == 0:
            empty_df[assigned_date] = np.nan

        elif n_selections == 1:
            single_name = poll[poll[assigned_date] == 'OK'].index.values[0]
            if single_name not in assignments:
                # If subject has not been assigned yet
                empty_df[assigned_date] = single_name
                assignments.append(single_name)
            else:
                empty_df[assigned_date] = np.nan

        elif n_selections > 1:
            multiple_names = poll[poll[assigned_date] == 'OK'].index.values
            chosen_name = np.random.choice(multiple_names)
            if chosen_name not in assignments:
                empty_df[assigned_date] = chosen_name
                assignments.append(chosen_name)
            else:
                chosen_name_2 = np.random.choice(multiple_names[multiple_names != chosen_name])
                if chosen_name_2 not in assignments:
                    empty_df[assigned_date] = chosen_name_2
                    assignments.append(chosen_name_2)
                else:
                    empty_df[assigned_date] = np.nan

    allocations = empty_df.copy()
    if transpose:
        allocations = pd.DataFrame.transpose(allocations)
        allocations = allocations.rename(columns={ allocations.columns[0]: "Participant" })

    # Export
    if export_to == 'csv':
        allocations.to_csv(filename + '.csv')
    elif export_to == 'xlsx':
        allocations.to_excel(filename + '.xlsx')

    # Feedback
    participants = poll.index[poll.index != 'Count'].tolist()
    for participant in participants:
        if participant not in assignments:
            print(f'{participant}' + ' could not be allocated.')
    if len(np.intersect1d(participants, assignments)) == len(participants):
        print('All participants successfully allocated.')


# =============================================================================
# Initialize Google API Console Credentials
# =============================================================================

def setup_oath(token_path, client_path):
    """
    Path containing token.pkl and client_secret.json respectively.
    """
    # Set up credentials
    scopes = ['https://www.googleapis.com/auth/calendar']

    # Token generated after first time code is run.
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            credentials = pickle.load(token)

    # If there are no (valid) credentials available, log in and enter authorization code manually
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_path, scopes=scopes)
            credentials = flow.run_local_server(port=0)

       # Save the credentials for the next run
        with open(token_path, 'wb') as token:
            pickle.dump(credentials, token)

    service = build("calendar", "v3", credentials=credentials)

    return service

# =============================================================================
# Get event information
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

def add_event(service, dates, start_points, end_points, locations, to_add,
              creator_email, event_name='Experiment', description='',
              timezone='Asia/Singapore', calendar_id='primary', silent=False,
              name_col=None, date_col=None, time_col=None, location_col=None,
              starttime_col=None, endtime_col=None):
    """Execute adding of event into google calendar. Set `service` as the resource built from the
    googleapiclient, e.g., service = autocalendar.setup_oath()

    If silent is set to True, print feedback of information that is added, columns to be denoted by
    `*_col` (Otherwise set to None).
    """

    events = []
    if len(to_add) > 1:  # If more than one event to add
        for date, start, end, location in zip(dates, start_points, end_points, locations):
            event, calendar_id = create_event(event_name=event_name, description=description,
                                              date=date, start=start, end=end, location=location,
                                              timezone=timezone,
                                              creator_email=creator_email,
                                              calendar_id=calendar_id)
            events.append(event)

    elif len(to_add) == 1:  # If only one event to add
        event, calendar_id = create_event(event_name=event_name, description=description,
                                          date=dates[0], start=start_points[0], end=end_points[0],
                                          location=locations[0],
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
