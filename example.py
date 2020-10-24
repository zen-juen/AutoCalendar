import autocalendar as autocalendar

# Set up Oauth to access Google API
service = autocalendar.setup_oath()

# Read and tidy excel sheet
participants = autocalendar.preprocess_file('../../Participants Scheduling/Master_Participant_List.xlsx',
                                            header_row=2)

### Session 1

# Extract info
dates, start_points, end_points, locations, to_add = autocalendar.extract_info(participants, date_col='Date_Session1', time_col='Timeslot_Session1', location_col='Location_Session1', filter_column='Calendar_Event', select='No')

# Execute
autocalendar.add_event(service, dates, start_points, end_points, locations, to_add,
                       event_name='fMRI study Session 1', description='', timezone='Asia/Singapore',
                       creator_email='decisiontask.study@gmail.com', calendar_id='Lab Use (NTU)',
                       silent=False, name_col='Participant Name', date_col='Date_Session1',
                       location_col='Location_Session1', time_col='Timeslot_Session1')

### Session 2

# Extract info
dates, start_points, end_points, locations, to_add = autocalendar.extract_info(participants, date_col='Date_Session2', time_col='Timeslot_Session2', location_col='Location_Session2', filter_column='Calendar_Event', select='No')

# Execute
autocalendar.add_event(service, dates, start_points, end_points, locations, to_add,
                       event_name='fMRI study Session 2', description='', timezone='Asia/Singapore',
                       creator_email='decisiontask.study@gmail.com', calendar_id='Lab Use (NTU)',
                       silent=False, name_col='Participant Name', date_col='Date_Session2',
                       location_col='Location_Session2', time_col='Timeslot_Session2')
