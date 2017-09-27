# outlook_google_calender_sync

This can be useed to sync outlook calender with google calender.


Currently it syncs following event types

  1. Single occurence events
  2. Recurring events with Weekly recurrence.

The script creates a new calendar with the used supplied name. It will first delete the calender
in google account if a calender with given name already exists.

The script uses exchangelib , google calendar apis and sqlite. Please download and install required packages for using this script.
Run on Python V3. Also you will need to create a google calender project and save the client secret file as client_secret.json
in the script's directoy.

Usage : Define following parameters in the init function :


new_cal_name = Calender Name for newly created cal
tz_name = Timezone to use e.g. 'America/Chicago'
outlook_user_name =
outlook_user_pass =
outlook_email =
google_email =
max_fut_days_to_check - Number of days to sync events in future
bReinit = False. Set This variable to True if you want to reinitialize ( delete and recreate calender on google ) or if there is some problem.
           Set this to True rarely or else it may cause google usage limit to exceed due to creation of calender multiple times    
