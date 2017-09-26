# outlook_google_calender_sync

This can be useed to sync outlook calender with google calender.

Currently it syncs following event types

  1. Single occurence events
  2. Recurring events with Weekly recurrence.

The script creates a new calendar with the used supplied name. It will first delete the calender 
in google account if a calender with given name already exists.

The script uses exchangelib and google calendar apis. Please download and install it before using this script.
Run on Python V3.

Usage : Define following parameters in the init function :

new_cal_name = Calender Name for newly created cal
tz_name = Timezone to use e.g. 'America/Chicago'
outlook_user_name =
outlook_user_pass = 
outlook_email = 
google_email =
