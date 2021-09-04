# Dispatcher-Web
Application for the dispatcher to work with the Google Calendar via Web API.

Now it works as web app.

It is accessible at https://dispatcher-omsu.googleusercontenturl.com:5000/

**Input files: options.txt**

It uses parameters given in file "options.txt".

- The first line - first parameter is the start date. All events *preceding* it will be omitted.

- The second line - second parameter is the end date. All events *following* it will be omitted.

- The third line - third parameter is the starting week of the semester. Is is used in output files to enumerate weeks of semester.

*Example input in options.txt*
```
2021-01-01 08:00:00
2021-12-01 00:00:00
34
```

This file can be modified at page /edit_options.html (accessible from top menu). This page contains interactive fields and buttons to save ptions and to restore default options.

**Input files: tutors.txt**

This file contains list of guests (tutors) to find in calendars

*Example input in tutors.txt*
```
deminie@college.omsu.ru
dan-909-o@college.omsu.ru
dtn-809-o@college.omsu.ru
```

This file can be modified at page /index.html (main page and is accessible from top menu). This page contains text area, where list of tutors may be inserted from clipboard.

**Output files**

Program creates multiple google spreadsheets with timetable for each tutor, given on input.

**uses**
``` 
flask==2.0.1
google==3.0.0
google-api-python-client==2.15.0
google_auth_oauthlib==0.4.4
``` 

Uses googleapiclient library from
https://github.com/googleapis/google-api-python-client

**Test options**
```
pip install selenium
pip install pynput

chromedriver
https://sites.google.com/chromium.org/driver/
```

**Setup**

There are a few setup steps you need to complete before you can use this library:

1.  If you don't already have a Google account, [sign up](https://www.google.com/accounts).
2.  If you have never created a Google APIs Console project, read the [Managing Projects page](http://developers.google.com/console/help/managing-projects) and create a project in the [Google API Console](https://console.developers.google.com/).
3.  [Install](http://developers.google.com/api-client-library/python/start/installation) the library.
4.  To add new tester while app is in test mode -> go to your [Google APIs Console](https://console.cloud.google.com/apis/credentials/consent)

**Functions**
```
def get_options():
    # It reads parameters from file: lower_date, upper_date, group
    # Those params are to be passed to function list_events_by_param
    # That function will list events, filtered by params
    
def get_calendar_dict():
    # This function retrieves list of calendars for user
    # Returns dict if calendar_ID: calenar_summary
    
def del_all_calendar_events(service, options):
    # This function deletes all events before given date
    # Also searches through all calendars for given user

def load_into_spreadsheet(service, options, timetable):
    # Input is: spreadsheet service, options dict and timetable object
    # Call the Sheets API, creates new spreadsheet
    # and loads data from timetable into spreadsheet
    # Outputs link to created spreadsheet on the screen
    
def get_authenticated_services():
    # The CLIENT_SECRETS_FILE variable specifies the name of a file that contains
    # the OAuth 2.0 information for this application, including its client_id and
    # client_secret.
    # opens in web browser
    webbrowser.register('chrome',
	None,
	webbrowser.BackgroundBrowser("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))
    # webbrowser.get('chrome').open(auth_url)
    webbrowser.get(None).open(auth_url)
    
def list_events_by_guest(service, options)
    # Makes timetable for given guest (tutor)
``` 
**structures**
```
periods_dict
    # translates hour of the day into period of the day

days_dict
    # days_dict translates day of the week into locl name of the day

class Timetable
    # timeteble stores all scheduled events (pairs)
    # .print() - prints timetable to the screen
```

