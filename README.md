# Dispatcher-Web
Application for the dispatcher to work with the Google Calendar via Web API.
Now it works as web app.

**Input files: options.txt**

It uses parameters given in file "options.txt". 

The first line - first parameter is the start date. All events *preceding* it will be omitted.

The second line - second parameter is the end date. All events *following* it will be omitted.

The third and subsequent lines - third parameter is the list of strings to find (e. g. names of study groups). In only one string to find: All events without this string in caption will be omitted. If many lines - many strings to find - executes program many times making many output files.

*Example input in options.txt*
```
2021-01-01 08:00:00
2021-12-01 00:00:00
```

**Input files: tutors.txt**

This file contais list of guests (tutors) to find

*Example input in options.txt*
```
deminie@college.omsu.ru
dan-909-o@college.omsu.ru
dtn-809-o@college.omsu.ru
```

**Output files**

Program creates google spreadsheet with timetable for given study group. If got many groups on input - it creates many timetables, one per group.

**uses**
``` 
google.oauth2.credentials
googleapiclient.discovery
googleapiclient.errors
google_auth_oauthlib.flow

os
datetime
webbrowser
``` 

Uses googleapiclient library from
https://github.com/googleapis/google-api-python-client

**Requres**
```
pip install google-api-python-client
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

**Features**
```
## default behavior - silently run byparam
list - lists all events in main calendar
quit - quits program
del -all - deletes all events in main calendar
           Also searches through all calendars for given user
byparam - lists all events in main calendar which are between dates, given in options
          Also it filters events by name. Choses events that contain give string in name
          It is implemented to get all events for given group e.g. DAN-909, DTN-809 etc.
          Also searches all calendars for given user
          Also accepts cyrillic letters
          Pretty prints timetable if there is only one group in input
          Ommits pretty print if there are many groups on input
          Creates google spreadsheet and loads timetable into it
          Shows link to the created sheet on a screen
byguest - Same as byparam, but makes timetable for given guest (tutor)  
	  Also lists all events in main calendar which are between two dates, given in options
cal_list - lists calendars for user
```
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
**usefull parameters for query**
```
Events: list
calendarId - string from calendarList: List
singleEvents = True - to expand recurring events into instances. False - to only return single one-off events and instances of recurring events, but not the underlying recurring events themselves. Optional. The default is False. 
```
