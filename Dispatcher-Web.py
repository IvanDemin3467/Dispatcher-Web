
from flask import Flask, render_template, request, url_for, flash, redirect, session
from werkzeug.exceptions import abort
from flask_talisman import Talisman

from os import environ

import google.oauth2.credentials

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import *

from datetime import datetime

# from waitress import serve


CLIENT_SECRETS_FILE = "client_secret.json"

# This access scope grants read-only access to the authenticated user's Calendar
# and Spreadsheets accounts.
SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/spreadsheets']
API_SERVICE_NAME = 'calendar'
API_VERSION = 'v3'

# periods_dict translates hour of the day into period of the day
periods_dict = {"08" : 1, "09" : 2, "10" : 2, "11" : 3, "12" : 3, "13" : 4,
                "14" : 4, "15" : 5, "16" : 5, "17" : 6, "18" : 7, "19" : 7}
# days_dict translates day of the week into locl name of the day
days_dict = {0: "Пн", 1: "Вт", 2: "Ср", 3: "Чт", 4: "Пт", 5: "Сб", 6: "Вс"}

# timeteble stores all scheduled events (pairs)
class Timetable:
    def __init__(self, name="Other"):
        self.timetable = [[["" for period in range(7)]
                               for day in range(7)]
                               for week in range(52)]
        self.name = name

    def put(self, value, period, day, week):
        stored_value = self.get(period, day, week)
        self.timetable[week-1][day-1][period-1] = value if stored_value == "" else "ОШИБКА! Две пары в одно время"

    def get(self, period, day, week):
        return self.timetable[week-1][day-1][period-1]

    def get_list(self):
        # .get_list() - prepares list to load into spreadsheet
        complete_list = []
        for week in range(52):
            empty_week = True
            complete_list.append(["Неделя: " + str(week+1), "", "", "", "", "", "", ""])
            complete_list.append(["", "Пара 1", "Пара 2", "Пара 3", "Пара 4", "Пара 5", "Пара 6", "Пара 7"])
            for day in range(7):
                row = []
                row.append(days_dict[day])
                for period in range(7):
                    value = self.timetable[week][day][period]
                    if value == "": value = " "
                    else: empty_week = False
                    row.append(value)
                    
                complete_list.append(row)
            if empty_week:
                del complete_list[-9:]
        return complete_list


def load_into_spreadsheet(service, list_timetable):
    # Input is: spreadsheet service, options dict and timetable object
    # Call the Sheets API, creates new spreadsheet
    # and loads data from timetable into spreadsheet
    # Outputs link to created spreadsheet

    # list of urls to ouput
    urls = []

    for timetable in list_timetable:
        flash("******************** Working on: timetable for: " + timetable.name + " ********************")

        # Create spreadsheet
        sheet = service.spreadsheets()
        spreadsheet = {
            'properties': {
                'title': timetable.name +
                datetime.now().strftime("-%Y-%m-%d-%H-%M-%S")
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                            fields='spreadsheetId').execute()
        spreadsheetId = spreadsheet.get('spreadsheetId')

        # Get into first list
        range_name = "Лист1!A1"

        # Load values from timetable in appropriate format for spreadsheet
        values = timetable.get_list()
        body = {
            'values': values
        }
        value_input_option = "USER_ENTERED"

        # Push values into spreadsheet via API
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=range_name,
            valueInputOption=value_input_option, body=body).execute()

        # How many calls updated?
        flash('{0} cells updated.'.format(result.get('updatedCells')))

        # The link to the result
        url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId
        flash(url)
        
        urls.append({timetable.name.split("@")[0]: url})
    return urls

_DEFAULT_AUTH_PROMPT_MESSAGE = (
    "Please visit this URL to authorize this application: {url}"
)
_DEFAULT_WEB_SUCCESS_MESSAGE = (
    "The authentication flow has completed. You may close this window."
)
def run_remote_server(
    flow,
    host="localhost",
    port=8080,
    authorization_prompt_message=_DEFAULT_AUTH_PROMPT_MESSAGE,
    success_message=_DEFAULT_WEB_SUCCESS_MESSAGE,
    open_browser=True,
    **kwargs
):
    """Run the flow using the server strategy.

    The server strategy instructs the user to open the authorization URL in
    their browser and will attempt to automatically open the URL for them.
    It will start a local web server to listen for the authorization
    response. Once authorization is complete the authorization server will
    redirect the user's browser to the local web server. The web server
    will get the authorization code from the response and shutdown. The
    code is then exchanged for a token.

    Args:
        host (str): The hostname for the local redirect server. This will
            be served over http, not https.
        port (int): The port for the local redirect server.
        authorization_prompt_message (str): The message to display to tell
            the user to navigate to the authorization URL.
        success_message (str): The message to display in the web browser
            the authorization flow is complete.
        open_browser (bool): Whether or not to open the authorization URL
            in the user's browser.
        kwargs: Additional keyword arguments passed through to
            :meth:`authorization_url`.

    Returns:
        google.oauth2.credentials.Credentials: The OAuth 2.0 credentials
            for the user.
    """

    
    wsgi_app = _RedirectWSGIApp(success_message)
    local_server = wsgiref.simple_server.make_server(
        host, port, wsgi_app, handler_class=_WSGIRequestHandler
    )

    flow.redirect_uri = "http://{}:{}/".format(host, local_server.server_port)
    auth_url, _ = flow.authorization_url(**kwargs)

    redirect(auth_url)

    local_server.handle_request()

    # Note: using https here because oauthlib is very picky that
    # OAuth 2.0 should only occur over https.
    authorization_response = wsgi_app.last_request_uri.replace("http://", "https://")
    flow.fetch_token(authorization_response=authorization_response)

    return flow.credentials


def get_authenticated_services():
    # The CLIENT_SECRETS_FILE variable specifies the name of a file that contains
    # the OAuth 2.0 information for this application, including its client_id and
    # client_secret.
    CLIENT_SECRETS_FILE = "client_secret.json"

    # This access scope grants read-only access to the authenticated user's Calendar
    # account.
    SCOPES = ['https://www.googleapis.com/auth/calendar',
              'https://www.googleapis.com/auth/spreadsheets']
    API_SERVICE_NAME = 'calendar'
    API_VERSION = 'v3'

    # Do auth
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)

    # Open in Google Chrome
    auth_url, _ = flow.authorization_url(prompt='consent')
    auth_url += "&redirect_uri=urn%3Aietf%3Awg%3Aoauth%3A2.0%3Aoob"
    webbrowser.register('chrome',
	None,
	webbrowser.BackgroundBrowser("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))
    
    # webbrowser.get(None).open(auth_url)

    flash("before credentials")
    credentials = run_remote_server(flow, host="34.136.33.201", port=8008)
    flash("after credentials")
    
    service_calendar = build('calendar', 'v3', credentials = credentials)
    service_sheets = build('sheets', 'v4', credentials = credentials)
    return service_calendar, service_sheets


def list_events_by_guest(service, options):
    # Same as byparam, but makes timetable for given guest (tutor)
    # lists all events in main calendar which are between two dates, given in options
    # Also it filters events by name of guest (tutor)
    page_token = None

    try: tutors = get_tutors()
    except: flash("get_tutors() FAILED")

    tutors = tutors_input.split("\r\n")

    # init timetables
    list_timetable = []
    for tutor in tutors:
        timetable = Timetable(tutor)
        list_timetable.append(timetable)
    
    # get calendar list
    calendar_dict = get_calendar_dict(service)

    # get events
    flash("******************** Working on: get all events from all calendars ********************")
    count_events = 0
    for calendar in calendar_dict:
        while True:
            events = service.events().list(calendarId=calendar, pageToken=page_token, singleEvents = True).execute()
            for event in events['items']:
                try:
                    event_name = event['summary']
                    event_start = datetime.strptime(event['start']['dateTime'],
                                                    "%Y-%m-%dT%H:%M:%S+06:00")
                    for timetable in list_timetable:
                        if event_start > options["lower_date"] and \
                           event_start < options["upper_date"]:
                            attendees = event['attendees']
                            for attendee in attendees:
                                if attendee['email'] == timetable.name:
                                    count_events += 1
                                    event_tutor = calendar_dict[calendar]
                                    week = int(event_start.strftime("%W").lstrip("0"))
                                    week = week - options["first_week"]
                                    day = int(event_start.strftime("%w"))
                                    period = event_start.strftime("%H")
                                    try: period = periods_dict[period]
                                    except: period = "other"
                                    data = event_name
                                    #print(data)
                                    try: timetable.put(value=data, period=period, day=day, week=week)
                                    except: print("error while put()")
                except:
                    pass

            page_token = events.get('nextPageToken')
            if not page_token:
                break
    flash(f"Got {count_events} events for {len(list_timetable)} tutors")

    return list_timetable

def get_tutors():
    # It reads parameters from file: 
    # Those params are to be passed to function list_events_by_guest
    # That function will list events, filtered by params
    options = {}
    s = open("tutors.txt", "rt", encoding = "utf-8")
    stream = list(s)
    s.close()
    tutors_list = []
    for i in range(0, len(stream)):
        nextline = stream[i].rstrip()
        tutors_list.append(nextline)
    return tutors_list

def get_options(path = "options.txt"):
    # It reads parameters from file: lower_date, upper_date, group
    # Those params are to be passed to function list_events_by_param
    # That function will list events, filtered by params
    options = {}
    s = open(path, "rt", encoding = "utf-8")
    stream = list(s)
    s.close()
    lower_date = stream[0].rstrip()
    options["lower_date"] = datetime.strptime(lower_date, "%Y-%m-%d %H:%M:%S")
    upper_date = stream[1].rstrip()
    options["upper_date"] = datetime.strptime(upper_date, "%Y-%m-%d %H:%M:%S")
    first_week = stream[2].rstrip()
    options["first_week"] = int(first_week)

    groups = []
    for i in range(3, len(stream)):
        nextline = stream[i].rstrip()
        groups.append(nextline)
    
    options["groups"] = groups
    return options

def set_options(options, path = "options.txt"):
    # It writes parameters to file: lower_date, upper_date, first_week
    try:
        s = open(path, "wt", encoding = "utf-8")
        
        s.write(str(options["lower_date"]) + "\n")
        s.write(str(options["upper_date"]) + "\n")
        s.write(str(options["first_week"]))

        s.close()
    except:
        result = "FAILED writing options"
        flash(result)
    else:
        result = "SUCCESS writing options"
    return result

def load_default_options():
    # It writes parameters to file: lower_date, upper_date, first_week
    try:
        default_options = get_options("default-options.txt")
    except:
        result = "FAILED reading default options"
        flash(result)
    else:
        result = "SUCCESS reading default options"

    try:
        set_options(default_options)
    except:
        result = "FAILED writing default options"
        flash(result)
    return result

def get_calendar_dict(service):
    # This function retrieves list of calendars for user
    # Returns dict if calendar_ID: calenar_summary
    page_token = None
    calendar_dict = {}
    while True:
      calendar_list = service.calendarList().list(pageToken=page_token).execute()
      for calendar_list_entry in calendar_list['items']:
        #print(calendar_list_entry['summary'])
          calendar_dict[calendar_list_entry['id']] = calendar_list_entry['summary']
      page_token = calendar_list.get('nextPageToken')
      if not page_token:
        break
    return calendar_dict




# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)
#talisman = Talisman(app)
app.config['SECRET_KEY'] = 'dfg90845j6lk4djfglsdfglkrm345m567lksdf657lkopmndrumjfrt26kbtyi'
urls_output = []
urls = []

@app.route('/', methods=['GET', 'POST'])
def index():
    i = 1
    if request.method == 'POST':
        if request.form.get('submit_button') == 'Do main job':
            pass
        else:
            pass # unknown
    elif request.method == 'GET':
        pass


    #get tutors list from file
    try: tutors = get_tutors()
    except: pass
    
    # put tutors into text area
    init_tutors = ""
    for tutor in tutors:
        init_tutors += tutor + "\n"
    init_tutors = init_tutors[:-1]
    
    return render_template('index.html',
                           init_tutors=init_tutors,
                           urls_output=urls,
                           urls=urls)


@app.route('/options', methods=('GET', 'POST'))
def options():
    options=get_options()

    return render_template('options.html',
                           lower_date=options["lower_date"],
                           upper_date=options["upper_date"],
                           first_week=options["first_week"])


@app.route('/edit_options', methods=('GET', 'POST'))
def edit_options():
    options=get_options()

    options["lower_date"] = request.form['lower_date']
    options["upper_date"] = request.form['upper_date']
    options["first_week"] = request.form['first_week']
    set_options(options)

    return redirect('options')

@app.route('/set_default_options', methods=('GET', 'POST'))
def set_default_options():
    load_default_options()

    return redirect('options')



@app.route('/test')
def test_api_request():
  if 'credentials' not in session:
    return redirect('authorize')

  # Load credentials from the session.
  credentials = google.oauth2.credentials.Credentials(
      **session['credentials'])

  youtube = googleapiclient.discovery.build(
      API_SERVICE_NAME, API_VERSION, credentials=credentials)

  channel = youtube.channels().list(mine=True, part='snippet').execute()

  # Save credentials back to session in case access token was refreshed.
  # ACTION ITEM: In a production app, you likely want to save these
  #              credentials in a persistent database instead.
  session['credentials'] = credentials_to_dict(credentials)

  return jsonify(**channel)


@app.route('/authorize', methods=['POST'])
def authorize():
    # The CLIENT_SECRETS_FILE variable specifies the name of a file that contains
    # the OAuth 2.0 information for this application, including its client_id and
    # client_secret.

    global tutors_input
    tutors_input = request.form['tutors_input']

    # Do auth
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)

    # The URI created here must exactly match one of the authorized redirect URIs
    # for the OAuth 2.0 client, which you configured in the API Console. If this
    # value doesn't match an authorized URI, you will get a 'redirect_uri_mismatch'
    # error.
    flow.redirect_uri = url_for('oauth2callback', _external=True)

    authorization_url, state = flow.authorization_url(
      # Enable offline access so that you can refresh an access token without
      # re-prompting the user for permission. Recommended for web server apps.
      access_type='offline',
      # Enable incremental authorization. Recommended as a best practice.
      include_granted_scopes='true')

    # Store the state so the callback can verify the auth server response.
    session['state'] = state

    return redirect(authorization_url)



@app.route('/oauth2callback')
def oauth2callback():
    # Specify the state when creating the flow in the callback so that it can
    # verified in the authorization server response.
    environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
    state = session['state']

    flow = InstalledAppFlow.from_client_secrets_file(
      CLIENT_SECRETS_FILE, scopes=SCOPES, state=state)
    redirect_uri = url_for('oauth2callback', _external=True)
    flow.redirect_uri = redirect_uri

    # Use the authorization server's response to fetch the OAuth 2.0 tokens.
    authorization_response = request.url
    # Note: using https here because oauthlib is very picky that
    # OAuth 2.0 should only occur over https.
    authorization_response = authorization_response.replace("http", "https")
    flow.fetch_token(authorization_response=authorization_response)

    # Store credentials in the session.
    # ACTION ITEM: In a production app, you likely want to save these
    #              credentials in a persistent database instead.
    credentials = flow.credentials

    service_calendar = build('calendar', 'v3', credentials = credentials)
    service_sheets = build('sheets', 'v4', credentials = credentials)
    try: options = get_options()
    except: flash("get_options() FAILED")

    #main job is here
    list_timetable = list_events_by_guest(service_calendar, options)
    global urls
    urls = load_into_spreadsheet(service_sheets, list_timetable)
    return redirect(url_for('index'))


@app.route('/revoke')
def revoke():
  if 'credentials' not in session:
    return ('You need to <a href="/authorize">authorize</a> before ' +
            'testing the code to revoke credentials.')

  credentials = google.oauth2.credentials.Credentials(
    **session['credentials'])

  revoke = requests.post('https://oauth2.googleapis.com/revoke',
      params={'token': credentials.token},
      headers = {'content-type': 'application/x-www-form-urlencoded'})

  status_code = getattr(revoke, 'status_code')
  if status_code == 200:
    return('Credentials successfully revoked.' + print_index_table())
  else:
    return('An error occurred.' + print_index_table())


@app.route('/clear')
def clear_credentials():
  if 'credentials' in session:
    del session['credentials']
  return ('Credentials have been cleared.<br><br>' +
          print_index_table())


def credentials_to_dict(credentials):
  return {'token': credentials.token,
          'refresh_token': credentials.refresh_token,
          'token_uri': credentials.token_uri,
          'client_id': credentials.client_id,
          'client_secret': credentials.client_secret,
          'scopes': credentials.scopes}

##def print_index_table():
##  return ('<table>' +
##          '<tr><td><a href="/test">Test an API request</a></td>' +
##          '<td>Submit an API request and see a formatted JSON response. ' +
##          '    Go through the authorization flow if there are no stored ' +
##          '    credentials for the user.</td></tr>' +
##          '<tr><td><a href="/authorize">Test the auth flow directly</a></td>' +
##          '<td>Go directly to the authorization flow. If there are stored ' +
##          '    credentials, you still might not be prompted to reauthorize ' +
##          '    the application.</td></tr>' +
##          '<tr><td><a href="/revoke">Revoke current credentials</a></td>' +
##          '<td>Revoke the access token associated with the current user ' +
##          '    session. After revoking credentials, if you go to the test ' +
##          '    page, you should see an <code>invalid_grant</code> error.' +
##          '</td></tr>' +
##          '<tr><td><a href="/clear">Clear Flask session credentials</a></td>' +
##          '<td>Clear the access token currently stored in the user session. ' +
##          '    After clearing the token, if you <a href="/test">test the ' +
##          '    API request</a> again, you should go back to the auth flow.' +
##          '</td></tr></table>')


if __name__ == '__main__':
    # This is used when running locally only. When deploying to Google App
    # Engine, a webserver process such as Gunicorn will serve the app. This
    # can be configured by adding an `entrypoint` to app.yaml.

    #app.run(host="0.0.0.0", port=5000, ssl_context=("certificate.pem", "key.pem"))
    app.run(host="127.0.0.1", port=5000, ssl_context=("certificate.pem", "key.pem"))
    #serve(app, host='127.0.0.1', port=5000)
# [END gae_python3_app]
# [END gae_python38_app]
    #app.run(host="0.0.0.0", port=5000, ssl_context="adhoc")


