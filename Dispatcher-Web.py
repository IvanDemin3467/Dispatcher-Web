# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START gae_python38_app]
# [START gae_python3_app]
from flask import Flask, render_template, request, url_for, flash, redirect, session
import sqlite3
from werkzeug.exceptions import abort

def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

def get_post(post_id):
    conn = get_db_connection()
    post = conn.execute('SELECT * FROM posts WHERE id = ?',
                        (post_id,)).fetchone()
    conn.close()
    if post is None:
        abort(404)
    return post


from os import environ

import google.oauth2.credentials

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import *

from datetime import datetime

import webbrowser

CLIENT_SECRETS_FILE = "client_secret.json"

# This access scope grants read-only access to the authenticated user's Calendar
# account.
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
        self.timetable[week-1][day-1][period-1] = value

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
    # Outputs link to created spreadsheet on the screen
    for timetable in list_timetable:
        flash("********************\nWorking on: timetable for:", timetable.name)
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
        range_name = "Лист1!A1"
        values = timetable.get_list()
        body = {
            'values': values
        }
        value_input_option = "USER_ENTERED"
        
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=range_name,
            valueInputOption=value_input_option, body=body).execute()
        
        flash('{0} cells updated.'.format(result.get('updatedCells')))
        url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId
        flash(url)
        
        # Open in Google Chrome
        # webbrowser.get('chrome').open(url)
        webbrowser.get(None).open(url)


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

##    if open_browser:
##        webbrowser.open(auth_url, new=1, autoraise=True)
    redirect(auth_url)

##    print(authorization_prompt_message.format(url=auth_url))

    local_server.handle_request()

    # Note: using https here because oauthlib is very picky that
    # OAuth 2.0 should only occur over https.
    authorization_response = wsgi_app.last_request_uri.replace("http", "https")
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

    # init timetables
    list_timetable = []
    for tutor in tutors:
        timetable = Timetable(tutor)
        list_timetable.append(timetable)
    
    # get calendar list
    calendar_dict = get_calendar_dict(service)

    # get events
    flash("********************\nWorking on: get all events from all calendars")
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
    flash(f"Got {count_events} events for {len(list_timetable)} groups")

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

def get_options():
    # It reads parameters from file: lower_date, upper_date, group
    # Those params are to be passed to function list_events_by_param
    # That function will list events, filtered by params
    options = {}
    s = open("options.txt", "rt", encoding = "utf-8")
    stream = list(s)
    s.close()
    lower_date = stream[0].rstrip()
    options["lower_date"] = datetime.strptime(lower_date, "%Y-%m-%d %H:%M:%S")
    upper_date = stream[1].rstrip()
    options["upper_date"] = datetime.strptime(upper_date, "%Y-%m-%d %H:%M:%S")

    groups = []
    for i in range(2, len(stream)):
        nextline = stream[i].rstrip()
        groups.append(nextline)
    
    options["groups"] = groups
    return options

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


def main():
    # When running locally, disable OAuthlib's HTTPs verification. When
    # running in production *do not* leave this option enabled.
    environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
    try: service_calendar, service_sheets = get_authenticated_services()
    except BaseException as e:
        flash(e)
        flash("get_authenticated_services() FAILED")

    first_run = True
    while True:
        if first_run:
            Input = "byguest"
            first_run = False
        else:
            break
            # Input = input("Next task: ")
        try: options = get_options()
        except: flash("get_options() FAILED")
        try:
            list_timetable = list_events_by_guest(service_calendar, options)
            load_into_spreadsheet(service_sheets, list_timetable)
        except HttpError as e:
            flash(e)
            flash("/n RETRY")
            first_run = True
        except BaseException as e:
            flash("some non HttpError in main module")
            flash(e)
##        try:
##            if Input == "list" or Input == "List" or Input == "LIST":
##                list_calendar_events(service_calendar)
##            if Input == "byparam":
##                list_timetable = list_events_by_param(service_calendar, options)
##                load_into_spreadsheet(service_sheets, list_timetable)
##            if Input == "byguest":
##                list_timetable = list_events_by_guest(service_calendar, options)
##                load_into_spreadsheet(service_sheets, list_timetable)
##            if Input == "cal_list":
##                get_calendar_list()
##            if Input == "q" or Input == "quit" or Input == "Quit" or Input == "QUIT":
##                break
##            if Input == "del -all" or Input == "quit" or Input == "Quit" or Input == "QUIT":
##                if(input("Are you sure? ") == "Yes"):
##                    del_all_calendar_events(service_calendar, options)
##        except HttpError as e:
##            flash(e)
##            flash("/n RETRY")
##            first_run = True
##        except BaseException as e:
##            flash("some non HttpError in main module")
##            flash(e)
    service_calendar.close()
    service_sheets.close()
    



# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)
app.config['SECRET_KEY'] = 'dfg90845j6lk4djfglsdfglkrm345m567lksdf657lkopmndrumjfrt26kbtyi'


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

    conn = get_db_connection()
    posts = conn.execute('SELECT * FROM posts').fetchall()
    conn.close()
    return render_template('index.html', posts=posts)


@app.route('/<int:post_id>')
def post(post_id):
    post = get_post(post_id)
    return render_template('post.html', post=post)


@app.route('/create', methods=('GET', 'POST'))
def create():
    if request.method == 'POST':
        title = request.form['title']
        content = request.form['content']

        if not title:
            flash('Title is required!')
        else:
            conn = get_db_connection()
            conn.execute('INSERT INTO posts (title, content) VALUES (?, ?)',
                         (title, content))
            conn.commit()
            conn.close()
            return redirect(url_for('index'))

    return render_template('create.html')


@app.route('/<int:id>/edit', methods=('GET', 'POST'))
def edit(id):
    post = get_post(id)

    if request.method == 'POST':
        title = request.form['title']
        content = request.form['content']

        if not title:
            flash('Title is required!')
        else:
            conn = get_db_connection()
            conn.execute('UPDATE posts SET title = ?, content = ?'
                         ' WHERE id = ?',
                         (title, content, id))
            conn.commit()
            conn.close()
            return redirect(url_for('index'))

    return render_template('edit.html', post=post)


@app.route('/<int:id>/delete', methods=('POST',))
def delete(id):
    post = get_post(id)
    conn = get_db_connection()
    conn.execute('DELETE FROM posts WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash('"{}" was successfully deleted!'.format(post['title']))
    return redirect(url_for('index'))


@app.route('/make', methods=('POST', 'GET'))
def make():
    main()
    return redirect(url_for('index'))

##@app.route('/')
##def index():
##  return print_index_table()


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


@app.route('/authorize')
def authorize():
    # The CLIENT_SECRETS_FILE variable specifies the name of a file that contains
    # the OAuth 2.0 information for this application, including its client_id and
    # client_secret.


    # Do auth
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)

    # The URI created here must exactly match one of the authorized redirect URIs
    # for the OAuth 2.0 client, which you configured in the API Console. If this
    # value doesn't match an authorized URI, you will get a 'redirect_uri_mismatch'
    # error.
    flow.redirect_uri = url_for('oauth2callback', _external=True)

##    # Open in Google Chrome
##    auth_url, _ = flow.authorization_url(prompt='consent')
##
##    auth_url += "&redirect_uri=urn%3Aietf%3Awg%3Aoauth%3A2.0%3Aoob"
##    webbrowser.register('chrome',
##	None,
##	webbrowser.BackgroundBrowser("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))
##    
##    # webbrowser.get(None).open(auth_url)

    authorization_url, state = flow.authorization_url(
      # Enable offline access so that you can refresh an access token without
      # re-prompting the user for permission. Recommended for web server apps.
      access_type='offline',
      # Enable incremental authorization. Recommended as a best practice.
      include_granted_scopes='true')

    # Store the state so the callback can verify the auth server response.
    session['state'] = state

    return redirect(authorization_url)

##    flash("before credentials")
##    credentials = run_remote_server(flow, host="34.136.33.201", port=8008)
##    flash("after credentials")
##    
##    service_calendar = build('calendar', 'v3', credentials = credentials)
##    service_sheets = build('sheets', 'v4', credentials = credentials)
##    return service_calendar, service_sheets
 ############################## 



@app.route('/oauth2callback')
def oauth2callback():
    # Specify the state when creating the flow in the callback so that it can
    # verified in the authorization server response.
    state = session['state']

    flow = InstalledAppFlow.from_client_secrets_file(
      CLIENT_SECRETS_FILE, scopes=SCOPES, state=state)
    redirect_uri = url_for('oauth2callback', _external=True)
    redirect_uri = redirect_uri.replace("http", "https")
    flow.redirect_uri = 

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
    list_timetable = list_events_by_guest(service_calendar, options)
    load_into_spreadsheet(service_sheets, list_timetable)

    session['credentials'] = credentials_to_dict(credentials)
    return redirect(url_for('index'))
    #return redirect(url_for('test_api_request'))


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
    app.run(host='127.0.0.1', port=8080, debug=True)
# [END gae_python3_app]
# [END gae_python38_app]


