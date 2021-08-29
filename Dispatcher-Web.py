import google.oauth2.credentials
from datetime import datetime
from flask import Flask, render_template, request, url_for, flash, redirect, session
from google_auth_oauthlib.flow import *
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from os import environ
from werkzeug.exceptions import abort

CLIENT_SECRETS_FILE = "client_secret.json"

# This access scope grants read-only access to the authenticated user's Calendar  and Spreadsheets accounts.
SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/spreadsheets']

# periods_dict translates hour of the day into period of the day
periods_dict = {"08": 1, "09": 2, "10": 2, "11": 3, "12": 3, "13": 4,
                "14": 4, "15": 5, "16": 5, "17": 6, "18": 7, "19": 7}
# days_dict translates day of the week into local name of the day
days_dict = {0: "Пн", 1: "Вт", 2: "Ср", 3: "Чт", 4: "Пт", 5: "Сб", 6: "Вс"}


# timetable stores all scheduled events (pairs)
class Timetable:
    def __init__(self, name="Other"):
        """
        Initializes list of events as 3-d array
        Also stores name of timetable
        """
        self.timetable = [[["" for period in range(7)]
                           for day in range(7)]
                          for week in range(53)]
        self.name = name

    def put(self, value, period, day, week):
        """
        Store the given value in the timetable.
        Notices the error if there are two events in the same period
        """
        stored_value = self.get(period, day, week)
        error_value = "ОШИБКА! Две пары в одно время: " + stored_value + "; " + value
        self.timetable[week - 1][day - 1][period - 1] = value if stored_value == "" else error_value

    def get(self, period, day, week):
        return self.timetable[week - 1][day - 1][period - 1]

    def get_list(self):
        """
        It prepares list of cells to load into spreadsheet
        """
        complete_list = []
        for week in range(53):
            empty_week = True
            complete_list.append(["Неделя: " + str(week + 1), "", "", "", "", "", "", ""])
            complete_list.append(["", "Пара 1", "Пара 2", "Пара 3", "Пара 4", "Пара 5", "Пара 6", "Пара 7"])
            for day in range(7):
                row = [days_dict[day]]
                for period in range(7):
                    value = self.timetable[week][day][period]
                    if value == "":
                        value = " "
                    else:
                        empty_week = False
                    row.append(value)

                complete_list.append(row)
            if empty_week: del complete_list[-9:]  # Do not print empty weeks
        return complete_list


def load_into_spreadsheet(service, list_timetable):
    """
    Input is: spreadsheet service and list of timetable objects
    It calls the Sheets API, creates new spreadsheets and loads data from timetables into spreadsheets
    Outputs list of links to created spreadsheets
    """

    # list of urls to output
    global urls
    urls = []

    for timetable in list_timetable:
        flash("******************** Working on: timetable for: " + timetable.name + " ********************")

        # Create spreadsheet
        spreadsheet = {
            'properties': {
                'title': timetable.name + datetime.now().strftime("-%Y-%m-%d-%H-%M-%S")
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        # Get into the first page in a spreadsheet
        range_name = "Лист1!A1"

        # Load values from timetable in appropriate format for spreadsheet
        values = timetable.get_list()
        body = {
            'values': values
        }
        value_input_option = "USER_ENTERED"

        # Push values into spreadsheet via API
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption=value_input_option, body=body).execute()

        # How many cells updated?
        flash('{0} cells updated.'.format(result.get('updatedCells')))

        # The link to the result
        url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id
        flash(url)

        # prepare the output
        urls.append({timetable.name.split("@")[0]: url})
    return urls


def list_events_by_guest(service, query):
    """
    Input is Google Calendar service and parameters of query
    It makes API call to Calendar , gets all events from all calendars and filters them according to query
    """
    # Same as byparam, but makes timetable for given guest (tutor)
    # lists all events in main calendar which are between two dates, given in options
    # Also it filters events by name of guest (tutor)
    page_token = None

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
            events = service.events().list(calendarId=calendar, pageToken=page_token, singleEvents=True).execute()
            for event in events['items']:
                try:
                    event_name = event['summary']
                    event_start = datetime.strptime(event['start']['dateTime'], "%Y-%m-%dT%H:%M:%S+06:00")
                except:
                    break
                for timetable in list_timetable:
                    if query["lower_date"] < event_start < query["upper_date"]:
                        attendees = event['attendees']
                        for attendee in attendees:
                            if attendee['email'] == timetable.name:
                                count_events += 1
                                try:
                                    week = int(event_start.strftime("%W").lstrip("0"))
                                    week = week - query["first_week"]
                                except:
                                    week = 0
                                    flash(timetable.name + ". Error with week")
                                try:
                                    day = int(event_start.strftime("%w"))
                                except:
                                    day = 0
                                    flash(timetable.name + ". Error with day")
                                try:
                                    period = event_start.strftime("%H")
                                    period = periods_dict[period]
                                except:
                                    period = 7
                                    flash(timetable.name + ". Error with period")
                                try:
                                    timetable.put(value=event_name, period=period, day=day, week=week)
                                except:
                                    flash(timetable.name + ". Error while put()")

            page_token = events.get('nextPageToken')
            if not page_token:
                break
    flash(f"Got {count_events} events for {len(list_timetable)} tutors")

    return list_timetable


def get_tutors():
    # It reads parameters from file: 
    # Those params are to be passed to function list_events_by_guest
    # That function will list events, filtered by params
    # options = {}
    s = open("tutors.txt", "rt", encoding="utf-8")
    stream = list(s)
    s.close()
    tutors_list = []
    for i in range(0, len(stream)):
        nextline = stream[i].rstrip()
        tutors_list.append(nextline)
    return tutors_list


def get_options(path="options.txt"):
    # It reads parameters from file: lower_date, upper_date, group
    # Those params are to be passed to function list_events_by_param
    # That function will list events, filtered by params
    options = {}
    s = open(path, "rt", encoding="utf-8")
    stream = list(s)
    s.close()
    lower_date = stream[0].rstrip()
    options["lower_date"] = datetime.strptime(lower_date, "%Y-%m-%d %H:%M:%S")
    upper_date = stream[1].rstrip()
    options["upper_date"] = datetime.strptime(upper_date, "%Y-%m-%d %H:%M:%S")
    first_week = stream[2].rstrip()
    options["first_week"] = int(first_week)

    return options


def set_options(options, path="options.txt"):
    # It writes parameters to file: lower_date, upper_date, first_week
    try:
        s = open(path, "wt", encoding="utf-8")

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
    default_options = {}
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
            # print(calendar_list_entry['summary'])
            calendar_dict[calendar_list_entry['id']] = calendar_list_entry['summary']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    return calendar_dict


# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)
# talisman = Talisman(app)
app.config['SECRET_KEY'] = 'dfg90845j6lk4djfglsdfglkrm345m567lksdf657lkopmndrumjfrt26kbtyi'
urls = []


@app.route('/', methods=['GET', 'POST'])
def index():
    i = 1
    if request.method == 'POST':
        if request.form.get('submit_button') == 'Do main job':
            pass
        else:
            pass  # unknown
    elif request.method == 'GET':
        pass

    # get tutors list from file
    tutors = []
    try:
        tutors = get_tutors()
    except:
        flash("FAILED get_tutors()")

    # put tutors into text area
    init_tutors = ""
    for tutor in tutors:
        init_tutors += tutor + "\n"
    init_tutors = init_tutors[:-1]

    return render_template('index.html',
                           init_tutors=init_tutors,
                           urls=urls)


@app.route('/options', methods=('GET', 'POST'))
def options():
    options = get_options()

    return render_template('options.html',
                           lower_date=options["lower_date"],
                           upper_date=options["upper_date"],
                           first_week=options["first_week"])


@app.route('/edit_options', methods=('GET', 'POST'))
def edit_options():
    options = get_options()

    options["lower_date"] = request.form['lower_date']
    options["upper_date"] = request.form['upper_date']
    options["first_week"] = request.form['first_week']
    set_options(options)

    return redirect('options')


@app.route('/set_default_options', methods=('GET', 'POST'))
def set_default_options():
    load_default_options()

    return redirect('options')


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

    service_calendar = build('calendar', 'v3', credentials=credentials)
    service_sheets = build('sheets', 'v4', credentials=credentials)

    options = {}
    try:
        options = get_options()
    except:
        flash("get_options() FAILED")

    # main job is here
    list_timetable = list_events_by_guest(service_calendar, options)
    global urls
    urls = load_into_spreadsheet(service_sheets, list_timetable)
    return redirect(url_for('index'))


if __name__ == '__main__':
    # This is used when running locally only. When deploying to Google App
    # Engine, a webserver process such as Gunicorn will serve the app. This
    # can be configured by adding an `entrypoint` to app.yaml.

    # app.run(host="0.0.0.0", port=5000, ssl_context=("certificate.pem", "key.pem"))
    app.run(host="127.0.0.1", port=5000, ssl_context=("certificate.pem", "key.pem"))
    # serve(app, host='127.0.0.1', port=5000)
