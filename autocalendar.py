from __future__ import print_function
import datetime, pickle, os, os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import imaplib, email, webbrowser
from email.header import decode_header
from getpass import getpass

# see https://www.thepythoncode.com/article/reading-emails-in-python for more info

# Prerequisite steps: enable the Google Developer Console for your iD Tech Account
# to enable Google Calendar Synchronization. Do so here:
# https://console.developers.google.com/projectselector2/apis/dashboard?pli=1&supportedpurview=project

# You must also enable 'less secure apps' to allow the script to read your emails. You can do so here:
# https://myaccount.google.com/lesssecureapps?pli=1

# Set up an example Quickstart API file here:
# https://developers.google.com/calendar/quickstart/python.
# Select 'Desktop App' and 'Download Client Configuration'.
# Save this file in the same folder as this script.

# Install the Google Client APIs using Python: (run this in your command prompt/terminal)
# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

# Also important! Delete your lesson confirmation emails once you receive them.
# This way, the program won't add multiple events over one another.

# US Time Zones:
HST = "-10:00" # Hawaiian Standard Time
ADT = "-08:00" # Alaskan Daylight Time
PDT = "-07:00" # Pacific Daylight Time
MST = "-07:00" # Mountain Standard Time
MDT = "-06:00" # Mountain Daylight Time
CDT = "-05:00" # Central Daylight Time
EDT = "-04:00" # Eastern Daylight Time

AST = "-09:00" # Alaskan Standard Time
PST = "-08:00" # Pacific Standard Time
MST = "-07:00" # Mountain Standard Time
CST = "-06:00" # Central Standard Time
EST = "-05:00" # Eastern Standard Time

# Your time zone: (default is Pacific Standard Time, UTC -8). NOT TESTED elsewhere.
default_timezone = PST
timezone_reference = 'America/Los_Angeles'


# SIGN-IN INFO
username = input("Your email: ")
if (not username.endswith("@idtech.com")):
    username += "@idtech.com"
    # print(f"Interpreting your email as {username}...")

password = getpass(f'Password for {username}: ')
tagline = "iD Tech Online - New Lesson Scheduled with" # look for this in subject line

# Number of recent emails to fetch; default is 20, but feel free to increase if
# you have a lot of new appointments to add.
# N = 20
try:
    N = int(input("How far back should I look for appointments? (default is 20)\n"))
except:
    print("Error: invalid num. Defaulting to 20.")
    N = 20
print("="*100)

def main():
    imap = imaplib.IMAP4_SSL("imap.gmail.com") # create an IMAP4 class with SSL
    imap.login(username, password) # authenticate
    status, messages = imap.select("INBOX")
    messages = int(messages[0]) # total number of emails

    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")

        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode()
                # email sender
                from_ = msg.get("From")


                # skip irrelevant emails:
                if tagline not in subject:
                    continue

                # print("Subject:", subject)
                # print("From:", from_)

                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # could print text/plain emails and skip attachments here
                            pass
                        elif "attachment" in content_disposition:
                            # print(body)
                            pass # ignore emails with attachments
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()

                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    maintype = msg.get_content_maintype()
                    if content_type == "text/plain":
                        # uncomment to print plain text emails
                        # print(body)
                        pass

                    if maintype == 'text':
                        initial_cut = body.split("A new lesson for ")[1]
                        second_cut = initial_cut.split(" has been scheduled for ")
                        third_cut = second_cut[1].split(" accessed via this link: ")
                        # fourth_cut = third_cut[1].split(".&nbsp")[0]

                        clientname = second_cut[0]
                        longdate = " ".join(third_cut[0].split(" ")[1::]) # ignore day of week
                        longdate_split = longdate.split(" ")[0::1]

                        date = f"{longdate_split[0]} {longdate_split[1]} {longdate_split[2][:-1]}"
                        # time = f"{longdate_split[3]} {longdate_split[4]}"
                        # time = " ".join(longdate_split[3::])
                        time = f"{longdate_split[3]} {longdate_split[4][:-1]}"
                        timezone = " ".join(longdate_split[5::])
                        zoomurl = third_cut[1].split(".&nbsp")[0]

                        print(f"CLIENT: {clientname}\nDATE: {date}\nTIME: {time}, {timezone}\nZOOM URL: {zoomurl}")
                        appt = create_event(clientname, date, time, timezone, zoomurl)

                        insert_event(appt)
                print("="*100)
    imap.close()
    imap.logout()

def convert_date(date):
    # takes in a 'date' string in format 'Month dd, yyyy' and returns a 'YYYY-MM-DD' string
    newdate = f"{date[-4:]}-"
    newdate += "{:02d}-".format(month_string_to_number(date.split(" ")[0]))
    newdate += "{:02d}".format(int(date.split(" ")[1][:-1]))
    return newdate

def month_string_to_number(string):
    # source: https://stackoverflow.com/questions/3418050/month-name-to-month-number-and-vice-versa-in-python
    m = {
        'jan':1,
        'feb':2,
        'mar':3,
        'apr':4,
         'may':5,
         'jun':6,
         'jul':7,
         'aug':8,
         'sep':9,
         'oct':10,
         'nov':11,
         'dec':12
        }
    s = string.strip()[:3].lower()

    try:
        out = m[s]
        return out
    except:
        raise ValueError('Not a month')

def conv_to_24hr(time): # convert a string in HH:MM AM/PM format to 24-hr HH:MM format
    times = time.split(":")
    times[0] = int(times[0])
    times[1] = int(times[1].split(" ")[0])

    if "pm" in time.lower() and times[0] != 12:
        times[0] += 12
    elif "pm" not in time.lower() and times[0] == 12:
        times[0] = 0

    return f"{times[0]:02d}:{times[1]:02d}"

def get_end_time(time):
    # input time is 24hr version, returns a 24-hr version of the end time
    times = time.split(":")
    times[0] = int(times[0]) + 1

    if times[0] >= 24:
        print("Error: Time overlaps two separate days. This functionality has not yet been implemented. Please add this session manually.")
        return None

    return f"{times[0]}:{times[1]}"

def create_event(client, date, time, timezone, url):
    starttime = f"{convert_date(date)}T{conv_to_24hr(time)}:00{default_timezone}"
    endtime = f"{convert_date(date)}T{get_end_time(conv_to_24hr(time))}:00{default_timezone}"

    event = {
      'summary': f"{client}: OPL",
      'location': url,
      'description': 'This event was automatically generated by autocalendar.py.',
      'start': {
        'dateTime': starttime,
        'timeZone': timezone_reference,
      },
      'end': {
        'dateTime': endtime,
        'timeZone': timezone_reference,
      },
      'recurrence': [],
      'attendees': [],
      'reminders': {
        'useDefault': True,
        'overrides': [],
      },
    }
    return event

def insert_event(event):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    event = service.events().insert(calendarId='primary', body=event).execute()
    print("Event created.")
    return

if __name__ == "__main__":
    main()
