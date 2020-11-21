## A Python script to integrate your newly booked sessions into your Zoom and Google calendars!

#### Overview
This script uses a combination of Python IMAP email libraries and the Google Calendar API to scan your email for new booking confirmations, and then automatically generates calendar events that align with the scheduled times. These events include the client's name, the proper start and end times, and the Zoom URL that you should use to start the meeting.  

#### Known Limitations:
- HAS NOT BEEN TESTED WITH OTHER TIME ZONES. I'm in Pacific Time and I've only been able to test this on my own machine.
- Currently does not support events that span more than a single calendar day (so any appointment that involves midnight will likely need to be added manually).
- If a session booking email is not deleted manually after being processed, running the script again will cause a duplicate event to be generated in your calendar.

#### Installation & Setup
If you don't have it already, install [Python](https://www.python.org/downloads/). I would recommend using version 3.8, because while earlier versions (3.7+) should work too, they haven't been tested.

You will also need to enable external access to your email to allow the script to fetch recent emails (you can do so here: https://myaccount.google.com/lesssecureapps?pli=1). Make sure you're signed in with your iD Tech email and not your personal one!

Following that, you will also need to enable access to the Google Calendar client API. This will enable the script to create new events on your calendar. Do so at the Google Developer Dashboard here: https://console.developers.google.com/projectselector2/apis/dashboard?pli=1&supportedpurview=project.

You will also need to install the Google Client API for Python. From your Terminal or Command Prompt, run:  
`pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib`  
This will enable the actual communication of the Python code and the Google Calendar API.

Finally, you will need to generate an access token for the code to access the calendar API. The easiest way I've found to do this is to go to https://developers.google.com/calendar/quickstart/python and click <b>'Enable the Google Calendar API'</b>. Name the example project whatever you'd like, select 'Desktop App', and click create. Finally, you'll want to <b>download the client configuration</b> and <b><u>save the .json file in the same location as the script</b></u>.

#### Usage
Assuming everything is set up properly, all you need to do is run the script. To do so, navigate to wherever the script is saved in your Terminal and run `python3 autocalendar.py`. It should prompt you for your account name and password (if you get tired of typing this in, you could easily hard code this into the script itself, but that's up to you!) and open a window for you to verify your access to Google Calendar via the client API.

(If `python3` doesn't work for you, try `py` or `python`.)
