#
#	uob_utils.py
#
#	These is a module that is required in some of the scripts provided.
#	It was created to provide some classes and functions to lighten various
#	scripts contained in Canvas.
#
#	First are two classes to simplifying emailing across all scripts. This
#	way, details of SMTP server and credentials only need to be changed in
#	one place for things to work.
#
#	Second are a series of functions that provide a number of useful date
#	operations which are specifically designed for the University of
#	Birmingham. The University uses a system of 'University' weeks to
#	calendarise everything from academic terms to open days. The first
#	'University' week (Week 1) is the week that contains the last Friday
#	of August. The relationship between 'University' weeks and 'Term' weeks
#	is fixed. Therefore, these function can find the relationship between
#	any date and the University calendar.
#
#	Things that need to be set:
#
#	smtp server and login credentials
smtp_server = 'mail.mottura.org' # smtp server
smtp_port = 465 # port number (ssl only, no insecure connections)
smtp_username = None # username (if this is left as None, make sure you set up ~/.mailcredentials)
smtp_password = None # password (if this is left as None, make sure you set up ~/.mailcredentials)
smtp_from_name = 'MetMat Canvas Bot' # Display name of sender
smtp_from_addr = 'canvasbots-noreply@mottura.org' # From address for all emails
#
#	SMTP server authentication is normally done using a username and a password.
#	If you do not wish to write you username and password above, you can write
#	these into a simple text file with name '.mailcredentials' and save the file
#	in your home directory. Make sure the file has restricted permissions (600).
#
#
#
#
#
#
import calendar
import datetime
import math
import smtplib
import os
import re
import sys
import json
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from copy import deepcopy


class MailAPI():
    """"""
    def __init__(self, username=smtp_username, password=smtp_password, server=smtp_server, port=smtp_port, path_to_credentials='~/.mailcredentials'):
        self.username = username
        self.password = password
        self.server = server
        self.port = port
        if username == None or password==None: # if username and password are not specified, look for .mailcredentials file
            if not os.path.isfile(os.path.expanduser(path_to_credentials)):
                raise RuntimeError('Provide a username and password in ~/.mailcredentials or as an argument of the call.')
            permissions = int(oct(os.stat(os.path.expanduser(path_to_credentials)).st_mode)[-3:])
            if permissions not in [400, 600]: # check that .mailcredentials file has appropriate file restrictions
                raise RuntimeError('Permissions of mail credentials are not secure enough.')
            with open(os.path.expanduser(path_to_credentials)) as f: # read Canvas token from .canvastoken file
                lines = f.readlines()
            self.username = lines[0].strip()
            self.password = lines[1].strip()
        try:
            self.ssl = smtplib.SMTP_SSL(self.server, self.port)
            self.ssl.ehlo()
            self.ssl.login(self.username, self.password)	
        except:
            raise RuntimeError('Could not connect to e-mail server.')
        
    def send(self, to_addr, msg):
        try: # tries to send email....
            return self.ssl.sendmail(smtp_from_addr, to_addr, msg.as_string())
        except smtplib.SMTPServerDisconnected: # ...if server is disconnected, tries to reconnect and then sends email...
            self.ssl = smtplib.SMTP_SSL(self.server, self.port)
            self.ssl.ehlo()
            self.ssl.login(self.username, self.password)
            return self.ssl.sendmail(smtp_from_addr, to_addr, msg.as_string())




class EMailMessage(MIMEMultipart):
    """"""
    def __init__(self, to_addr, subj, cc_addr=None):
        MIMEMultipart.__init__(self)
        self['From'] = '"' + smtp_from_name + '" <' + smtp_from_addr + '>'
        self['To'] = to_addr
        if not cc_addr==None:
            self['Cc'] = cc_addr
        self['Date'] = formatdate(localtime=True)
        self['Subject'] = subj
    
    def body(self, text):
        self.attach(MIMEText(text))

    def attach_file(self, path_to_file):
        with open(path_to_file, 'rb') as f:
            part = MIMEApplication(f.read(),Name=os.path.basename(path_to_file))
            part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(path_to_file)
        self.attach(part)


def AugustBankHoliday(year):
    """Returns the date of the August Bank Holiday for a given year."""
    cal = calendar.Calendar()
    month = cal.monthdatescalendar(year,8)
    lastweek = month[-1]
    monday = lastweek[0]
    return monday

def WeekOne(year):
    """Returns the Monday of week 1 for a given academic year (e.g. 2014 for 14/15)."""
    cal = calendar.Calendar(4)
    month = cal.monthdatescalendar(year,8)
    lastweek = month[-1]
    friday = lastweek[0]
    monday = friday - datetime.timedelta(days=4)
    return monday

def UniversityWeek(date):
    """Returns the university week for a given date"""
    if date >= WeekOne(date.year):
        return int(math.floor((date-WeekOne(date.year)).days/7))+1
    else:
        return int(math.floor((date-WeekOne(date.year-1)).days/7))+1

def DateFromUniversityWeek(ayear,uweek,dow):
    """Takes the academic year (e.g. 2014 for 14/15), university week and day of the week (e.g. Mon = 0, Tues = 1, etc.) as input and returns a date."""
    return WeekOne(ayear) + datetime.timedelta(weeks=uweek-1) + datetime.timedelta(days=dow)

def AcademicYear(date):
    """Returns the academic year for a given date."""
    if date >= WeekOne(date.year):
        return date.year
    else:
        return date.year - 1

def TermWeek(date):
    """Returns a list that contains the term, term week and university week for a given date."""
    uweek = UniversityWeek(date)
    if uweek <= 16 and uweek >= 5:
        return [1,uweek-5,uweek]
    elif uweek <= 31 and uweek >= 21:
        return [2,uweek-20,uweek]
    elif uweek <= 43 and uweek >= 36:
        return [3,uweek-35,uweek]
    else:
        return [0,0,uweek]

def FindCorrespondingDate(date,ayear):
    """Returns a date in any academic year that corresponds to the input date for its academic year."""
    return DateFromUniversityWeek(ayear,UniversityWeek(date),date.weekday())
    

def days_since_deadline(deadline): # returns number of days since submission deadline and the marking deadline
    DAYS_TO_MARK = 15
    today = datetime.datetime.now()
    READABLE_CLOSED_DAYS = [
        [2022, 12, 19], [2022, 12, 20], [2022, 12, 21], [2022, 12, 22], [2022, 12, 23],
        [2022, 12, 26], [2022, 12, 27], [2022, 12, 28], [2022, 12, 29], [2022, 12, 30],
        [2023, 1, 2], 
        [2023, 4, 7], [2023, 4, 10], [2023, 4, 11], [2023, 4, 12],
        [2023, 5, 1], [2023, 5, 29], [2023, 8, 28]
    ]
    closed_days = [datetime.datetime(i[0], i[1], i[2]) for i in READABLE_CLOSED_DAYS]
    closed_dates = [closed_day.date() for closed_day in closed_days] # Closed dates without times
    days_elapsed = math.ceil((today - deadline).total_seconds() / 86400)
    for i in range(1, days_elapsed + 1): # Remove weekends
        date = deadline + datetime.timedelta(days = i)
        day_of_the_week = date.weekday()
        if day_of_the_week > 4: # Sat/Sun
            days_elapsed -= 1
    for closed_day in closed_days: # Remove closed days
        if ((today - closed_day).total_seconds() > 0 and (deadline - closed_day).total_seconds() < 0):
            days_elapsed -= 1
    # Now calculate the number of working days
    working_days_from_deadline = 0
    days_from_deadline = 1
    while working_days_from_deadline < DAYS_TO_MARK:
        date = deadline + datetime.timedelta(days = days_from_deadline)
        day_of_the_week = date.weekday()
        if ((day_of_the_week < 5) and (date.date() not in closed_dates)):
            working_days_from_deadline += 1
        days_from_deadline += 1
        if days_from_deadline > 1000:
            sys.exit('Error calculating working days from deadline, exiting to break infinite loop')
    marking_deadline = deadline + datetime.timedelta(days = days_from_deadline - 1)
    is_working_day = True if ((today.weekday() < 5) and (today.date() not in closed_dates)) else False
    return days_elapsed - 1, marking_deadline, is_working_day

def produce_email(days_left, assignment, TSO_email, ws, col_sub, i):
    mail = MailAPI()
    with open('marking_messages.json', 'r') as raw_json:
        messages = json.loads(raw_json.read())
    recipients = deepcopy(TSO_email)
    if days_left in ['1 day left', '5 days left']:
        recipients.extend(re.findall('[A-Za-z0-9\.]*@bham.ac.uk',assignment['description']))
    message_subject = messages['subject'][days_left]
    message_body = messages['body'][days_left] % (assignment['name'], assignment['html_url'], ws[col_sub+str(i)].value, assignment['needs_grading_count'])
    msg = EMailMessage(", ".join(recipients), message_subject)
    msg.body(message_body)
    mail.send(recipients, msg) # Send email
    # Code below creates a text file for debug
#	with open('email.txt', 'w') as f:
#	    f.write(message_subject + "\n" + message_body)
