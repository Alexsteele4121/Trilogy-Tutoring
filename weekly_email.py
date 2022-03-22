# To install all the requirements: pip install -r requirements.txt
import gspread
from google.oauth2.service_account import Credentials
from termcolor import cprint
import threading
from email.message import EmailMessage
import smtplib
from time import sleep
# To install all the requirements: pip install -r requirements.txt


# Google Doc Info
tutor_tracking_spreadsheet = '<NAME OF YOUR TUTOR TRACKING SPREADSHEET>'
student_roster_worksheet = 'Student Roster'
student_roster_cells_range = 'A3:G'

# Emails You Would Like To Prevent From Receiving Emails
blacklist = ['Example@gmail.com']

# Google Doc Credential Info
api_credentials_file = r'<PATH\TO\YOUR\API\CREDS>'  # Youtube video to help setup credentials: https://youtu.be/cnPlKLEGR7E
api_scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Gmail Login Info
gmail_username = '<YOUR EMAIL>@instructors.2u.com'
gmail_app_password = '<YOUR APP KEY>'  # You can create an app key here: https://myaccount.google.com/security

# Weekly Message
full_name = '<YOUR NAME>'
calendly_link = '<YOUR CALENDLY LINK>'
weekly_subject = 'Cyber Boot Camp - Tutorial Available'
weekly_message = f'''Hi Everyone!

I hope you had a great week! I have attached a link below to schedule another tutoring session if you wish. If you are already scheduled, please ignore this email.

{calendly_link}
On the Calendly page, be sure you have the correct time zone selected in the section labeled "Times are in" 
If our availability doesn’t sync, let me know and I'll see if we can figure something out.

Maximum tutorial sessions per week - our week is Monday - Sunday.
    - Part-time (6 month boot camp) students are entitled to 1 session per week.
    - Full-time (3 month boot camp) students are entitled to 2 sessions per week. 

If you have any questions or none of the times available work for you please let me know and I would be happy to help.

If you would like to schedule regular, recurring sessions at the same day/time each week, just let me know by REPLY ALL and we can work it out.  This is particularly useful if you have a strict schedule so you won't have to compete for time on my calendar.

CC Central Support on all email by always using REPLY ALL.

Sincerely,
{full_name}'''

# Nothing To Fill Out Past This Point.

class Gmail:

    def __init__(self, username, password):
        self.username = username
        self.password = password

    def email(self, Receiver: str, Message: str, Subject: str = None, From: str = None, Cc: str = None):
        msg = EmailMessage()
        msg['Subject'] = Subject
        msg['From'] = From
        msg['To'] = Receiver
        msg['Cc'] = Cc
        msg.set_content(Message)

        threading.Thread(target=self.send, args=(msg,)).start()

    def send(self, msg):
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.username, self.password)
                smtp.send_message(msg)
            return True
        except Exception as e:
            print(e)
            return False


class Student:

    def __init__(self, class_code: str, graduation_date: str, student_name: str, student_email: str,
                 timezone_difference: str):
        self.class_code = class_code
        self.graduation_date = graduation_date
        self.name = student_name
        self.email = student_email
        self.timezone_difference = timezone_difference

    def __str__(self):
        return f'{self.name}: {self.email}'


class Roster:

    def __init__(self, spreadsheet: str, worksheet: str, cells_range: str, creds_file: str, scopes: list):
        self.spreadsheet = spreadsheet
        self.worksheet = worksheet
        self.cells_range = cells_range
        self.creds_file = creds_file
        self.scopes = scopes

        self.students = self.retrieve_roster()

    def authenticate(self):
        credentials = Credentials.from_service_account_file(
            filename=self.creds_file,
            scopes=self.scopes
        )
        return gspread.authorize(credentials)

    def retrieve_roster(self) -> list:
        gc = self.authenticate()
        sheet = gc.open(self.spreadsheet).worksheet(self.worksheet)
        raw_data = sheet.get(self.cells_range)

        students = []
        for data in raw_data:
            students.append(Student(class_code=data[0],
                                    graduation_date=data[1],
                                    student_name=data[2],
                                    student_email=data[3],
                                    timezone_difference=data[5]))
        return students

    def get_student_count(self) -> int:
        return len(self.students)

    def get_all_students(self) -> list:
        return self.students

    def get_all_emails(self) -> list:
        return [student.email for student in self.students]


def get_roster():
    return Roster(spreadsheet=tutor_tracking_spreadsheet,
                  worksheet=student_roster_worksheet,
                  cells_range=student_roster_cells_range,
                  creds_file=api_credentials_file,
                  scopes=api_scopes)


def main():
    cprint('[-] Loading Roster...', 'yellow')
    try:
        database = get_roster()
    except gspread.SpreadsheetNotFound:
        cprint(f'[!] Could Not Load Your Roster.\n'
               f'[!] Please Ensure You Have Shared The Spreadsheet With Your Service Account.', 'red')
        exit(1)
    except gspread.exceptions.WorksheetNotFound:
        cprint(f'[!] Could Not Load Your Worksheet.\n'
               f'[!] The Spreadsheet Was Accessible But No Worksheet Named "{student_roster_worksheet}" Was Found.',
               'red')
        exit(2)
    except FileNotFoundError as e:
        cprint(f'[!] Could Not Find Spreadsheet Credential File. {e}\n'
               f'[!] Please Activate/Download Credentials: https://console.cloud.google.com', 'red')
        exit(3)
    except Exception as e:
        cprint(f'[!] Fatal Error When Loading Roster. {e}', 'red')
        exit(4)
    else:
        cprint('[*] Roster Loaded Successfully!', 'green')

        cprint('[-] Sending Weekly Emails...', 'yellow')
        gmail = Gmail(gmail_username, gmail_app_password)
        for student in database.get_all_students():
            if student.email in blacklist:
                continue
            cprint(f'\r[-] Sending An Email To {student.name}'.ljust(50), 'yellow', end='')
            gmail.email(Receiver=student.email, Message=weekly_message, Subject=weekly_subject,
                        Cc='centraltutorsupport@bootcampspot.com')
            sleep(2)
        cprint(f'\r[*] Emails Have Successfully Sent!'.ljust(50), 'green')


if __name__ == '__main__':
    main()
