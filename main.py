import smtplib
import ssl
import argparse
import csv
import datetime

smtp_server = 'smtp.mail.umich.edu'
port = 465

sender = input('Enter your email here:')
password = input('Enter your password here:')
receiver = ''

signature = """\
Please send mail to icpsr-sptechsupp@umich.edu if you have questions about the this communication.

Sincerely,
Edward J. Czilli
Project Manager, ICPSR Summer Program Computing Services
--
Computing Support Services
ICPSR Summer Program in Quantitative Methods of Social Research
Inter-university Consortium for Political and Social Research
University of Michigan"""

context = ssl.create_default_context()

parser = argparse.ArgumentParser(description='ICPSR Email Sender')
parser.add_argument("-f", "--filename", nargs=1, help="filename", required=True)
parser.add_argument("-i", "--intro", action='store_true', help="Sends spam list email")
parser.add_argument("-c", "--compsupp", action='store_true', help="Sends support documentation")
parser.add_argument("-u", "--send_username", action='store_true', help="Sends username to participant")
parser.add_argument("-p", "--send_password", action='store_true', help="Sends password to participant")

args = parser.parse_args()

exclude = 'umich'

with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    with open(args.filename[0]) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        # creates the logfile
        logfile = open('logs.csv', 'w')
        server.login(sender, password)
        log = {'connected_to_server': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
        logfile.write(str(log) + '\n')
        for row in csv_reader:
            if row[0] == "GivenName":
                continue
            receiver = row[2]
            if args.intro:
                intro_message = """\
This is a test message because I dont have the boilerplate email yet, make sure you exclude this email address from your spam list

"""
                # possibly have this log also show who the recipient is
                # creates logs
                server.sendmail(sender, receiver, intro_message + signature)
                log = {'intro_message':'{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log) + '\n')

            if args.compsupp:
                compsupp_message = """\
This is a test message because I dont have the boilerplate email yet, here is a link to the documentation homepage www.google.com

"""
                server.sendmail(sender, receiver, compsupp_message + signature)
                log = {'compsupp_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log)+ '\n')
            if args.send_username:
                if exclude in row[2]:
                    print("umich email found")
                    continue
                username_message = """\
From: {}
To: {}
Subject: ICPSR Summer Program Computing Support Message 1

Dear {} {},

Welcome to the ICPSR Summer Program.

This is the first of two emails that we mentioned in a previous communication. Please retain this message for reference.

UNIQNAME: Your University of Michigan uniqname is: {}

EXPIRATION: Your access to Canvas resources will expire at 12:00 AM (EDT), {}

""".format(sender, receiver, row[0], row[1], row[3], row[5])
                server.sendmail(sender, receiver, username_message + signature)
                log = {'username_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log)+ '\n')
            if args.send_password:
                if exclude in row[2]:
                    print("umich email found")
                    continue
                password_message = """\
From: {}
To: {}
Subject: ICPSR Summer Program Computing Support Message 2
                
Dear {} {},
                
Welcome to the ICPSR Summer Program.
                
This is the second of two emails that we mentioned in a previous communication. Please retain this message for reference.
                
Your University of Michigan UMICH password is : {}
                
If the password field is blank, your UMICH password has already been assigned to you.

""".format(sender, receiver, row[0], row[1], row[4])
                server.sendmail(sender, receiver, password_message + signature)
                log = {'password_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log)+ '\n')
