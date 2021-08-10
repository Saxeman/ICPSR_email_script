import smtplib
import ssl
import argparse
import csv
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase

smtp_server = 'smtp.mail.umich.edu'
port = 465

sender = 'ajsaxe@umich.edu'
    # input('Enter your email here:')
sender_icpsr = 'icpsr-sptechsupp@umich.edu'
password = 'CircleofSin123!'
    # input('Enter your password here:')
receiver = ''

# turn these into MIME objects so they can be easily formatted as emails
instructor_msg = """\
<p>This message is sent to you on behalf of ICPSR Summer Program Computing Support. Please retain it for future 
reference.</p>
<p>To ensure that messages from ICPSR Summer Program Computing Support are not flagged as spam by your email 
application, please add the following email address to your contact list: icpsr-sptechsupp@umich.edu.</p>
<p>We are writing to provide you with important information about the University of Michigan account credentials that you 
will use to participate in your short workshop(s).</p>
<p>A U-M uniqname, UMICH password, and enrollment in Duo two-factor authentication are required for participation in the 
ICPSR Summer Program.</p>
<b>If you are not a regular University faculty, staff or student</b>, you will receive an additional email message 
from icpsr-sptechsupp@umich.edu that will contain your U-M uniqname (the U-M term for a user name). Upon receipt 
of your U-M uniqname, complete the following steps:<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1. <b>Set your UMICH password</b>. To set your password, visit 
https://password.it.umich.edu/password/forgot and enter your date of birth. You will receive a password reset code. 
Enter the code at the prompt and set your password. Refer to http://documentation.its.umich.edu/node/240 for additional 
guidance about passwords. If you need assistance with a password reset, contact the U-M ITS Service Center at 
4help@umich.edu  or 734.764.HELP.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2. <b>Enroll in Duo</b>. Many U-M web resources employ Duo for 
two-factor authentication (2FA). Two-factor authentication requires that you provide two proofs of identity: 
something you know (password) and something you have (a device enrolled in Duo). Duo two-factor authentication is 
required to access to access Summer Program Canvas sites. For a description of Duo enrollment options and enrollment 
instructions for each option, visit https://documentation.its.umich.edu/2fa/options-two-factor-authentication. 
To enroll a device in Duo or to manage your Duo devices, visit 
U-M Account Management at https://password.it.umich.edu.
<p><b>We strongly encourage all faculty and instructional aides</b> to visit the Summer Program Instructional Staff 
Canvas site and review the module titled Instructor Preparation - Information Technology 
(https://umich.instructure.com/courses/485137/modules). There you will find important information about your 
U-M account, Canvas, Zoom and additional steps that you must take to prepare for the Program.</p>
<p><i>Failure to complete these tasks in advance will impede your ability to engage in Summer Program instructional 
activities.</i></p>
<p><i>The Responsible Use of Information Resources</i> [Standard Practice Guide (SPG) 601.07] applies to all members 
of the University community and refers to all information resources (http://spg.umich.edu/policy/601.07). 
By using the university's technology services, you agree to follow U-M information technology policies and guidelines 
for responsible use. Inappropriate use of U-M technology resources may result in termination of access, disciplinary 
review, expulsion from the University, termination of employment, legal action or other disciplinary action. 
For information about responsible and appropriate use, see https://it.umich.edu/information-technology-policies.</p>
<p>If you have questions about the contents of this message, or if you are in need of Summer Program computing support, 
send email to icpsr-sptechsupp@umich.edu.</p>
<p>Welcome to the ICPSR Summer Program!</p>
"""
participant_msg = """\
<p>This message is sent to you on behalf of ICPSR Summer Program Computing Support. Please retain it for future 
reference.</p>
<p>To ensure that messages from ICPSR Summer Program Computing Support are not flagged as spam by your email application, 
please add the following email address to your contact list: icpsr-sptechsupp@umich.edu.</p>
<p>We are writing to provide you with important information about the University of Michigan account credentials that 
you will use to participate in your short workshop(s).</p>
<p>A U-M uniqname, UMICH password, and enrollment in Duo two-factor authentication are required for participation in 
the ICPSR Summer Program.</p>
<b>If you are not a regular University faculty, staff or student</b> and if you have not already received a U-M uniqname, 
you will receive an additional email message from icpsr-sptechsupp@umich.edu that will contain your U-M uniqname 
(the U-M term for a user name). Upon receipt of your U-M uniqname, complete the following steps:<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1. <b>Set your UMICH password</b>. To set your password, visit 
https://password.it.umich.edu/password/forgot and enter your uniqname. When prompted to verify your identity, 
enter the email address that you provided when you registered for the ICPSR Summer Program (i.e., the email 
address at which you received this message). You will be prompted to enter the password reset code that was 
sent to that email address. Enter the code and set your password. Refer to http://documentation.its.umich.edu/node/240 
for additional guidance about passwords.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2. <b>Enroll in Duo</b>. Many U-M web resources employ Duo for 
two-factor authentication (2FA). Two-factor authentication requires that you provide two proofs of identity: 
something you know (password) and something you have (a device enrolled in Duo). Duo two-factor authentication 
is required to access to access Summer Program Canvas sites. For a description of Duo enrollment options and enrollment 
instructions for each option, visit https://documentation.its.umich.edu/2fa/options-two-factor-authentication. 
To enroll a device in Duo or to manage your Duo devices, visit 
U-M Account Management at https://password.it.umich.edu.
<p><b>All participants must</b> visit the main Summer Program Canvas site and review the module titled 
Participant Preparation - Information Technology (https://umich.instructure.com/courses/483457/modules). 
There you will find important information about your U-M account, instructional resources, software and 
additional steps that you must take to prepare for participation in the Program.</p>
<p><i>Failure to complete these tasks in advance will impede your participation in Summer Program instructional 
activities</i>.</p>
<p>Your access to Summer Program Canvas resources will expire at 12:00 AM Eastern Daylight Time (EDT) 
14 days after the last day of your workshop/course/lecture unless your instructor has requested different arrangements. 
Please consult your instructor to confirm the period of extended access. When the period of extended 
access expires, you will lose access to U-M Canvas and all related materials.</p>
<p>The Responsible Use of Information Resources [Standard Practice Guide (SPG) 601.07] applies to all members of the 
University community and refers to all information resources (http://spg.umich.edu/policy/601.07). By using the 
university's technology services, you agree to follow U-M information technology policies and guidelines for 
responsible use. Inappropriate use of U-M technology resources may result in termination of access, disciplinary 
review, expulsion from the University, termination of employment, legal action or other disciplinary action. 
For information about responsible and appropriate use, see https://it.umich.edu/information-technology-policies.</p>
<p>If you have questions about the contents of this message, or if you are in need of Summer Program computing support, 
send email to icpsr-sptechsupp@umich.edu.</p>
<p>Welcome to the ICPSR Summer Program!</p>
"""
username_msg = """\
Content-Type: text/html
From: {}
To: {}
Subject: ICPSR Summer Program Computing Support Message
<body><html>
Dear {} {},<br>
<br>
Welcome to the ICPSR Summer Program.<br>
<br>
UNIQNAME: Your University of Michigan uniqname is: {}<br>
<br>
EXPIRATION: Your access to Canvas resources will expire at 12:00 AM (EDT), {}<br>
<br>
"""
password_msg = """\
Content-Type: text/html
From: {}
To: {}
Subject: ICPSR Summer Program Computing Support Message    
</body></html>
Dear {} {},<br>
<br>
Welcome to the ICPSR Summer Program.<br>
<br>
Your University of Michigan UMICH password is : {}<br>
<br>     
If the password field is blank, your UMICH password has already been assigned to you.<br>
<br>
"""

signature = """\
Please send mail to icpsr-sptechsupp@umich.edu if you have questions about this communication.<br>
<br>
Sincerely,<br>
Edward J. Czilli<br>
Project Manager, ICPSR Summer Program Computing Services<br>
<br>
--<br>
Computing Support Services<br>
ICPSR Summer Program in Quantitative Methods of Social Research<br>
Inter-university Consortium for Political and Social Research<br>
University of Michigan</body></html>"""

context = ssl.create_default_context()

parser = argparse.ArgumentParser(description='ICPSR Email Sender')
parser.add_argument("-f", "--filename", nargs=1, help="filename", required=True)
parser.add_argument("-i", "--intro", action='store_true', help="Sends introduction email")
parser.add_argument("-u", "--send_username", action='store_true', help="Sends username to participant")
parser.add_argument("-p", "--send_password", action='store_true', help="Sends password to participant")

args = parser.parse_args()

exclude = 'umich'

# Loads the ICPSR Logo and sets up headers
fp = open('ICPSR_logo.png', 'rb')
image = MIMEImage(fp.read())
fp.close()
image.add_header('Content-ID', '<0>')

# Creates the instructor intro email
intro_msg_instructor = MIMEMultipart()
intro_msg_instructor['From'] = sender_icpsr
intro_msg_instructor['Subject'] = "ICPSR Summer Program: Information Technology, Instructor Preparation"
msg_content = MIMEText('<html><body><p><img src="cid:0"></p>' + instructor_msg + signature, 'html', 'utf-8')
intro_msg_instructor.attach(msg_content)

# Creates the participant intro email
intro_msg_participant = MIMEMultipart()
intro_msg_participant['From'] = sender_icpsr
intro_msg_participant['Subject'] = "ICPSR Summer Program: Information Technology, Participant Preparation"
msg_content = MIMEText('<html><body><p><img src="cid:0"></p>' + participant_msg + signature, 'html', 'utf-8')
intro_msg_participant.attach(msg_content)

# Attaches the image to the intro emails
intro_msg_instructor.attach(image)
intro_msg_participant.attach(image)


with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    with open(args.filename[0]) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        # creates the logfile
        logfile = open('logs.csv', 'w')
        server.login(sender, password)
        log = {'connected_to_server': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
        logfile.write(str(log) + '\n')
        for row in csv_reader:
            # checks if first row
            if row[0] == "GivenName":
                continue
            receiver = row[2]
            if args.intro:
                role = row[6].capitalize()
                if role == 'Instructor':
                    intro_msg_instructor['To'] = receiver
                    server.sendmail(sender, receiver, intro_msg_instructor.as_string())
                    log = {'intro_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                    logfile.write(str(log) + '\n')
                elif role == 'Participant':
                    intro_msg_participant['To'] = receiver
                    server.sendmail(sender, receiver, intro_msg_participant.as_string())
                    log = {'intro_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                    logfile.write(str(log) + '\n')
                else:
                    log = {'input_error': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                    logfile.write(str(log) + '\n')
                    exit(1)
            if args.send_username:
                if exclude in row[2]:
                    log = {'umich_error': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                    logfile.write(str(log) + '\n')
                    continue
                username_msg = username_msg.format(sender, receiver, row[0], row[1], row[3], row[5])
                server.sendmail(sender, receiver, username_msg + signature)
                log = {'username_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log) + '\n')
            if args.send_password:
                if exclude in row[2]:
                    log = {'umich_error': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                    logfile.write(str(log) + '\n')
                    continue
                password_msg = password_msg.format(sender, receiver, row[0], row[1], row[4])
                server.sendmail(sender, receiver, password_msg + signature)
                log = {'password_message': '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now())}
                logfile.write(str(log) + '\n')
