#!/usr/bin/python
import os
from email.mime.application import MIMEApplication
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from appscript import *
import smtplib
import pdb

fro = '<your-email-address>'
nu = app( 'Numbers' )
table = nu.documents.first.sheets()[0].tables()[0]
rows = table.rows.get()
submissions = os.listdir( '.' )
mailer = smtplib.SMTP('smtp.gmail.com', 587)
mailer.ehlo()
mailer.starttls()
mailer.login( fro, "<email-password>" )
SUBJECT = "<Course Number> HW<homework-number> Gradesheet (%s)"
UNSUBMITTED_MAIL = "Hi,\n\nYou have not submitted this homework <X>. Please reply to this email if you think that you had submitted."
SUBMITTED_MAIL = "Hi,\n\nPlease find attached, your gradesheet for the homework."

for i in range( 1, len( rows ) ):
    cells = rows[ i ].cells()
    print "Sending Email to: %s at %s" % ( cells[ 2 ].value(), cells[ 1 ].value() )
    netid = cells[ 2 ].value()
    studentEmail = cells[ 1 ].value()
    msg = MIMEMultipart()
    msg[ 'Subject' ] = SUBJECT % netid
    msg[ 'From' ] = fro
    msg[ 'To' ] = studentEmail
    if netid not in submissions:
        print netid, studentEmail, "did not submit this homework"
        text = UNSUBMITTED_MAIL
    else:
        text = SUBMITTED_MAIL
        gradesheetName = "%s/FALL2015_HW<X>_Grading-%s.pdf" % ( netid, netid )
        print "Sending file:", gradesheetName
        pdfFile = open( gradesheetName, 'rb' ).read()
        msgPdf = MIMEApplication( pdfFile, 'pdf' )
        msgPdf.add_header( 'Content-Disposition', 'attachment', filename=gradesheetName )
        msgPdf.add_header( 'Content-Disposition', 'inline', filename=gradesheetName )
        msg.attach( msgPdf )

    body = MIMEMultipart( 'alternative' )
    part1 = MIMEText( text, 'plain' )
    body.attach( part1 )
    msg.attach( body )
    mailer.sendmail(fro, studentEmail, msg.as_string())

mailer.close()
