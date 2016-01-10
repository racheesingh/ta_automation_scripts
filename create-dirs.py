#!/usr/bin/python

'''
This script will read a spreadsheet open in Numbers with student names/IDs/emails
and create a directory for every student. It will then copy a blank grade-sheet file
to each student's directory and rename it after the student.

'''
from appscript import *
import os
import shutil
nu = app( 'Numbers' )
table = nu.documents.first.sheets()[0].tables()[0]
rows = table.rows.get()
BLANK_GRADE_SHEET = "FALL2015_HW4_Grading.numbers"
USER_GRADE_SHEET = BLANK_GRADE_SHEET.split( '.' )[ 0 ] + '-%s.numbers'
GRADE_DIR = 'gradesheets/'

for i in range( 1, len( rows ) ):
    cells = rows[ i ].cells()
    netid = cells[ 2 ].value()
    try:
        if not os.path.exists( netid ):
            os.makedirs( GRADE_DIR + '/' + netid )
        fname = USER_GRADE_SHEET % netid
        shutil.copy2( BLANK_GRADE_SHEET, "%s/%s/%s" % ( GRADE_DIR, netid, fname ) )
    except: continue
    
