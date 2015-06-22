from open_file import *
from emailHTMLAttachments import *
from datetime import date, timedelta
import time

def main():
# Define some constants and inputs here: 
	# Output format is 2015-06-20
	yesterday = date.today() - timedelta(1) 

	# Change format to 20150620 for consistency since my Excel macro files save output in this format
	yrmoday = yesterday.strftime('%Y%m%d') 

	# Reassign value to yesterday to take on this format for email subject line: Sat, June 20, 2015
	yesterday = yesterday.strftime('%a, %B %d, %Y') 

	fromEmail = "Your name here. Not email address. E.g. 'The Dude's Reporting Bot' "
	# List of Lists : Highly recommend use team distribution email like team@yourcompany.com
	recipients = [['1st@gmail.com'],['2nd@gmail.com'],['etc@someemail.com']]  
	textMessage = 'Dude has nothing nice to say so dude keeps sipping on a white russian.'
	HTMLMessage = 'Put your email body text here.'

# Locate macro files and their output xlsx files:
	# Path of Your1stDash macro file
	Your1stDashMacro = 'C:/Users/Dude/Dail_Report/Your1stDashMacro.xlsm' 
	# Path of the output file that Your1stDash macro generates
	# like C:/Users/Dude/Dail_Report/20150620/Your1stDash_20150620.xlsx
	Your1stDashPath = 'C:/Users/Dude/Dail_Report/%s/Your1stDash_%s.xlsx' %(yrmoday,yrmoday) 

	Your2ndDashMacro = 'C:/Users/Dude/Dail_Report/Your2ndDashMacro.xlsm'
	Your2ndDashPath = 'C:/Users/Dude/Dail_Report/%s/Your2ndDash_%s.xlsx' %(yrmoday,yrmoday)

	macros = [
		Your1stDashMacro,\
		Your2ndDashMacro\
	]

# Open Excel Macro file so it can update itself, require import of open_file.py
# Ideally each of your macro file should execute some sort of query, update itself, 
# and save an output file for that day to be published:
	for f in macros:
		filename = f
		print 'Opening %s' %(f)
		open_file(filename)
		print 'Updating...'
	print 'Updating Dashboards done...\n'

# Load up email and send: 
	for toEmailList in recipients:
		print 'Loading files to email...'
		subject = '%s: This is your email subject line' % yesterday

		# Your attachments' paths go into this []
		attachment_paths = [Your1stDashPath, Your2ndDashPath] 

		# Call emailHTMLAttachments.py and pass arguments to it:
		emailHTMLAttachments(subject, fromEmail, toEmailList, textMessage, HTMLMessage, attachment_paths)
		print 'Dashboards are sent to %s\n' %(toEmailList);

	print 'Done. Check your email.'

main()
