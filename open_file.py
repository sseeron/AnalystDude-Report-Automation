import os, sys, subprocess

def open_file(filename):
	# filename = '/Users/Dude/Daily_Report/Your1stDash_20150620.xlsx' for Mac
	# filename = 'C:/Users/Dude/Dail_Report/20150620/Your1stDash_20150620.xlsx' for PC
	if sys.platform == 'win32':
		os.startfile(filename)
	else:
		opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
		subprocess.call([opener, filename])
