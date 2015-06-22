# AnalystDude
This is a collection of Python and VBA scripts that help me automate daily reports I send to company over email.

# Context
Part of my daily job is to update company dashboards (several Excel xlsx files) by running a bunch of SQL queries. And then email the updated xlsx files as attachments to the whole company. Without any BI solutions, this practice isn't sacalable so I automated the refresh&publish process to save time, and relieve me from repeating myself.

# How It Works
1. Embed relevant SQL queries with PowerQuery add-on in Excel
2. Use the VBA macro file (.xlsm) to automatically execute embedded queries and save the output as an xlsx file with the dashboard's date. (In my case, I'm always updating dashboard with end-of-day yesterday's data, so Dashboard has yesterday's date on it.)
3. Running the Python script (PublishButton.py) in this repo will first execute the macro file, and load the output xlsx file into gmail, and publish. The goal is "Push Button Publishing".

# How you can use it
1. Install <a href="https://www.python.org/downloads/">Python</a> environment on your computer.
2. Install <a href="https://support.office.com/en-us/article/Introduction-to-Microsoft-Power-Query-for-Excel-6E92E2F4-2079-4E1F-BAD5-89F6269CD605">PowerQuery</a> as Excel Add-on. You can skip this and the next step if you have other ways of updating/generating Excel spreadsheet to be published.
3. Load and organize SQL and its outputs in your spreadsheet according to your business needs.
4. Tweak the attached macro_template file to your specific diectory/path, this should be consistent with the Python code
5. Tweak the Python scripts to link both the macro file and output xlsx file. Specify your email inputs, namely: subject, recipients, text, password, etc.
6. Run PublishButton.py, sit back and sip on your coffee.

# Notes
1. I'm an analyst and by no means a good programmer. I researched around for this project and used various people's code I found through Googling. (Goolge is your best friend when you have a question.)
2. Please do whatever you want with the code and help improve it if possible. I'm new to GitHub and would love some guidance/suggestions/criticisms/.
3. I really think there are better ways of achieving the same 'push button publishing' goal than this one (with less steps and less components). So please do leave comments :) 
