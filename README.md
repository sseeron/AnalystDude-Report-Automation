# AnalystDude
This is a collection of Python and VBA scripts that help me automate daily reports I send to company over email.

# Context
Part of my daily job is to update company dashboards (several Excel xlsx files) by running a bunch of SQL queries. And then email the updated xlsx files as attachments to the whole company. Without any BI solutions, this practice isn't sacalable so I automated the refresh&publish process to save time, and relieve me from repeating myself.

# How It Works
1. Embed relevant SQL queries with PowerQuery add-on in Excel
2. Use the VBA macro file (.xlsm) to automatically execute embedded queries and save the output as an xlsx file with today's date.
3. Running the Python script in this repo will first execute the macro file, and load the output xlsx file into gmail, and publish. The goal is "Push Button Publishing".

# How you can use it
1. Install <a href="https://www.python.org/downloads/">Python</a> environment on your computer.
2. Install PowerQuery as Excel Add-on. You can skip this and the next step if you have other ways of updating/generating Excel spreadsheet to be published.
3. Load and organize SQL and its outputs in your spreadsheet according to your business needs.
4. Tweak the attached macro file to your specific diectory/path, this should be consistent with the Python code
5. Tweak the Python code to link both the macro file and output xlsx file. Specify your email inputs, namely: subject, recipients, text, password, etc.

# Notes
1. I'm an analyst and by no means a good programmer. I researched around for this project and used various people's code I found through Googling. (Goolge is your best friend when you have a question.)
2. Please do whatever you want with the code and help improve it if possible. I'm new to GitHub and would love some guidance/suggestions/criticisms/.
3. I really think there are better ways of achieving the same 'push button publishing' goal than this one (with less steps and less components). So please do leave comments :) 
