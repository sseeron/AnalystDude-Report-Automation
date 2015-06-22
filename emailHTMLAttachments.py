import smtplib
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email import Encoders
import os
import mimetypes

def emailHTMLAttachments(subject, fromEmail, toEmailList, textMessage, HTMLMessage, attachment_paths):
    dctype='application/octet-stream'
    COMMASPACE = ', '

# Create the container (outer) email message.
    msg = MIMEMultipart()
    msg['Subject'] = subject
# me == the sender's email address
    msg['From'] = fromEmail
    msg['To'] = COMMASPACE.join(toEmailList)
    text = textMessage
    html = """\
    <html>
      <head></head>
      <body>
        %s
      </body>
    </html>
    """ % (HTMLMessage)
    
    body = MIMEMultipart('alternative')
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    body.attach(part1)
    body.attach(part2)
    msg.attach(body)

# add mutiple attachments to an Email
# attachment_paths is a list, like this:['/home/x/a.pdf', '/home/x/b.txt']
    for file_path in attachment_paths:
        ctype, encoding = mimetypes.guess_type(file_path)
        if ctype is None or encoding is not None:
            ctype = dctype
        maintype, subtype = ctype.split('/', 1)
        try:
            with open(file_path, 'rb') as f:
                part = MIMEBase(maintype, subtype)
                part.set_payload(f.read())
                Encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                print os.path.basename(file_path)
                msg.attach(part)
        except IOError:
             print "error: Can't open the file %s" % file_path

# Send the email via our own SMTP server.
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login("your@email.com","Your Password Here")
    s.sendmail(fromEmail, toEmailList, msg.as_string())
    s.quit()
