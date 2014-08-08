__author__ = 'Ashar Malik'

import csv, time
import smtplib, os
import imaplib
import email
from shutil import rmtree
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders
from os import listdir

sender_username = ""
sender_pass = ""

smtp_address = ""
smtp_port = 80

#send_mail([to], subject, text, [attachments], [cc])
def send_mail(send_to, subject, text, files=[], cc=[]):
    assert isinstance(send_to, list)
    assert isinstance(files, list)
    assert isinstance(cc, list)

    msg = MIMEMultipart()
    msg['From'] = sender_username
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    #new things...
    msg['CC'] = COMMASPACE.join(cc)
    send_to = send_to + cc

    msg.attach( MIMEText(text) )

    for f in files:
        print "Attaching %s..." % os.path.basename(f),
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    print "Sending...",
    smtp = smtplib.SMTP('smtpout.secureserver.net', 80)
    smtp.login(sender_username, sender_pass)
    smtp.sendmail(sender_username, send_to, msg.as_string())
    smtp.close()

def read_csv(file_name):
    csv_info = []
    with open(file_name, 'rb') as f:
        reader = csv.reader(f)
        for row in reader:
            csv_info.append(row)

    return csv_info

def log(text):
    global current_log
    print "Log: %s" % text

    text = "[" + time.strftime("%d/%m/%Y") +" "+ time.strftime("%I:%M:%S") + "] " + text
    current_log = current_log + text +"\n"
    with open("bin/log.txt", 'a') as f:
        f.write(text+"\n")

def filter_whitespace(name):
    while name.endswith(" "):
        name = name[0:-1]
        if name.__len__() == 0:
            return name
    return name

def correct_email(email):
    email = email.replace(" ", "") #get rid of whitespace
    while email[-1:].isalpha() == False: #get rid of any trailing punctuation (non-alphabetic characters)
        email = email[:-1]
    return email

def run_campaign(campaign_loc):
    global current_log
    begin = time.time()
    email_list = []
    email_problems = []
    cc = []
    current_log = ""

    if os.path.isfile(campaign_loc + 'to_email.csv') == False:
        print "No emailing list provided in this email! Please make sure you saved it as 'to_email' as 'CSV (Comma Delimited)' or 'Windows Comma Separated' on Mac."
        return -1

    try:
        csv_info = read_csv(campaign_loc + 'to_email.csv')
    except Exception as e:
        print "Unable to open CSV file. Did this file come from a Mac? Make sure you saved as 'Windows Comma Separated'"
        print "You can also open the file on this computer and resave it as CSV."
        return  -1
    #ask for subject of email

    if os.path.isfile(campaign_loc + 'template.txt'):
        with open(campaign_loc + 'template.txt', 'rb') as f:
            template = f.read()
    else:
        print "No template in this email! Make sure you save it as 'template.txt'."
        return -1

    #change %Name to %name
    template = template.replace("%Name", "%name")

    subject = raw_input("Subject: ")

    headers = csv_info.pop(0)

    #make headers lowercase
    for i, item in enumerate(headers):
        headers[i] = str(item.lower()).replace(" ", "") #remove whitespace as well

    email_index = headers.index("email")
    replace_list = []
    #compile index/phrase list to replace in template
    for i, to_replace in enumerate(headers):
        if to_replace != "email":
            replace_list.append((to_replace, i))

    #fill in templates
    for item in csv_info:
        email_str = template
        to_email = item[email_index]
        if len(to_email) != 0:
            for replace_token in replace_list:
                replacement = filter_whitespace(item[replace_token[1]]) #removes any trailing spaces

                email_str = email_str.replace('%'+replace_token[0], replacement)
            #template is done being filled in
            to_email = correct_email(to_email)
            email_list.append((email_str, to_email))

    #.DS_Store
    #grab attachments
    if os.path.isdir(campaign_loc+'Attachments') == False:
        print "No attachments found."
        attachments = []
    else:
        attachments = listdir(campaign_loc+'Attachments')
        if ".DS_Store" in attachments:
            attachments.remove(".DS_Store")


    if template.lower().__contains__("attach") and attachments.__len__() == 0:
        while True:
            print "The template mentions attachments. But I did not find any attachments. Continue? y/n"
            usr_input = raw_input('>> ')
            if usr_input == "y":
                break
            elif usr_input == "n":
                return -1

    cc_term = "a"
    while True:
        print "Would you like to add %s CC? (y/n)" % cc_term
        cc_decision = raw_input(">> ")

        if cc_decision == "y":
            cc_person = raw_input("Type an email: ")
            cc.append(cc_person)
            cc_term = "another"
        elif cc_decision == "n":
            break

    print "Filled in Template:\n"+email_list[0][0]
    print "----------------"
    print "From: %s" % sender_username
    print "Subject: %s" % subject
    print "CC: %s" % cc
    print "Attachments: %s" % attachments

    while True:
        usr_input = ""
        usr_input = raw_input("Continue (y/n): ")

        if usr_input == "n":
            return  -1
        elif usr_input == "y":
            break

    if attachments.__len__() > 0:
        for i, attachment in enumerate(attachments):
            attachments[i] = campaign_loc+'Attachments/'+attachment

    #sending emails now
    emails_sent = []

    while True:
        email_problems = []
        total_emails = len(email_list)
        for i, email in enumerate(email_list):
            to = email[1]
            text = email[0]
            print "(%s/%s) Emailing %s..." % (i+1, total_emails, to),
            try:
                send_mail([to], subject, text, attachments, cc)
            except Exception as e:
                print "An error occurred."
                email_problems.append(email)
            else:
                print "Done."
                emails_sent.append(to)

        if email_problems.__len__() > 0:
            print "These emails were not sent(%s):" % len(email_problems)
            for email in email_problems:
                print email[1]

            skip_problems = False
            while True:
                retry_reply = raw_input("Retry? (y/n): ")
                if retry_reply == "n":
                    skip_problems = True
                    break
                elif retry_reply == "y":
                    email_list = email_problems
                    break

            if skip_problems == True:
                break
        else:
            break

    #done sending emails
    elapsed_time = time.time()-begin #in seconds
    emails_sent = len(email_list) #not true

    print "These people were successfully emailed:\n%s" % '\n'.join(zip(*email_list)[1])+"\n"

def loadSettings():
    global smtp_address
    global smtp_port
    global sender_username
    global sender_pass

    if os.path.isfile("settings.txt") == False:
        to_write = "address=%s\n" % smtp_address
        to_write = to_write+"port=%s\n" % smtp_port
        to_write = to_write+"username=%s\n" % sender_username
        to_write = to_write+"pass=%s" % sender_pass

        with open('settings.txt', 'w') as f:
            f.write(to_write)

        print "Created settings.txt:\n%s" % to_write
        return

    with open('settings.txt', 'r') as f:
        variables = f.readlines()

        for i, line in enumerate(variables):
            variables[i] = line.replace("\n", "")
            print variables[i]
            variables[i] = variables[i].split("=")

            if variables[i][0] == 'address':
                smtp_address = variables[i][1]
            elif variables[i][0] == 'port':
                smtp_port = int(variables[i][1])
            elif variables[i][0] == 'username':
                sender_username = variables[i][1]
            elif variables[i][0] == 'pass':
                sender_pass = variables[i][1]
    print "Successfully loaded settings."

def list_dir():
    print "Available campaigns:"

    listing = [name for name in os.listdir("Email Campaigns") if os.path.isdir(os.path.join("Email Campaigns", name))]
    for i, directory in enumerate(listing):
        print "%s: %s" % (i, directory)
    print "Type a number or press ENTER to refresh:"
    return listing

def main():
    global sender_username
    global sender_pass
    #create skeleton
    if os.path.isdir('bin') == False:
        os.makedirs('bin')
    if os.path.isdir('Email Campaigns') == False:
        os.makedirs('Email Campaigns')
    #maybe skip if there is only one campaign? (ask if they want to use that one y/n
    options = list_dir()
    while True:
        usr_input = raw_input('>> ')
        print usr_input
        if usr_input.isdigit():
            index = int(usr_input)
            if index+1 <= len(options):
                if run_campaign('Email Campaigns/'+options[index]+'/') == -1:
                    print "Campaign could not run. Please type a number or press ENTER to refresh:"
                else:
                    print "Campaign finished."
                    options = list_dir()
            else:
                print "Number is not an option. Type a number or press ENTER to refresh:"
        elif usr_input == "set_user":
            sender_username = raw_input("Type username: ")
            print "Username set to %s." % sender_username
            sender_pass = raw_input("Type password: ")
            print "Password set to %s." % sender_pass
        else:
            options = list_dir()

pass

loadSettings()
main()

#todo: show astericks for password
#add multi account options
#error detection for names/emails
#export unsent emails?