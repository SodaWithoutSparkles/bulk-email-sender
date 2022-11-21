#!/usr/bin/env python3
# standard lib
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr
from email.mime.text import MIMEText
import smtplib
# need to install
from pandas import ExcelFile

# This code is heavily commented. You can read the comments (lines with #) to see what this is doing
# module not found error? you may need to install pandas and openpyxl
# pip install pandas openpyxl

# edit the following configs... --------------------------------------------------
# your email to login, also appear in the From field of email
SENDER = ""

# your name, appear in email
YOUR_NAME = ""

# event name, appear in email
EVENT = ""

# NOT YOUR GOOGLE PASSWORD!!! It is a 16 char pw
# https://support.google.com/accounts/answer/185833?hl=en
# select "mail" for app and give it a name
# I use "bulk email sender"
# You can disable 2FA after using, or you may unable to login google without phone (at school, for example).
# You can do so at https://myaccount.google.com/security
# You *may* need to remove/revoke the app PW before you disable 2FA
# You dont need to remember it. remove the app PW after run and generate a new PW next time
# run the code with a empty app PW will trigger a guide
# Example: GOOGLE_APP_PW = "aaaabbccddddeeff"
GOOGLE_APP_PW = ""

# excel path for recipient and their address
# Case sensitive, and use absolute path, i.e. start from C: if on windows, start from / if on *nix
# This is a raw string, no need to escape the \
# yes, the r before quotes is intended, to indicate its a raw string
# on windows, just right-click and copy path, make sure theres only 1 set of quotes
# On how the excel should be written, see readExcelFile()
# Example: r"C:\Users\alice\Documents\sendEmail.xlsx"
# NOT: ""Documents\sendEmail.xlsx""
EXCEL_PATH = r""

# email subject
SUBJECT = ""

# use {VARIABLE} for placeholder, change fillTemplate() and the excel file accordingly
# New lines/enter, tabs, spaces are preserved, just type it like normal email
# Yes, there are three double-quotes before and after the string, it is intended. 
# Dont change them, they will not be visible in the email
# Please use only alphanumeric characters, a-z, A-Z, 0-9
# unless you know what you are doing and escape them
TEMPLATE = """Dear {NAME}, 

Here are photos from {EVENT}, please check attached links/files.

{LINKS}

Best regards, 
{YOUR_NAME}
Photography Club"""

# Disable config check
I_UNDERSTAND_THE_RISKS_AND_WANT_TO_DISABLE_CONFIG_CHECKS = False

# Disable coloured printing. Set to True if see werid printing
DISABLE_COLOURS_AND_ANSI_ESCAPE_CODES = False
# DISABLE_COLOURS_AND_ANSI_ESCAPE_CODES = True

# End configs -------------------------------------------------------------------


# Helper vars for pretty print
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


DEBUG_HEAD = "[" + bcolors.ENDC + "{:^6}".format("DBG1") + bcolors.ENDC + "]"
INFO_HEAD = "[" + bcolors.OKCYAN + "{:^6}".format("INFO") + bcolors.ENDC + "]"
OK_HEAD = "[" + bcolors.OKGREEN + "{:^6}".format("OK") + bcolors.ENDC + "]"
WARD_HEAD = "[" + bcolors.WARNING + "{:^6}".format("WARN") + bcolors.ENDC + "]"
ERROR_HEAD = "[" + bcolors.FAIL + "{:^6}".format("ERR!") + bcolors.ENDC + "]"
# End vars for pretty print


def readExcelFile():
    # read excel file
    # contain at least 3 case-sensitive columns, "name", "to", "links"
    # dont create more than 1 columns with the above names
    # the code will find columes named "name", "to", and "links" and fill the placeholders
    # that should roughly look like this
    #    name   |         to         |       links
    #    Alice  | alice@example.com  | https://example.com
    #    Bob    |  bob@example.com   | https://example.net
    #    Chris  | chris@example.com  | https://example.org

    print(INFO_HEAD, "reading excel files", end='\r')
    xls = ExcelFile(EXCEL_PATH)
    print(INFO_HEAD, "parsing excel files", end='\r')
    # read the first sheet. Python is 0-indexed, i.e. first item is 0, second item is 1
    df = xls.parse(xls.sheet_names[0])
    if (df.isnull().values.any()):
        print(ERROR_HEAD, "Excel has empty cell(s). Fix the excel first.")
        exit(1)
    elif len(df.index) == 0:
        print(ERROR_HEAD, "Excel has no rows.     ")
        exit(1)
    print(OK_HEAD,
          "finished parsing excel file with {} rows".format(len(df.index)))
    return df


def sendEmail(emailObjs):
    # Give a list of email objects, send for each obj.
    # Return number of errors
    # count num of errors
    errorCount = 0
    # setup email server
    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:
        try:
            print(INFO_HEAD, "Connecting to mail server", end='\r')
            smtp.ehlo()  # verify smtp server
            smtp.starttls()  # establish TLS
            smtp.login(SENDER, GOOGLE_APP_PW)  # Login as user w/ cred
            print(OK_HEAD, "Connected to mail server              ")

            for emailobj in emailObjs:
                try:
                    print(INFO_HEAD,
                          "sending email: {}".format(emailobj['to']),
                          end='\r')
                    smtp.send_message(emailobj)
                    print(OK_HEAD, "sending email: {}".format(emailobj['to']))
                except Exception as e:
                    errorCount += 1
                    print(
                        ERROR_HEAD, "Error when sending mail to: {}".format(
                            emailobj['to']))

        except Exception as e:
            errorCount += 1
            print(ERROR_HEAD, "Error when setup email server: \n", e)
            raise Exception(ERROR_HEAD + "Error when setup email server")

    if errorCount > 0:
        print(WARD_HEAD,
              "send email finished with {} error(s)".format(errorCount))
        return errorCount
    else:
        print(INFO_HEAD, "send email finished with 0 error(s)")
    return 0


def setEmail(subj: str, to: str, body: str):
    # set email format to be used later
    content = MIMEMultipart()
    content["subject"] = subj
    content["from"] = SENDER
    content["to"] = to
    content.attach(MIMEText(body))
    return content


def fillTemplate(name, event, links, your_name):
    # replace text in template to entries in excel
    # if you add a new placeholder called {foo} in template
    # add "foo=foo, " without quotes in front of YOUR_NAME
    # add "foo, " without quotes in front of the 'name' in 'def fillTemplate(name...)' above
    # then follow instructions before the "body = fillTemplate" line in main()
    out = TEMPLATE.format(YOUR_NAME=your_name,
                          LINKS=links,
                          EVENT=event,
                          NAME=name)
    return out


def printMail(content):
    # get mail main body
    if content.is_multipart():
        for part in content.get_payload():
            body = part.get_payload()
    else:
        body = content.get_payload()
    print(INFO_HEAD, "Check if the following is fine")
    if '{' in body or '}' in body:
        # maybe not filled placeholders?
        print(
            WARD_HEAD,
            "'{' or '}' in body, check CAREFULLY for un-filled placeholders")
    print("-------------------------------------------------------")
    print("From: {}".format(content['from']))
    print("To: {}".format(content['to']))
    print("Subject: {}".format(content['subject']))
    print(body)
    print("-------------------------------------------------------")


def main():
    excel = readExcelFile()
    to_send = excel.to_dict()
    # print(to_send)
    emailObjList = []
    firstMail = True
    for i in range(len(to_send['to'])):
        # set items to those in excel file
        # if you add a new placeholder {foo}
        # add foo = to_send['foo'][i]
        # add column 'foo' in excel
        name = to_send['name'][i]
        to = to_send['to'][i]
        links = to_send['links'][i]
        # attempt to allow multiple links
        links = links.replace(' ', '\n')
        # if add foo, change to: body = fillTemplate(foo, name, EVENT, links, YOUR_NAME)
        body = fillTemplate(name, EVENT, links, YOUR_NAME)
        content = setEmail(SUBJECT, to, body)
        if firstMail:
            # show the first email to check
            printMail(content)
            input(INFO_HEAD + " Press Enter to continue, Ctrl-C to break...\r")
            if not DISABLE_COLOURS_AND_ANSI_ESCAPE_CODES:
                print(
                    "\033[A                                                            \033[A"
                )
            # anything else is not first mail from now on
            firstMail = False
        # append to pending email list
        emailObjList.append(content)
    # send all email in pending list, sendEmail will return number of failed mails
    failedMailCount = sendEmail(emailObjList)
    if failedMailCount == 0:
        print(OK_HEAD, "Done!")
    print(INFO_HEAD, "You may want to disable 2FA, read line 28 of the file: ")
    # print line 28 and 29
    with open(__file__) as f:
        for i, line in enumerate(f):
            if 26 < i < 29:
                print(line, end='')
            elif i > 28:
                break


def checkCred():
    # check sender
    if '@' not in parseaddr(SENDER)[1]:
        raise Exception(ERROR_HEAD + "SENDER address not valid")
    if YOUR_NAME == "":
        print(WARD_HEAD, "YOUR_NAME field is empty. ", end='')
        input("Press Enter to continue, Ctrl-C to break...")
    if EVENT == "":
        raise Exception(ERROR_HEAD + "EVENT cannot be empty")
    if GOOGLE_APP_PW == "":
        print(
            '''GOOGLE_APP_PW is empty. Create one and fill it in the config section of this code.
        1. Visit https://myaccount.google.com/security
        2. Enable 2-Step Verification (also known as 2FA)
        3. Follow prompt. You will need a phone number
        4. After you enable 2FA, go back to https://myaccount.google.com/security
        5. In "App Passwords", create a app password
        6. Select "Mail" for app, "Other" for device
        7. Give it a name, like "bulk email sender"
        8. Click Generate, a new window appeared
        9. Copy the 16 character password into the script GOOGLE_APP_PW variable. Just select and copy.
        10. Confirm GOOGLE_APP_PW has no spaces and is contained in brackets just like the examples
        11. Click Done on the app password page
        12. Run the script again
        13. When you finished using this script, click the trash bin icon to delete the password
        14. Do not try to remember it. Just delete after use.
        15. You may want to disable 2FA after use. Or you may not be able to login without your phone (for example at school).
        
        The above is adopted from the official google help: https://support.google.com/accounts/answer/185833?hl=en'''
        )
        print(ERROR_HEAD + " GOOGLE_APP_PW not valid")
        exit(1)
    if len(GOOGLE_APP_PW) != 16:
        raise Exception(
            ERROR_HEAD +
            " GOOGLE_APP_PW not valid. It has to be length 16. Did you used your google account password?"
        )
    if EXCEL_PATH == "":
        raise Exception(ERROR_HEAD + " EXCEL_PATH cannot be empty")
    if not (EXCEL_PATH.startswith('C:') or EXCEL_PATH.startswith('c:')
            or EXCEL_PATH.startswith('/')):
        raise Exception(ERROR_HEAD +
                        " EXCEL_PATH has to be full path, not relative path")
    if SUBJECT == "":
        raise Exception(ERROR_HEAD + " SUBJECT cannot be empty")
    if TEMPLATE == "":
        raise Exception(ERROR_HEAD + " TEMPLATE cannot be empty")


if __name__ == "__main__":
    if not (I_UNDERSTAND_THE_RISKS_AND_WANT_TO_DISABLE_CONFIG_CHECKS):
        checkCred()
    else:
        print(WARD_HEAD, "CONFIG CHECKS DISABLED, proceed at your own risk")
    if DISABLE_COLOURS_AND_ANSI_ESCAPE_CODES:
        DEBUG_HEAD = "[" + "{:^6}".format("DBG1") + "]"
        INFO_HEAD = "[" + "{:^6}".format("INFO") + "]"
        OK_HEAD = "[" + "{:^6}".format("OK") + "]"
        WARD_HEAD = "[" + "{:^6}".format("WARN") + "]"
        ERROR_HEAD = "[" + "{:^6}".format("ERR!") + "]"
        print(WARD_HEAD, "colours disabled")
    main()