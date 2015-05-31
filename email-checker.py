import argparse
#from BeautifulSoup import BeautifulSoup
import datetime
import email
from HTMLParser import HTMLParser
import imaplib
import logging
import os
import platform
import random
import string
import subprocess
import sys
import time
import uuid
import webbrowser
try:
    import _winreg
except ImportError:
    # We're not on Windows
    # Nothing else to do
    pass

DATE_FORMAT = '%m/%d/%Y %H:%M:%S'

# Configure the argument parser
parser = argparse.ArgumentParser(description="Automated IMAP email checker and opener.")
parser.add_argument("-s", "--server", dest="server", required=True,
    help="IMAP email server"
)
parser.add_argument("-u", "--username", dest="username", required=True,
    help="User's email address/login"
)
parser.add_argument("-p", "--password", dest="userpass", required=True,
    help="User's password"
)
parser.add_argument("-e", "--ssl", dest="ssl",
    action='store_true', default=False,
    help="Use IMAP over SSL"
)
parser.add_argument("-Smn", "--min-wait", dest="sleepMin",
    default=5, type=int,
    help="Minimum amount of time (in minutes) to wait before checking email again"
)
parser.add_argument("-Smx", "--max-wait", dest="sleepMax",
    default = 10, type = int,
    help="Maximum amount of time (in minutes) to wait before checking email again"
)
parser.add_argument("-k", "--keywords", dest="keywords",
    default = "",
    help="Comma-separated list of keywords. Enables social engineering checks for 1 or more keywords before opening emails or clicking links. Example: 'resume,job,opening,applicant'"
)
parser.add_argument("-f", "--log-file", dest="logfile",
    default="email-checker.log",
    help = "Filename to save logging output to. Must specify full file path"
)
parser.add_argument("-d", "--debug", dest="loglevel",
    action = "store_const", const = logging.DEBUG, default = logging.INFO,
    help="Log debug messages"
)
parser.add_argument("-i", "--install", dest="install",
    action='store_true', default=False,
    help="Install this script to start when the system boots"
)
parser.add_argument("-c", "--decode", dest="decode",
    action="store_true", default=False,
    help=" Decode the parameters (created during install)"
)
# Parse the arguments
args = parser.parse_args()

# Simple substitution cipher to obfuscate server name, login,
# password, and keywords stored in the registry
#
# Don't translate spaces, or it will break things when calling parameters
decode = string.maketrans(
    'TFGQberD2Oz9yAi6UW7ox01mSvIgCKMnPupJRcNjVd3Hq85La4wEsXhYkltfZB',
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890'
)
encode = string.maketrans(
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890',
    'TFGQberD2Oz9yAi6UW7ox01mSvIgCKMnPupJRcNjVd3Hq85La4wEsXhYkltfZB'
)

# Set variable values from argument parser
if args.decode:
    server = string.translate(args.server, decode)
    username = string.translate(args.username, decode)
    userpass = string.translate(args.userpass, decode)
    keywords = string.translate(args.keywords, decode).split(',')
else:
    server = args.server
    username = args.username
    userpass = args.userpass
    keywords = args.keywords.split(',')

ssl = args.ssl
sleepMin = args.sleepMin
sleepMax = args.sleepMax
if sleepMin > sleepMax:
    tmp = sleepMin
    sleepMin = sleepMax
    sleepMax = tmp
install = args.install
logLevel = args.loglevel

# Configure logging to file
logging.basicConfig(
    level = logLevel,
    filename = args.logfile,
    filemode = 'a',
    datefmt = DATE_FORMAT,
    format = '%(asctime)s: %(levelname)-8s %(message)s'
)

# Configure console logging
console = logging.StreamHandler()
console.setLevel(logLevel)
consoleFormat = logging.Formatter(
    fmt = '%(asctime)s: %(message)s',
    datefmt = DATE_FORMAT
)
console.setFormatter(consoleFormat)
logging.getLogger('').addHandler(console)

# Seed the random number generator
# No argument seeds with the current system time
random.seed()

#####################
# CLASS DEFINITIONS #
#####################

class UrlParser(HTMLParser):

    urls = []

    def get_urls(self):
        return self.urls

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            # Links
            for attr in attrs:
                if attr[0] == 'href':
                    self.urls.append(attr[1])
        elif tag == 'img':
            # Webbug
            for attr in attrs:
                if attr[0] == 'src':
                    self.urls.append(attr[1])

    def reset(self):
        HTMLParser.reset(self)
        self.urls = []

########################s
# FUNCTION DEFINITIONS #
########################

def check_keywords(msg):
    if any(keywords):
        # Keywords defined -- perform the checks
        logging.debug("Performing keyword check...")
        if msg.is_multipart():
            # Message is multipart
            # Find the plain text part to search
            for part in msg.walk():
                contentType = part.get_content_type()
                if contentType == 'text/plain':
                    text = part.get_payload()
                    # prefer plain text over html -- break
                    break
                elif contentType == 'text/html':
                    text = part.get_payload(decode=True)
        else:
            # Message is not multipart
            # The payload is the plain text email
            text = msg.get_payload()

        # Search the message sender, subject, and body
        # Everything in lowercase
        sender = msg['from'].lower()
        text = text.lower()
        subject = msg['subject'].lower()
        # Cycle through the array of keywords
        for word in keywords:
            word = word.lower()
            # Compare lowercase keyword to lowercase text
            if word in sender or word in subject or word in text:
                # We can return True as soon as we find a match
                # We don't need to iterate through the rest of teh keywords
                logging.info("Keyword match found: " + word)
                return True

        logging.info("Keyword match NOT found.")
        return False
    else:
        # No keywords defined -- skip the check
        # Return True so we "pass" the test
        logging.info("No keywords defined -- skipping keyword check.")
        return True

def check_mail():
    logging.debug("--------------------------------------------------")
    logging.debug("Checking email...")
    try:
        # Two-part setup:
        # 1) Connect to the server
        if ssl:
            imap = imaplib.IMAP4_SSL(server)
        else:
            imap = imaplib.IMAP4(server)
        logging.debug("Connected to IMAP server " + server)
        # 2) Log in to the server
        imap.login(username, userpass)
        logging.debug("Logged into IMAP server as " + username)

        # Select the default mailbox
        retcode, count = imap.select()
        #retcode, count = imap.select(readonly=1) # open readonly for testing (don't mark messages as SEEN)
        # retcode != 'OK' --> failed to select the malbox
        # count <= 0      --> no messages in mailbox
        if retcode != 'OK' or count <= 0:
            # Either way, we can't continue
            if retcode == 'OK':
                logging.debug("Failed to select the mailbox")
                # The mailbox exists and select succeeded with no messages
                # Close the mailbox
                imap.close()
            elif count <= 0:
                logging.debug("No messages in mailbox")
            # Log out and return so we can sleep and try again later.
            imap.logout()
            return
        logging.debug("Selected the default mailbox (Inbox)")

        # Search for UNSEEN messages
        retcode, msgIds = imap.search(None, '(UNSEEN)')
        if retcode != 'OK':
            logging.debug("Search for UNSEEN emails failed")
            # Something went wrong.
            # Close the mailbox, log out, and return so we can sleep and try again later.
            imap.close()
            imap.logout()
            return

        # Convert the msg IDs string to an array
        msgIds = msgIds[0].split(' ')
        logging.info("--------------------------------------------------")
        if not any(msgIds):
            logging.info("Found 0 UNREAD emails.")
            # No messages to process.
            # Close the mailbox, log out, and return so we can sleep and try again later.
            imap.close()
            imap.logout()
            return
        else:
            logging.info("Found " + str(len(msgIds)) + " UNREAD emails.")


        # Try to fetch each returned msg ID
        logging.debug("Retrieving message(s): " + ', '.join(msgIds))
        for msgId in msgIds:
            logging.info("--------------------------------------------------")
            retcode, message = imap.fetch(msgId, '(RFC822)')
            if retcode != 'OK':
                # Couldn't fetch the message. Continue with the next one.
                logging.debug("Could not FETCH message " + str(msgId))
            else:
                # Message successfully received
                # USe the email module -- it's easier than parsing imaplib messages
                msg = email.message_from_string(message[0][1])
                logging.debug("Message " + str(msgId))
                logging.info("Subject: '" + msg['subject'] + "'")
                logging.info("   From: " + msg['from'])

                # Perform keyword checks to determine if we'll process the message
                if check_keywords(msg):
                    # Passed keyword check (or no keywords defined)
                    # Process the message
                    process_message(msgId, msg)
                else:
                    # Failed keyword check -- skip this message
                    logging.debug("Skipping message " + str(msgId))

        # We're done with this round.
        # Close and log out
        imap.close()
        imap.logout()

    except Exception, err:
        logging.error(str(err))
        if imap:
            imap.logout()

    return

def clean_build_linux(specName):
    subprocess.call(["rm", "-rf", "./build"])
    subprocess.call(["rm", "-rf", "./dist"])
    subprocess.call(["rm", "-f", specName])
    logging.debug("Removed build-related files")

def clean_build_windows(specName):
    subprocess.call("cmd /C rmdir /S /Q  build\\")
    subprocess.call("cmd /C rmdir /S /Q  dist\\")
    subprocess.call("cmd /C del /F /Q " + specName)
    logging.debug("Removed build-related files")

def do_install():
    print
    logging.info("==================================================")
    logging.info("Email-checker Installation")
    logging.info("==================================================")
    print
    print "This will generate a standalone executable from this script and install"
    print "it using a user-specified name and location. It will also configure the"
    print "operating system to start the executable at boot time using the options"
    print "supplied by the user. This may require administrator/root privileges. If"
    print "you are not running this script with these privileges, it is STRONGLY"
    print "recommended that you exit and run this script with appropriate privileges"
    print "or ensure you the executable is installed to an unprivileged location."
    print
    response = get_yes_no_with_default("Continue?", True)
    if not response:
        sys.exit()

    print
    print "You can specify a name (e.g., Outlook, Firefox) for the executable file"
    print "so it blends in when viewed in a process list. On Windows systems, this"
    print "name is also used as the registry key name."
    print
    installName = raw_input("What should the file be named? [email-checker]: ")
    installName = installName or 'email-checker'

    cwd = get_path()
    print
    print "You can specify a location for the executable file. This make the program"
    print "appear to be running from a more legitimate location."
    print
    while True:
        installPath = raw_input("Where should the executable be located? [" + cwd + "]: ")
        # Get the full/absolute path
        installPath = os.path.abspath(installPath or cwd)
        # verify we have a valid path
        if os.path.isdir(installPath):
            break

    # Obfuscate some of the stored settings.
    # Use the "args" version since we transform some (like splitting keywords)
    # when the script starts.
    cryptServer = string.translate(args.server, encode)
    cryptUsername = string.translate(args.username, encode)
    cryptUserpass = string.translate(args.userpass, encode)
    cryptKeywords  = string.translate(args.keywords, encode)

    print
    print "--------------------------------------------------"
    print
    print "Name:      " + installName
    print "Path:      " + installPath
    print "Server:    " + cryptServer + " (" + args.server + ")"
    print "Use SSL:   " + str(args.ssl)
    print "Username:  " + cryptUsername + " (" + args.username + ")"
    print "Password:  " + cryptUserpass + " (" + args.userpass + ")"
    print "Sleep Min: " + str(args.sleepMin)
    print "Sleep Max: " + str(args.sleepMax)
    if any(keywords):
        print "Keywords:  " + cryptKeywords + " (" + args.keywords + ")"
    else:
        print "Keywords:  [None]"
    print "Log File:  " + args.logfile
    print "Debug:     " + "True" if args.loglevel == logging.DEBUG else "False"
    print
    response = get_yes_no_with_default("Continue?", True)
    if not response:
        sys.exit()

    print
    print "--------------------------------------------------"

    argString = "-s " + cryptServer + " " + \
                ("-e " if args.ssl else "") + \
                "-u " + cryptUsername + " " + \
                "-p " + cryptUserpass + " " + \
                "-Smn " + str(args.sleepMin) + " " + \
                "-Smx " + str(args.sleepMax) + " " + \
                ("-k " + cryptKeywords + " " if any(cryptKeywords) else "") + \
                "-f " + os.path.abspath(args.logfile) + " " + \
                ("-d " if args.loglevel == logging.DEBUG else "") + \
                "-c"

    # Do OS-specific build and installation
    operatingSystem = platform.system()
    if operatingSystem == 'Linux':
        do_install_linux(installPath, installName, argString)
    elif operatingSystem == 'Windows':
        do_install_windows(installPath, installName, argString)
    else:
        logging.debug("Unsupported operating system: " + os + ". Installation aborted.")

def do_install_linux(path, name, args):
    logging.debug("Generating Linux executable...")

    DEVNULL = open(os.devnull, 'w')
    exeAbsPath = path + "/" + name

    # Generate the .spec file and executable using pyinstaller
    subprocess.call(
        ["pyinstaller", "--onefile", "--name=" + name, "email-checker.py"],
        stdout=DEVNULL, stderr=subprocess.STDOUT
    )
    subprocess.call(
        ["pyinstaller", name + ".spec"],
        stdout=DEVNULL, stderr=subprocess.STDOUT
    )
    logging.info("Generated " + name)

    # Move/rename the generated file to it's final location
    try:
        os.rename("dist/" + name, exeAbsPath)
        logging.info("Moved " + name + " to " + path)
    except OSError, err:
        logging.error("Error moving " + name + " to " + path)
        logging.error(str(err))
        clean_build_linux(name + ".spec")
        return

    # Create a .desktop file
    f = open(name + ".desktop", 'w')
    f.write("[Desktop Entry]\n")
    f.write("Type=Application\n")
    f.write("Exec=" + exeAbsPath + " " + args + "\n")
    f.write("Hidden=false\n")
    f.write("X-GNOME-Autostart-enabled=true\n")
    f.write("Name=" + name + "\n")
    f.write("Comment=")
    f.close()
    logging.info("Desktop launcher file created")

    homedir = os.path.expanduser("~")
    if not os.path.isdir(homedir + "/.config/autostart"):
        os.makedirs(homedir + "/.config/autostart")
        os.chmod(homedir + "/.config/autostart", 0700)
        logging.debug("Created directory " + homedir + "/.config/autostart")

    try:
        os.rename(name + ".desktop", homedir + "/.config/autostart/" + name + ".desktop")
        logging.info("Moved " + name + ".desktop to " + homedir + "/.config/autostart/" + name + ".desktop")
    except OSError, err:
        logging.error("Error moving " + name + ".desktop to " + homedir + "/.config/autostart/" + name + ".desktop")
        logging.error(str(err))
        os.remove(exeAbsPath)
        logging.info("Removed " + exeAbsPath)
        clean_build_linux(name + ".spec")
        return

    # Clean up after the build
    clean_build_linux(name + ".spec")

def do_install_windows(path, name, args):
    logging.debug("Generating Windows executable...")

    DEVNULL = open(os.devnull, 'w')
    exeName = name + ".exe"
    exeAbsPath = path + "\\" + exeName

    # Generate the .spec file and executable using pyinstaller
    subprocess.call(
        "pyinstaller --onefile --name=" + name + " email-checker.py",
        stdout=DEVNULL, stderr=subprocess.STDOUT
    )
    subprocess.call(
        "pyinstaller " + name + ".spec",
        stdout=DEVNULL, stderr=subprocess.STDOUT
    )
    logging.info("Generated " + exeName)

    # Move/rename the generated file to it's final location
    try:
        os.rename("dist/" + exeName, exeAbsPath)
        logging.info("Moved " + exeName + " to " + path)
    except WindowsError, err:
        logging.error("Error moving " + exeName + " to " + path)
        logging.error(str(err))
        clean_build_windows(name + ".spec")
        return

    # Create the HKCU run key
    runKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    data = "\"" + exeAbsPath + "\" " + args
    try:
        rootKey = _winreg.OpenKey(
            _winreg.HKEY_CURRENT_USER,
            runKey,
            0,
            _winreg.KEY_SET_VALUE
        )
        _winreg.SetValueEx(
            rootKey,
            name,
            0,
            _winreg.REG_SZ,
            data
        )
        _winreg.CloseKey(rootKey)

        logging.info("Registry HKCU run key created")
        logging.debug("Key:  HKCU\\" + runKey)
        logging.debug("Name: " + name)
        logging.debug("Type: REG_SZ")
        logging.debug("Data: " + data)
    except WindowsError, err:
        logging.error("Error creating registry HKCU run key")
        logging.error(str(err))

    # Clean up after the build
    clean_build_windows(name + ".spec")

def get_path():
    if getattr(sys, 'frozen', False):
       # we are running in a |PyInstaller| bundle
        return os.path.dirname(sys.executable)
    else:
        # we are running in a normal Python environment
        return os.path.dirname(os.path.abspath(__file__))

def get_yes_no_with_default(text, default=False):
    valid = ['y', 'yes', 'n', 'no']
    while True:
        response = raw_input(text + " " + ("([Yes] / No)" if default else "(Yes / [No])") + ": ")
        # lower() to prevent case issues, set to default if no input
        response = response.lower() or ('yes' if default else 'no')
        if response in valid:
            break

    # Typed 'y' or 'yes'
    if response == 'y' or response == 'yes':
        return True
    else:
        return False

def main_loop():
    # Display some information about how we're configured
    print
    logging.info("==================================================")
    logging.info("Email checker started. Press [CTRL+C] to stop")
    logging.info("==================================================")
    logging.info("Checking email for: " + username)
    logging.debug("Using user password: " + userpass)
    logging.debug("Checking on server: " + server)
    if ssl:
        logging.debug("Using IMAP over SSL")
    logging.debug("Wait " + str(sleepMin) + " to " + str(sleepMax) + " minutes between checks")
    if any(keywords):
        logging.debug("Checking for " + str(len(keywords)) + " keywords: " + str(keywords))
    else:
        logging.debug("Not checking for keywords.")

    # Loop indefinitely
    while True:
        check_mail()
        # Randomize how long we sleep for
        randSleep = random.randint(sleepMin, sleepMax)
        logging.debug("--------------------------------------------------")
        logging.debug("Sleeping for " + str(randSleep) + " minute(s).")
        # Sleep
        time.sleep(randSleep * 60)

    return

def open_file(filename):
    logging.debug("Opening file \"" + filename + "\"")
    try:
        if sys.platform == "win32":
            # os.startfile is Windows-specific
            os.startfile(filename)
        else:
            # We're on a *nix box
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, filename])
    except:
        logging.debug("Error opening file \"" + filename + "\"")

    return

def process_attachment(part, ext='.bin'):
    # Attempt to save the file and get it's returned filename
    filename = save_attachment(part, ext)
    if filename:
        # Save was successful. Attempt to open the file.
        open_file(filename)

    return

def process_html(htmlPart):
    #soup = BeautifulSoup(htmlPart.get_payload(decode=True))
    #for link in soup.findAll('a'):
    #   if link['href']:
    #       url = link['href']

    # BeautifulSoup is a better way to extract links, but isn't
    # guaranteed to be installed. Use custom HTMLParser class instead.
    parser = UrlParser()
    parser.feed(htmlPart.get_payload(decode=True))
    parser.close()
    urls = parser.get_urls()

    if any(urls):
        for url in urls:
            logging.info("!!! Found a URL: " + url)
            # Attempt to open the url in the default browser
            # Use a new window to de-conflict other potential exploits
            # new=0 -> same window
            # new=1 -> new window
            # new=2 -> new tab
            logging.debug("Opening URL " + url)
            webbrowser.open(url, 1)

    return

def process_message(msgId, msg):
    logging.debug("Processing message " + str(msgId) + "...")

    # Uncomment below to save each email to a txt file for debugging
    #f = open(str(msgId) + ".txt", 'w')
    #f.write(msg.as_string())
    #f.close()

    if msg.is_multipart():
        # The message is multipart.
        # Walk through all parts looking for content types we know
        logging.debug("Message is multi-part.")
        for part in msg.walk():
            # Get the MIME type of the current message part
            contentType = part.get_content_type()
            logging.debug("Part: " + contentType)
            if contentType == 'multipart/mixed':
                pass   # do nothing
            elif contentType == 'multipart/alternative':
                pass   # do nothing
            elif contentType == 'text/plain':
                # Can't do anything with a plain text email...
                logging.info("Found a plain text email part. Not doing anything.")
            elif contentType == 'text/html':
                # Can look for links in HTML emails or html attachments
                logging.info("!!! Found an HTML email part/attachment")
                process_html(part)
            # The rest are all file attachments.
            # If we find one we recognize, process the attachment.
            elif contentType == 'application/pdf':
                logging.info("!!! Found a PDF")
                process_attachment(part, '.pdf')
            elif contentType == 'application/msword':
                logging.info("!!! Found a Word document")
                process_attachment(part, '.doc')
            elif contentType == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                logging.info("!!! Found a Word document")
                process_attachment(part, '.docx')
            elif contentType == 'application/vnd.ms-excel':
                logging.info("!!! Found an Excel worksheet")
                process_attachment(part, '.xls')
            elif contentType == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                logging.info("!!! Found an Excel worksheet")
                process_attachment(part, '.xlsx')
            elif contentType == 'application/vnd.ms-powerpoint':
                logging.info("!!! Found a PowerPoint presentation")
                process_attachment(part, '.ppt')
            elif contentType == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
                logging.info("!!! Found a PowerPoint presentation")
                process_attachment(part, '.pptx')
            elif contentType == 'application/vnd.oasis.opendocument.text':
                logging.info("!!! Found an OpenDocument Text Document")
                process_attachment(part, '.odt')
            elif contentType == 'application/vnd.oasis.opendocument.spreadsheet':
                logging.info("!!! Found an OpenDocument Spreadsheet")
                process_attachment(part, '.ods')
            elif contentType == 'application/vnd.oasis.opendocument.presentation':
                logging.info("!!! Found an OpenDocument Presentation")
                process_attachment(part, '.odp')
            elif contentType == 'application/vnd.oasis.opendocument.graphics':
                logging.info("!!! Found an OpenDocument Graphics Document")
                process_attachment(part, '.odg')
            elif contentType == 'application/vnd.oasis.opendocument.formula':
                logging.info("!!! Found an OpenDocument Formula")
                process_attachment(part, '.odf')
            else:
                # Unsupported/unknown ContentType
                pass   # do nothing
    else:
        # The message is not multipart -- it's palin text.
        logging.debug("Message is NOT multi-part.")
        # Not a whole lot we can do...
        logging.info("Got a plain text email. Not doing anything.")

    return

def save_attachment(part, ext='.bin'):
    logging.debug("Saving attachment...")
    # Try to get the filename from the part header
    filename = part.get_filename()
    if not filename:
        # Couldn't get the filename. Use a temporary one.
        filename = 'unknown' + ext

    counter = 1
    # Check if the filename already exists
    while os.path.isfile(filename):
        # Generate a one-up file name until we find one that doesn't exist.
        filename = 'part-%04d%s' % (counter, ext)
        counter += 1
        # After we reach 10000 files with the same extension, the filenames repeat.
        # Give ourselves an out so we don't loop forever.
        # This will cause file part-0000.ext to be overwritten
        if counter >= 10000:
            break

    try:
        # Decode the email part and write it to the file
        f = open(filename, 'wb')
        f.write(part.get_payload(decode=True))
        f.close()
    except:
        logging.error("An error occurred while saving the file")
        # Don't return a filename
        return None

    logging.debug("Attachment saved as \"" + filename + "\"")
    # Return the filename of the saved file
    return filename

################
# MAIN PROGRAM #
################

# Substitution codebook generator
#keyspace = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890'
#str1 = keyspace
#str2 = ""
#for i in range(0, len(str1)):
#    pos = random.randint(0, len(str1) - 1)
#    str2 += str1[pos]
#    str1 = str1[:pos] + str1[(pos+1):]
#print keyspace
#print str2

try:
    if install:
        do_install()
    else:
        main_loop()
except KeyboardInterrupt:
    # die gracefully on [ctrl+c]
    sys.exit()

