import argparse
import email
from HTMLParser import HTMLParser
import imaplib
import logging
import os
import platform
import random
import shutil
import signal
import string
import subprocess
import sys
import time
import threading
import webbrowser
try:
    import _winreg
except ImportError:
    pass #  # We're not on Windows -- do nothing

################################################################################

#############
# Constants #
#############

CONSOLE_LOG_FORMAT = "%(asctime)s: %(message)s"
DATE_FORMAT = '%m/%d/%Y %H:%M:%S'
FILE_LOG_FORMAT = "%(asctime)s: %(levelname)-8s %(message)s"
PLATFORMS = ['Linux', 'Windows']
SAVE_DIR = ".email-checker"

# Simple substitution cipher to obfuscate server name, login,
# password, and keywords stored for auto start
#
# Don't translate spaces, or it will break things when calling parameters
KEYSPACE = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890'
CODEBOOK = 'TFGQberD2Oz9yAi6UW7ox01mSvIgCKMnPupJRcNjVd3Hq85La4wEsXhYkltfZB'
DECODE = string.maketrans(CODEBOOK, KEYSPACE)
ENCODE = string.maketrans(KEYSPACE, CODEBOOK)

# MIME data
DOC = ["!!! Found a Word document", '.doc']
DOCX = ["!!! Found a Word document", '.docx']
HTML = ["!!! Found an HTML email part/attachment", '.html']
ODF = ["!!! Found an OpenDocument formula", '.odf']
ODG = ["!!! Found an OpenDocument graphics document", '.odg']
ODP = ["!!! Found an OpenDocument presentation", '.odp']
ODS = ["!!! Found an OpenDocument spreadsheet", '.ods']
ODT = ["!!! Found an OpenDocument text document", '.odt']
PDF = ["!!! Found a PDF", '.pdf']
PPT = ["!!! Found a PowerPoint presentation", '.ppt']
PPTX = ["!!! Found a PowerPoint presentation", '.pptx']
RTF = ["!!! Found a Rich Text Format file", '.rtf']
TXT = ["Found a plain text email part. Not doing anything.", '.txt']
XLS = ["!!! Found an Excel worksheet", '.xls']
XLSX = ["!!! Found an Excel worksheet", '.xlsx']

MIME_DICT = {
	'text/plain': TXT,
    'text/html': HTML,
    'application/pdf': PDF,
    'application/rtf': RTF,
    'application/msword': DOC,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        DOCX,
    'application/vnd.ms-excel': XLS,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': XLSX,
    'application/vnd.ms-powerpoint': PPT,
    'application/vnd.openxmlformats-officedocument.presentationml.presentation':
        PPTX,
    'application/vnd.oasis.opendocument.text': ODT,
    'application/vnd.oasis.opendocument.spreadsheet': ODS,
    'application/vnd.oasis.opendocument.presentation': ODP,
    'application/vnd.oasis.opendocument.graphics': ODG,
    'application/vnd.oasis.opendocument.formula': ODF        
}

###################
# Argument Parser #
###################

parser = argparse.ArgumentParser(
    description="Automated IMAP email checker and opener."
)

# Add a parser for sub-commands
subparsers = parser.add_subparsers(
    title = "Email-Checker Commands",
    dest = "cmd"
)

# Note: The -f FILE and -d options appear under all sub-command parsers so
# that the arguments are specified after the sub-command instead of having
# to be specified before the sub-command
# (i.e., 'script.py subcommand [arguments] -f FILE -d' instead of
# 'script.py -f FILE -d subcommand [arguments]')

# Add and configure 'check' sub-command
checkParser = subparsers.add_parser(
    'check',
    help = "Check an IMAP mailbox"
)
checkParser.add_argument(
    "-s", "--server", dest = "server", required = True,
    help = "IMAP email server"
)
checkParser.add_argument(
    "-u", "--username", dest = "username", required = True,
    help = "User's email address/login"
)
checkParser.add_argument(
    "-p", "--password", dest = "password", required = True,
    help = "User's password"
)
checkParser.add_argument(
    "-e", "--ssl", dest = "ssl",
    action = 'store_true', default = False,
    help = "Use IMAP over SSL"
)
checkParser.add_argument(
    "-mn", "--min-wait", dest = "waitMin",
    default = 5, type = int,
    help = "Minimum amount of time (in minutes) to wait before checking " \
            "email again"
)
checkParser.add_argument(
    "-mx", "--max-wait", dest = "waitMax",
    default = 10, type = int,
    help = "Maximum amount of time (in minutes) to wait before checking " \
           "email again"
)
checkParser.add_argument(
    "-k", "--keywords", dest = "keywords",
    default = "",
    help = "Comma-separated list of keywords. Enables social engineering " \
           "checks for 1 or more keywords before opening emails or clicking " \
           "links. Example: 'resume,job,opening,applicant'"
)
checkParser.add_argument(
    "-c", "--decode", dest = "decode",
    action = "store_true", default = False,
    help = "Decode the parameters (created during install)"
)
checkParser.add_argument(
    "-f", "--log-file", dest="logfile",
    default="email-checker.log",
    help = "Filename to save logging output to"
)
checkParser.add_argument(
    "-d", "--debug", dest="loglevel",
    action = "store_const", const = logging.DEBUG, default = logging.INFO,
    help="Log debug messages"
)

# Add and configure 'install' sub-command
installParser = subparsers.add_parser(
    'install',
    help = "Install a built email checker executable"
)
installParser.add_argument(
    "-f", "--log-file", dest="logfile",
    default="email-checker-install.log",
    help = "Filename to save logging output to"
)
installParser.add_argument(
    "-d", "--debug", dest="loglevel",
    action = "store_const", const = logging.DEBUG, default = logging.INFO,
    help="Log debug messages"
)

# Add and configure 'codebook' subcommand
codebookParser = subparsers.add_parser(
    'codebook',
    help = "Generate a new substitution cipher codebook"
)
codebookParser.add_argument(
    "-f", "--log-file", dest="logfile",
    default="email-checker.log",
    help = "Filename to save logging output to"
)
codebookParser.add_argument(
    "-d", "--debug", dest="loglevel",
    action = "store_const", const = logging.DEBUG, default = logging.INFO,
    help="Log debug messages"
)

# Parse the arguments
args = parser.parse_args()

####################
# Global Variables #
####################

if args.cmd == 'check':
    checkerRunEvent = threading.Event()
    checkerSleepEvent = threading.Event()
    checkerThread = None

    if args.decode:
        server = string.translate(args.server, DECODE)
        username = string.translate(args.username, DECODE)
        password = string.translate(args.password, DECODE)
        keywords = string.translate(args.keywords, DECODE).split(',')
    else:
        server = args.server
        username = args.username
        password = args.password
        keywords = args.keywords.split(',')
    
    ssl = args.ssl
    if args.waitMin > args.waitMax:
        checkParser.error("Min wait time cannot be greater than max wait " \
                          "time")
    waitMin = args.waitMin
    waitMax = args.waitMax

#####################
# Configure Logging #
#####################

logging.basicConfig(
    level = args.loglevel,
    filename = args.logfile,
    filemode = 'a',
    datefmt = DATE_FORMAT,
    format = FILE_LOG_FORMAT
)

# Configure console logging
console = logging.StreamHandler()
console.setLevel(args.loglevel)
consoleFormat = logging.Formatter(
    fmt = CONSOLE_LOG_FORMAT,
    datefmt = DATE_FORMAT
)
console.setFormatter(consoleFormat)
logging.getLogger('').addHandler(console)

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

########################
# Function Definitions #
########################

def build_exe(name):
    logging.debug("Generating executable")

    DEVNULL = open(os.devnull, 'w')

    # Generate the .spec file and executable using pyinstaller
    subprocess.call(
        ["pyinstaller", "--onefile", "--name=" + name, "email-checker.py"],
        stdout=DEVNULL, stderr=subprocess.STDOUT
    )

    exeName = name
    if platform.system() == 'Windows':
        exeName += '.exe'
    exeFile = os.path.abspath('dist/' + exeName)
    logging.info("Generated " + exeFile)

    return exeFile

def check_create_save_dir():
    # Check that SAVE_DIR exists.
    # If not, create it.
    savePath = os.path.abspath(SAVE_DIR)
    if not os.path.isdir(savePath):
        try:
            os.mkdir(savePath)
            logging.debug("Created save directory " + savePath)
            return savePath
        except OSError as err:
            logging.error("Error creating attachment save directory.")
            logging.error(str(err))
            # We need to return something
            # Get the current path and return it
            return get_path()
    else:
        return savePath

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
        subject = msg['subject'].lower()
        text = text.lower()
        # Cycle through the array of keywords
        for word in keywords:
            word = word.lower()
            # Compare lowercase keyword to lowercase text
            if word in sender or word in subject or word in text:
                # We can return True as soon as we find a match
                # We don't need to iterate through the rest of the keywords
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
    logging.info("--------------------------------------------------")
    logging.debug("Checking email...")

    imap = None
    try:
        # Connect to the server
        if ssl:
            imap = imaplib.IMAP4_SSL(server)
        else:
            imap = imaplib.IMAP4(server)

        logging.debug("Connected to IMAP server " + server)

    except Exception as err:
        logging.error("Failed to connect to IMAP server")
        logging.error(str(err))
        return
    
    try:    
        # Log in to the server
        retcode, data = imap.login(username, password)
        # Don't need to test retcode for 'OK' -- raises exception on failure
        logging.debug("Logged into IMAP server as " + username)
        
        # Select the default mailbox
        retcode, data = imap.select()
        # Open readonly for testing (don't mark messages as SEEN)
        #retcode, data = imap.select(readonly=1)
        
        # retcode != 'OK' --> failed to select the malbox
        # data <= 0      --> no messages in mailbox
        if retcode != 'OK' or int(data[0]) <= 0:
            # Either way, we can't continue
            if retcode != 'OK':
                logging.debug("Failed to select the mailbox")
            elif int(data[0]) <= 0:
                logging.debug("No messages in mailbox")
                # Close the mailbox
                retcode, data = imap.close()
                logging.debug("Closed the IMAP mailbox")

            retcode, data = imap.logout()
            logging.debug("Closed connection to IMAP server")
            return

        logging.debug("Selected the default mailbox (Inbox)")
        logging.debug(data[0] + " messages in mailbox")

    except Exception as err:
        logging.error(str(err))
        retcode, data = imap.logout()
        logging.debug("Closed connection to IMAP server")
        return

    try:
        # Search for UNSEEN messages
        retcode, data = imap.search(None, '(UNSEEN)')
        if retcode != 'OK':
            logging.debug("Search for UNSEEN emails failed")
            # Something went wrong. Return and clean up in finally block
            return

        # Convert the data string to an array
        msgIds = data[0].split(' ')
        if not any(msgIds):
            logging.info("Found 0 UNREAD emails.")
            # No messages to process. Return and clean up in finally block
            return
        else:
            logging.info("Found " + str(len(msgIds)) + " UNREAD emails.")

        # Try to fetch each returned msg ID
        logging.debug("Retrieving message(s): " + ', '.join(msgIds))
        for msgId in msgIds:
            logging.info("--------------------------------------------------")
            retcode, data = imap.fetch(msgId, '(RFC822)')
            if retcode != 'OK':
                # Couldn't fetch the message. Continue with the next one.
                logging.debug("Could not FETCH message " + str(msgId))
            else:
                # Message successfully received
                # Use the email module -- it's easier than parsing imaplib 
                # messages
                msg = email.message_from_string(data[0][1])
                logging.debug("Message " + str(msgId))
                logging.info("Subject: " + msg['subject'])
                logging.info("   From: " + msg['from'])

                # Perform keyword checks to determine if we'll process the 
                # message
                if check_keywords(msg):
                    # Passed keyword check (or no keywords defined)
                    # Process the message
                    process_message(msgId, msg)
                else:
                    # Failed keyword check -- skip this message
                    logging.debug("Skipping message " + str(msgId))

        logging.debug("--------------------------------------------------")

    except Exception as err:
        logging.error(str(err))

    finally:
        retcode, data = imap.close()
        logging.debug("Closed the IMAP mailbox")
        retcode, data = imap.logout()
        logging.debug("Closed connection to IMAP server")

def checker_loop():
    # Wait for the main thread to signal that it's ok to proceed
    checkerSleepEvent.wait()
    logging.debug("Checker thread running")

    while checkerRunEvent.is_set():
        check_mail()
        # Randomize how long we sleep for
        randSleep = random.randint(waitMin, waitMax)
        logging.debug("--------------------------------------------------")
        logging.info("Waiting " + str(randSleep) + " minute(s). Press " \
                     "[ENTER] to check now")
        # Sleep
        checkerSleepEvent.wait(randSleep * 60)

def clean_build(name):
    logging.debug("Cleaning build artifacts")

    try:
        shutil.rmtree("build")
        logging.debug("Removed 'build' directory")
    except shutil.Error as err:
        logging.error("Error removing 'build' directory")
        logging.error(str(err))

    try:
        shutil.rmtree("dist")
        logging.debug("Removed 'dist' directory")
    except shutil.Error as err:
        logging.error("Error removing 'dist' directory")
        logging.error(str(err))

    try:
        os.remove(name + '.spec')
        logging.debug("Removed '" + name + ".spec'")
    except shutil.Error as err:
        logging.error("Error removing '" + name + ".spec'")
        logging.error(str(err))

    logging.info("Removed build artifacts")

def generate_arg_string(arguments):
    if arguments['decode']:
        arguments['server'] = string.translate(arguments['server'], ENCODE)
        arguments['username'] = string.translate(arguments['username'], ENCODE)
        arguments['password'] = string.translate(arguments['password'], ENCODE)
        arguments['keywords'] = string.translate(arguments['keywords'], ENCODE)

    argString = "check " + \
                "-s " + arguments['server'] + " " + \
                ("-e " if arguments['ssl'] else "") + \
                "-u " + arguments['username'] + " " + \
                "-p " + arguments['password'] + " " + \
                "-mn " + str(arguments['waitMin']) + " " + \
                "-mx " + str(arguments['waitMax']) + " " + \
                ("-k " + arguments['keywords'] + " " 
                    if any(arguments['keywords']) else "") + \
                "-f " + arguments['logfile'] + " " + \
                ("-d " if arguments['debug'] else "") + \
                ("-c" if arguments['decode'] else "")

    return argString

def generate_filename(ext):
    counter = 0
    while True:
        # Attempt to pad to four digits.
        # Once we go above 10000, the number of digits will increase.
        filename = 'part-%04d%s' % (counter, ext)
        if os.path.exists(filename):
            counter += 1
        else:
            return filename

def get_arguments(path, name):
    while True:
        print
        print "Enter the IMAP server to use for checking mail."
        server = get_string_input("Server")
        ssl = get_yes_no("Use SSL?", False)

        print
        print "Enter the user's email credentials."
        username = get_string_input("Username")
        password = get_string_input("Password")

        print
        print "Specify the amount of time to wait between email checks (in " \
              "minutes)."
        while True:
            while True:
                waitMin = get_integer_input("Min Wait Time", 5)
                if waitMin < 1:
                    print "Error: Wait time must be at least 1 minute."
                else:
                    break
            while True:
                waitMax = get_integer_input("Max Wait Time", 10)
                if waitMax < 1:
                    print "Error: Wait time must be at least 1 minute."
                else:
                    break
            if waitMin > waitMax:
                print "Error: Min wait time cannot be greater than max wait " \
                      "time."
            else:
                break

        print
        print "Specify a comma-separated list of keywords that will trigger " \
              "the user to \"open\" an email. No keywords will result in " \
              "all emails being processed."
        keywords = raw_input2("Keywords: ")

        print
        fullLogPath = os.path.join(path, name + '.log')
        print "Specify the path and name of the log file."
        while True:
            logfile = get_string_input("Logfile", fullLogPath)
            logfile = os.path.abspath(logfile)
            logpath, logname = os.path.split(logfile)
            if os.path.isdir(logpath) and not os.path.isdir(logfile):
                break

        print
        print "Enabling debug output logs additional messages to assist " \
              "with troubleshooting."
        debug = get_yes_no("Enable debug output?", False)

        print
        print "Stored arguments such as server, username, password, and " \
              "keywords can be stored in an obfuscated format."
        decode = get_yes_no("Obfuscate arguments?", True)
        
        print
        print "The following arguments have been specified:"
        print "Server:    " + server
        print "Use SSL:   " + str(ssl)
        print "Username:  " + username
        print "Password:  " + password
        print "Min Wait:  " + str(waitMin)
        print "Max Wait:  " + str(waitMax)
        if any(keywords):
            print "Keywords:  " + keywords
        else:
            print "Keywords:  [None]"
        print "Log File:  " + logfile
        print "Debug:     " + str(debug)
        print "Obfuscate: " + str(decode)
        print
        if get_yes_no("Are the above settings correct?", True):
            break

    return {
        'server': server,
        'ssl': ssl,
        'username': username,
        'password': password,
        'waitMin': waitMin,
        'waitMax': waitMax,
        'keywords': keywords,
        'logfile': logfile,
        'debug': debug,
        'decode': decode
    }

def get_integer_input(prompt, default=None):
    if default:
        try:
            int(default)
        except ValueError:
            raise 
    input = None
    error = False

    while not input:
        if default:
            tmpPrompt = prompt + " [" + str(default) + "]: "
        else:
            tmpPrompt = prompt + ": "
        
        if error:
            tmpPrompt = "Invalid value. " + tmpPrompt

        input = raw_input2(tmpPrompt) or default
        try:
            int(input)
        except ValueError:
            input = None
            error = True

    return int(input)

def get_path():
    if getattr(sys, 'frozen', False):
       # we are running in a |PyInstaller| bundle
        return os.path.dirname(sys.executable)
    else:
        # we are running in a normal Python environment
        return os.path.dirname(os.path.abspath(__file__))

def get_string_input(prompt, default=None):
    input = None
    while not input:
        if default:
            input = raw_input2(prompt + " [" + default + "]: ") or default
        else:
            input = raw_input2(prompt + ": ")

    return input

def get_yes_no(prompt, default=None):
    valid = ['y', 'yes', 'n', 'no']
    if default == True:
        prompt = prompt + " ([Yes] / No): "
        defaultResponse = 'yes'
    elif default == False:
        prompt = prompt + " (Yes / [No]): "
        defaultResponse = 'no'
    else:
        prompt = prompt + " (Yes / No): "
        defaultResponse = ''

    while True:
        response = raw_input2(prompt) or defaultResponse
        # lower() to prevent case issues, set to default if no input
        response = response.lower()
        if response in valid:
            break

    # Typed 'y' or 'yes'
    if response == 'y' or response == 'yes':
        return True
    else:
        return False

def install_linux(name, installExe, argString):
    logging.debug("Installing Linux executable")

        # Create a .desktop file
    f = open(name + ".desktop", 'w')
    f.write("[Desktop Entry]\n")
    f.write("Type=Application\n")
    f.write("Exec=xterm -e " + installExe + " " + argString + "\n")
    f.write("Hidden=false\n")
    f.write("X-GNOME-Autostart-enabled=true\n")
    f.write("Name=" + name + "\n")
    f.write("Comment=")
    f.close()
    logging.info("Desktop launcher file created")

    homedir = os.path.expanduser("~")
    # Check whether the autostart directory exists
    if not os.path.isdir(homedir + "/.config/autostart"):
        # Create the autostart directory
        os.makedirs(homedir + "/.config/autostart")
        # Set the autostart directory permissions
        os.chmod(homedir + "/.config/autostart", 0700)
        logging.debug("Created directory " + homedir + "/.config/autostart")

    try:
        # Move the .desktop file to the autostart directory
        shutil.move(
            name + ".desktop", homedir + "/.config/autostart/"
        )
        logging.info(
            "Moved " + name + ".desktop to " + homedir + "/.config/autostart/"
        )
    except shutil.Error as err:
        logging.error(
            "Error moving " + name + ".desktop to " + homedir +
            "/.config/autostart/"
        )
        logging.error(str(err))

def install_windows(name, installExe, argString):
    logging.debug("Installing Windows executable")
    
    # Create the HKCU run key
    runKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    data = "\"" + installExe + "\" " + argString
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
    except WindowsError as err:
        logging.error("Error creating registry HKCU run key")
        logging.error(str(err))

def main():
    # Seed the random number generator
    # No argument seeds with the current system time
    random.seed()

    if args.cmd == 'check':
        run_check(args)
    elif args.cmd == 'install':
        try:
            run_install(args)
        except KeyboardInterrupt:
            return
    elif args.cmd == 'codebook':
        run_codebook(args)
    else:
        # Should never get here
        logging.error("Invalid email checker command: " + cmd)
        sys.exit(1)

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
    # The ext argument is only used when we know the ContentType from the part
    # header and no filename is specified (see save_attachment()). This script 
    # will always use the provided attachment filename if present. This may 
    # cause an attachment to attempt to open using the wrong application if 
    # the file extension does not match the file type.

    # Attempt to save the file and get it's returned filename
    filename = save_attachment(part, ext)
    if filename:
        # Save was successful. Attempt to open the file.
        open_file(filename)

    return

def process_html(htmlPart):
    # Example using BeautifulSoup
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

    filename = htmlPart.get_filename()
    if filename:
        # We have an HTML attachment
        # Attempt to save and open it.
        process_attachment(htmlPart, '.html')

def process_message(msgId, msg):
    logging.debug("Processing message " + str(msgId) + "...")

    # Uncomment below to save each email to a txt file for debugging
    #savePath = check_create_save_dir()
    #f = open(os.path.join(savePath, str(msgId) + ".txt"), 'w')
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
            # Handle some specific MIME types first
            if contentType == 'multipart/mixed':
                pass   # do nothing
            elif contentType == 'multipart/related':
                pass   # do nothing
            elif contentType == 'multipart/alternative':
                pass   # do nothing
            elif contentType == 'text/plain':
                # Can't do anything with a plain text email...
                logging.info(MIME_DICT[contentType][0])
            elif contentType == 'text/html':
                # Can look for links in HTML emails or html attachments
                logging.info(MIME_DICT[contentType][0])
                process_html(part)
            # The rest are file attachments. Search MIME_DICT.
            # If we find one we recognize, process the attachment.
            elif contentType in MIME_DICT:
                logging.info(MIME_DICT[contentType][0])
                process_attachment(part, MIME_DICT[contentType][1])
            # Octet-stream attachments from some tools
            # (e.g., Social Engineering Toolkit)
            elif contentType == 'application/octet-stream':
                process_attachment(part)
            else:
                # Unsupported/unknown ContentType
                pass   # do nothing
    else:
        # The message is not multipart -- it's palin text.
        logging.debug("Message is NOT multi-part.")
        # Not a whole lot we can do...
        logging.info("Got a plain text email. Not doing anything.")

def raw_input2(prompt=''):
    # Workaround for raw_input raising EOFError & KeyboardInterrupt on Ctrl-C
    try:
        return raw_input(prompt)
    except EOFError as err:
        # If KeyboardInterrupt not raised in 50ms, it's a real EOF event
        # Raise the EOFError
        time.sleep(.05)
        raise

def run_check(args):
    # Display some information about how we're configured
    print
    logging.info("==================================================")
    logging.info("Email checker started. Press [CTRL+C] to stop")
    logging.info("==================================================")
    logging.info("Checking email for: " + username)
    logging.debug("Using user password: " + password)
    logging.debug("Checking on server: " + server)
    if ssl:
        logging.debug("Using IMAP over SSL")
    logging.debug("Wait " + str(waitMin) + " to " + str(waitMax) + 
                  " minutes between checks")
    if any(keywords):
        logging.debug("Checking for " + str(len(keywords)) + " keywords: " + 
                      str(keywords))
    else:
        logging.debug("Not checking for keywords.")

    logging.debug("--------------------------------------------------")
    logging.debug("Creating checker thread")
    # Create the checker thread
    global checkerThread
    checkerThread = threading.Thread(target = checker_loop)
    # Set checker thread run event (signal to run)
    checkerRunEvent.set()
    # Start the thread
    checkerThread.start()

    # Sleep for 3 seconds to ensure checker thread
    # has time to create and start
    time.sleep(3)

    while True:
        # Signal checker thread to proceed
        checkerSleepEvent.set()
        # Clear the event so we only iterate once
        checkerSleepEvent.clear()
        
        try:
            # Wait for the user to hit [ENTER]
            raw_input2()
            # Old way:
            #sys.stdin.readline()
            # [CTRL+C] doesn't seem to work correctly unless something is 
            # written to stdout. Write something with a carriage return to 
            # return to the start of the line.
            # \1xb[2K -> escape sequence to clear the line
            #sys.stdout.write("\x1b[2K\r")
            #logging.info("User forced immediate email check")
        except KeyboardInterrupt:
            # The user hit [CTRL-C]
            # Break out of the loop so we can tear things down
            stop_checker_thread()
            break
        except:
            # Don't know what happened -- grab the exception and 
            # break out of the loop so we can tear things down
            exctype, value = sys.exc_info()[:2]
            logging.error("Unexpected error: " + exctype.__name__)
            stop_checker_thread()
            break

    logging.info("--------------------------------------------------")
    logging.info("Email checking terminated")

def run_install(args):
    print
    logging.info("==================================================")
    logging.info("Email-checker Installation")
    logging.info("==================================================")
    
    operatingSystem = platform.system()
    if operatingSystem not in PLATFORMS:
        logging.debug("Unsupported operating system: " + os + ". Installation aborted.")
        return

    frozen = getattr(sys, 'frozen', False)
    if frozen:
        logging.debug("Running from PyInstaller bundle")
    else:
        logging.debug("Not running from PyInstaller bundle")

    print
    print "This will generate a standalone executable from this script and " \
          "install it using a user-specified name and location. It will " \
          "also configure the operating system to start the executable with " \
          "the supplied arguments when the installing user account logs in."
    print
    print "NOTE: This may require administrator/root privileges if you " \
          "attempt to install to privileged locations. If you are not " \
          "running this script with administrator/root privileges, it is " \
          "STRONGLY recommended that you exit and run this script again " \
          "with the appropriate privileges or ensure the executable is " \
          "installed to an unprivileged location."
    print
    response = get_yes_no("Continue?", True)
    if not response:
        return

    print
    print "You can specify a name (e.g., Outlook, Firefox) for the " \
          "executable file so it blends in when viewed in a process " \
          "list. On Windows systems, this name is also used as the name " \
          "for the registry run-key value."
    print
    installName = get_string_input(
        "What should the file be named?",
        "email-checker"
    )

    cwd = get_path()
    print
    print "You can specify a location for the executable file. This can " \
          "make the program appear to be running from a more legitimate " \
          "location."
    print
    while True:
        installPath = get_string_input(
            "Where should the executable be located?",
            cwd
        )
        # Get the full/absolute path
        installPath = os.path.abspath(installPath or cwd)
        # verify we have a valid path
        if os.path.isdir(installPath):
            break

    print
    if not frozen:
        # Build the executable
        try:
            exeFile = build_exe(installName)
            exePath, exeName = os.path.split(exeFile)
        except:
            clean_build(installName)
            return

        # Move the generated file to it's final location
        try:
            shutil.move(exeFile, installPath)
            logging.info("Moved " + exeFile + " to " + installPath)
        except shutil.Error as err:
            logging.error("Error moving " + exeFile + " to " + installPath)
            logging.error(str(err))
            clean_build(installName)
            return
    else:
        exeFile = sys.executable
        exePath, exeName = os.path.split(exeFile)

        installExeAbsPath = os.path.join(installPath, installName)
        if operatingSystem == 'Windows':
            installExeAbsPath += '.exe'

        if os.path.exists(installExeAbsPath):
            # The PyInstaller bundle is located in the install directory
            # with the specified executable name
            logging.info(installExeAbsPath + " already exists")
        else:
            # The PyInstaller bundle is not located in the install directory
            # Copy the PyInstaller bundle to the specified location
            try:
                shutil.copy(exeFile, installPath)
                logging.info("Copied " + exeFile + " to " + installPath)
            except shutil.Error as err:
                logging.error(
                    "Error copying " + exeFile + " to " + installPath
                )
                logging.error(str(err))
                clean_build(installName)
                return

    arguments = get_arguments(installPath, installName)
    argString = generate_arg_string(arguments)

    print
    installExe = os.path.join(installPath, exeName)
    if operatingSystem == 'Linux':
        install_linux(installName, installExe, argString)
    elif operatingSystem == 'Windows':
        install_windows(installName, installExe, argString)

    if not frozen:
        clean_build(installName)

def run_codebook(args):
    print
    print "=================================================="
    print "Email-checker Codebook Generator"
    print "=================================================="

    if getattr(sys, 'frozen', False):
        logging.debug("Running from PyInstaller bundle")
        logging.error("Cannot generate a new codebook for a PyInstaller " \
                      "bundle")
        logging.error("Run the codebook generator using the Python script, " \
                      "replace the script's previous codebook with the new " \
                      "one, and build a new bundle with the new codebook.")
        return

    # Generate the new codebook
    tmpKeyspace = KEYSPACE
    newCodebook = ""
    # Interate the length of the keyspace
    for i in range(0, len(tmpKeyspace)):
        # Pick a random character in the keyspace
        pos = random.randint(0, len(tmpKeyspace) - 1)
        # Add that character to the codebook
        newCodebook += tmpKeyspace[pos]
        # Remove that character from the keyspace (so it's not used again)
        tmpKeyspace = tmpKeyspace[:pos] + tmpKeyspace[(pos + 1):]

    print
    print "Keyspace: " + KEYSPACE
    print "Codebook: " + newCodebook
    print
    print "Copy the new codebook above and replace the 'CODEBOOK' constant " \
          "in the email-checker Python script (email-checker.py) with the " \
          "new codebook value."
    print

def save_attachment(part, ext='.bin'):
    # The ext argument is only used if we know the ContentType from the part 
    # header and there is no filename specified. In this case, a generic 
    # filename with the specified extension will be created. The function will 
    # save the file to the directory specified in the 'SAVE_DIR' constant.
    # If the directory does not exist, it is created.

    logging.debug("Saving attachment...")

    savePath = check_create_save_dir()

    # Try to extract the filename from the part header
    filename = part.get_filename()
    if not filename:
        # Couldn't get the filename.
        # Use a generic one with the provided extension
        filename = generate_filename(ext)

    fileAbsPath = os.path.join(savePath, filename)
    if os.path.exists(fileAbsPath):
        # The filename was already used. Extract the file's extension.
        _, fileExt = os.path.splitext(filename)
        # Generate a new filename with the same extension
        filename = generate_filename(fileExt)
        fileAbsPath = os.path.join(savePath, filename)

    try:
        # Decode the email part and write it to the file
        f = open(fileAbsPath, 'wb')
        f.write(part.get_payload(decode=True))
        f.close()
    except Exception, err:
        logging.error("An error occurred while saving the file")
        logging.error(str(err))
        # Don't return a filename
        return None

    logging.debug("Attachment saved as \"" + fileAbsPath + "\"")
    # Return the filename of the saved file
    return fileAbsPath

def stop_checker_thread():
    if checkerThread and checkerThread.is_alive():
        # Clear thread run event (signal to quit)
        checkerRunEvent.clear()
        # Signal the checker thread proceed (in case it's waiting)
        checkerSleepEvent.set()
        # Join the thread (waiting for it to terminate)
        checkerThread.join()

################
# Main Program #
################

main()
