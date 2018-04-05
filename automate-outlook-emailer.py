
import win32com.client
import time
import threading
import re
from gsheets import Sheets
import pandas as pd

# Global variable flag
# flag to check for interrupt.
flag = threading.Event()


# Download and update the blacklist
def create_blacklist(url):
    '''
    Generate a list of emails from the Google sheets to blacklist from participation

    :param url: url containing google sheet with the users to blacklist from participation
    :return: return a list of emails to blacklist and exclude from emailing (if url not found, return an empty list)
    '''
    # Authorization
    sheets = Sheets.from_files('client_secrets.json', 'storage.json')
    # Fetch sheet by url
    s = sheets.get(url)
    if s is None:
        return []
    # list comprehension to flatten list
    flattened = [email.lower() for emails in s.sheets._items[0]._values for email in emails]
    return flattened


def create_mailing_list(url):
    '''
    Create a mailing list of just emails from a Google sheet with the database of volunteers and their information
    (Format may differ based on mailing list format and questions asked)

    :param url: url containing google sheet with the users to email
    :return: return list of user emails to contact (if url not found, return an empty list)
    '''
    sheets = Sheets.from_files('client_secrets.json', 'storage.json')
    s = sheets.get(url)
    if s is None:
        return []
    # list comprehension to flatten list
    flattened = [email.lower() for emails in s.sheets._items[0]._values[1:] for ind, email in enumerate(emails) if
                 ind == 1]

    return flattened


def read_email_body(email_html):
    '''
    Reads a file into a string object.

    :param email_html: html file containing email body to send
    :return: a read html-formatted email body
    '''
    initial_email_file = open(email_html, "r")
    initial_email = initial_email_file.read()
    return initial_email


def is_email(string):
    '''
    Basic check for correct syntax of an email

    :param string: email string to check
    :return: True if email format is correct and False if not
    '''
    if re.search(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", string):
        return True
    return False


def edit_email_body(date, time_slot):
    '''
    Generate an email html body that is catered to each user based on their day and time of appointment.
    Assumes you have a html file with the template, and with [DATE] and [TIME] in appropriate locations.

    :param date: string of date which user signed up for
    :param time_slot: string of time slot which user signed up for
    :return: None
    '''
    with open("confirmation_email_template.html", "rt") as fin:
        with open("confirmation_email.html", "w") as fout:
            for line in fin:
                fout.write(line.replace('[DATE]', date).replace('[TIME]', time_slot))


def get_subjects(bad_emails):
    '''
    This function uses the blacklisted emails and does more filtering and processing to subjects
    who have signed up for an appointment on the Google Form.
    Format of the Google Form should be similar to this:
    https://docs.google.com/forms/d/1IC6WzGsP7JvMVndYohK5qDUNuoutvCkxeB8PfkvmkAI/edit

    :param bad_emails: a list of strings containing blacklisted emails
    :return: returns a list of strings containing filtered emails that are ready to send
    '''
    response_list_url = raw_input("Enter the url of the Google sheet containing the responses "
                                  "to the Google survey with appointment times "
                                  "(2nd column should be of all emails): \n")

    sheets = Sheets.from_files('client_secrets.json', 'storage.json')
    s = sheets.get(response_list_url)
    if s is None:
        print("The Google sheet could not be found at the url provided. Please try again. \n")
        exit(1)

    response_list = s.sheets[0].to_frame().values.tolist()

    date_list = [response for response in s.sheets._items[0]._values[0]]
    date_list = date_list[3:8]
    list_of_subjects = []

    for elem in response_list:

        time_slot = []
        email_list = [current_list[0] for current_list in list_of_subjects]
        email = elem[1].lower()
        # if email found in blacklisted emails, skip this user
        if email in bad_emails or email in email_list:
            continue

        # Use for a full work week 5 days
        # check if no empty slots for date or is already confirmed
        if pd.isnull(elem[8]):
            # For every day of the week
            time_slot = elem[3:8]
            date_index = next((i for i, j in enumerate(time_slot) if not pd.isnull(j)), None)
            if date_index is None:
                continue
            elif date_index == 0:
                date = date_list[0]
            elif date_index == 1:
                date = date_list[1]
            elif date_index == 2:
                date = date_list[2]
            elif date_index == 3:
                date = date_list[3]
            elif date_index == 4:
                date = date_list[4]
            time_slot = [x for x in time_slot if not pd.isnull(x)]

            # If no time slot or more than one time slot, then skip user
            if len(time_slot) != 1:
                continue
            time_slot = ''.join(time_slot)
            list_of_subjects.append([email, elem[2], date, time_slot])

    return list_of_subjects


def user_confirm_email(email_html):
    '''
    Function for the user to confirm if the email is in the intended form.

    :param email_html: a path to the html file containing the body of email
    :return: True if users confirms, False if not
    '''
    with open(email_html, "r") as confirm_text_file:
        initial_email = confirm_text_file.read()

    email_volunteer("[PERSONAL EMAIL GOES HERE]", "Test Email", initial_email, True, True)
    works = raw_input("Does the email appear as you intended (y/n):\n")
    if works == "y":
        return True
    return False


def threaded_function():
    '''
    Call back function for threading which prompts user input to stop program

    :return: None to gracefully end the program
    '''
    stop_condition = raw_input("Enter any key to stop:\n")
    while stop_condition == '  ':
        stop_condition = raw_input("Enter any key other than '  ' to stop:\n")
    print "Stopping the emailer program..."
    flag.set()
    return


def email_volunteer(volunteer, subject, body, test=True, oversee=True):
    '''
    Using a running Microsoft Outlook application, this function emails an individual person
    with subject and html formatted body

    :param volunteer: string of volunteer email
    :param subject: string containing subject of email
    :param body: string containing html formatted body of the email message
    :param test: boolean value which indicates whether or not emailer is in test mode
    :param oversee: boolean value which indicates whether or not user wants to oversee sending emails one at a time
    :return: None
    '''

    emailer = win32com.client.Dispatch("Outlook.Application")

    # Create an email message.
    message = emailer.CreateItem(0)
    # If you have multiple emails on your outlook, email on behalf of other alternate email
    message.SentOnBehalfOfName = "OtherOutlookEmail@outlook.com"
    message.To = volunteer
    # Email to CC
    # message.CC = "example@gmail.com"
    message.Subject = subject
    message.htmlBody = body
    # Sends the message item. Once a message item is sent it cannot be sent again. Doing so would crash the program.
    if oversee:
        if not test:
            # We are not checking to confirm the email so we show it to the user.
            message.display()
            # Indirectly wait for 5 seconds, while checking to see if the thread threw a flag (i.e. flag is set & True)
            for i in range(5):
                if not flag.is_set():
                    time.sleep(1)
                    print i
            # No flags were set to True, so continue to send emails
            if not flag.is_set():
                message.Send()
            # Flag was set to True, so we prompt user to ask them what they want to do
            else:
                decision = raw_input("\nWould you like to skip this message(s), continue sending message (c), "
                                     "or exit (e): \n")

                if decision == "s":
                    message.To = "MyEmail@gmail.com"
                    message.Subject = "SKIPPED"
                    print ("Skipping")
                    message.Send()

                elif decision == "c":
                    message.Send()

                elif decision == "e":
                    message.To = "MyEmail@gmail.com"
                    message.Subject = "EXIT WAS CALLED"
                    message.Send()
                    exit(2)
        # Test mode to display the complete email before sending
        else:
            message.Display()
    else:
        message.Send()


def initial_email(mailing_list_url):
    '''
    This function defines the initial email to send for asking subjects if they are interested in participation by
    signing up via a Google Form located within this email.

    :param mailing_list_url: a string containing the url to the Google sheet that contains the mailing list
    :return: None to gracefully end program
    '''
    blacklist_url = raw_input("Enter the url of the Google sheet containing"
                               " a column/list of blacklisted emails"
                              "(If no blacklist, type None): \n")
    if blacklist_url != 'None':
        blacklist = create_blacklist(blacklist_url)
    else:
        blacklist = []

    mailing_list = create_mailing_list(mailing_list_url)
    all_emails = mailing_list

    try:
        with open("already_emailed.txt", 'r') as already_emailed_file:
            already_emailed = already_emailed_file.read()
            already_emailed = already_emailed.split('\n')

        emails = list(set(all_emails) - set(already_emailed))

    # Create file if does not exist (IOError)
    except IOError:
        emails = all_emails
        already_emailed = open("already_emailed.txt", 'w')

    emails = list(set(emails) - set(blacklist))

    # Basic email syntax check using regular expressions
    emails = [x for x in emails if is_email(x)]

    print("There are, " + str(len(emails)) + ' emails in the mailing list. The list starts and ends with the elements '
                                             'below:')
    try:
        print(emails[0], emails[len(emails) - 1])
    except IndexError:
        print("There are no emails in the given list or all users are blacklisted")
        exit(1)

    # email_html_body = raw_input("Enter the path to the html formatted body of email: \n")
    email_html_body = 'initial_email.html'
    # Reads the content that will compose the body of the email.
    initial_email = read_email_body(email_html_body)

    # Double check with user of script to check the html file and email to be sent to other user
    if not user_confirm_email(email_html_body):
        exit(1)

    # Ask the user if they would prefer to oversee the emailer bot (recommended use) or
    # send a chunk or all of emails instantly
    modular_sender = raw_input("Would you like to instantly send any number of emails(i) or oversee the emailing(o)? \n"
                               "Please enter either i or o respectively: ")
    # If we want to send the lump sum
    if modular_sender == 'i':
        # How many do you want to send
        number_to_send = input("Enter the number of emails you would like to send: ")
        # Error checking just in case user enters more emails than we have
        if number_to_send > len(emails):
            number_to_send = len(emails)

        # Send the emails
        for email in emails[:number_to_send]:
            print "Now sending email to %s..." % email
            email_volunteer(email, "Appointment", initial_email, False, False)
            with open("already_emailed.txt", "a") as out_file:
                out_file.write(email + '\n')
        exit(0)

    # Create a new thread with the function to execute threaded_function.
    thread = threading.Thread(target=threaded_function)
    # Set thread to be a daemon, which will allow the system to exit when the main function is completed,
    # regardless of whether or not the thread has finished its job.
    thread.daemon = True
    thread.start()

    for email in emails:

        if "stopped" in str(thread):
            # Thread has been stopped, so prompt user for next steps
            proceed = raw_input("Would you like to continue? y/n\n")
            if proceed == 'y':
                # If user wants to continue.
                # We need to complete the previous thread, and start a new one.
                # Additionally we should clear the flag that was set for later use.

                # Reset flag to false
                flag.clear()
                # Wait until thread terminates, this blocks calling thread until thread who calls join terminates
                thread.join()
                thread = threading.Thread(target=threaded_function)
                thread.daemon = True
                thread.start()
                time.sleep(1)
                email_volunteer(email, "Appointment", initial_email, False, True)

                with open("already_emailed.txt", "a") as out_file:
                    out_file.write(email + '\n')

            else:
                # Stop program by breaking out of loop
                print("Stopping the emailer program...\n")
                break
        else:
            time.sleep(1)
            print "Now sending email to %s..." % email
            email_volunteer(email, "Appointment", initial_email, False, True)
            with open("already_emailed.txt", "a") as out_file:
                out_file.write(email + '\n')


def confirm_email():
    '''
    This function is used when there exists a Google Form containing responses from subjects
    to confirm their appointment with a date and time.

    :return: None to gracefully end program
    '''
    blacklist_url = raw_input("Enter the url of the Google sheet containing"
                              " a column/list of blacklisted emails"
                              "(If no blacklist, type None): \n")
    if blacklist_url != 'None':
        blacklist = create_blacklist(blacklist_url)
    else:
        blacklist = []

    # all the processing and filtering of subjects is done in get_subjects function
    list_of_subjects = get_subjects(blacklist)

    thread = threading.Thread(target=threaded_function)
    # Set thread to be a daemon, which will allow the system to exit when the main function is completed,
    # regardless of whether or not the thread has finished its job.
    thread.daemon = True
    thread.start()

    for elem in list_of_subjects:

        # Just so it is clear:
        email = elem[0]
        name = elem[1]
        date = elem[2]
        time_slot = elem[3]

        # Replaces template email with actual date and time of appointment,and writes to a confirmation_email.html file
        edit_email_body(date, time_slot)
        confirmation_text = read_email_body("confirmation_email.html")

        if "stopped" in str(thread):
            # Thread has been stopped, so prompt user for next steps
            proceed = raw_input("Would you like to continue? y/n\n")
            if proceed == 'y':
                # If user wants to continue.
                # We need to complete the previous thread, and start a new one.
                # Additionally we should clear the flag that was set for later use.

                # Reset flag to false
                flag.clear()
                # Wait until thread terminates, this blocks calling thread until thread who calls join terminates
                thread.join()
                thread = threading.Thread(target=threaded_function)
                thread.daemon = True
                thread.start()
                time.sleep(1)
                email_volunteer(email, "Appointment", initial_email, False, True)

            else:
                # Stop program by breaking out of loop
                print("Stopping the emailer program...\n")
                break
        else:
            time.sleep(1)
            print "Now sending email to %s..." % email
            email_volunteer(email, "Appointment", initial_email, False, True)
            with open("already_emailed.txt", "a") as out_file:
                out_file.write(email + '\n')


def main():
    user_choice = raw_input("Is this the first initial email you are sending to the mailing list (i), "
                            "or a confirmation of appointment email(c)? "
                            "Please enter 'i' for initial or 'c' for confirmation. \n")
    if user_choice == 'i':
        mailing_list_url = raw_input("Enter the url of the Google sheet containing a mailing list "
                                     "with the second column being desired emails: \n")
        initial_email(mailing_list_url)
    elif user_choice == 'c':
        confirm_email()
    else:
        print("Please re-run the script, and enter i for sending initial emails to mailing list  "
              "or c for sending confirmation emails to scheduled participants. \n")
        exit(1)


if __name__ == "__main__":
    main()
