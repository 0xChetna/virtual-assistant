import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import speech_recognition as sr
import os
import pyttsx3
import webbrowser
import openai
import datetime
import wikipedia
import calendar

def say(text):



    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()



def take_command():
    # Initialize recognizer
    recognizer = sr.Recognizer()

    # Use microphone as source
    with sr.Microphone() as source:
        print("Listening...")

        # Adjust for ambient noise
        recognizer.adjust_for_ambient_noise(source)

        # Capture audio input
        audio = recognizer.listen(source)

    try:
        # Recognize speech using Google Speech Recognition
        print("Recognizing...")
        query = recognizer.recognize_google(audio)

        # Print what was heard
        print(f"User said: {query}")

    except sr.UnknownValueError:
        # Handle errors when speech cannot be recognized
        print("Sorry, I didn't get that. Please try again.")
        return "Sorry, I didn't get that. Please try again."

    return query


def send_email(subject, message, to_email, from_email, password, smtp_server='smtp.gmail.com', smtp_port=587):
    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Add message body
    msg.attach(MIMEText(message, 'plain'))

    # Connect to SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(from_email, password)

    # Send email
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)

    # Close the connection
    server.quit()


def send_email_assistant(from_email, password):
    try:
        # Prompt for recipient's email address
        say("Enter the recipient's email address: ")
        to_email = input("Enter the recipient's email address: ")

        # Prompt for email subject
        say("Enter the email subject: ")
        subject = input("Enter the email subject: ")

        # Prompt for email content
        say("Enter the email content: ")
        content = input("Enter the email content: ")

        send_email(subject, content, to_email, from_email, password)
        print("Email sent successfully!")
        say("Email sent successfully!")
    except Exception as e:
        print("An error occurred while sending the email:", e)
        say("Sorry, I couldn't send the email. Please try again later.")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        say("Good Morning!")
    elif hour>=12 and hour<18:
        say("Good Afternoon!")
    else:
        say("Good Evening!")
def date():
    now = datetime.datetime.now()
    month_name = now.month
    day_name = now.day
    month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                   'November', 'December']
    ordinalnames = ['1st', '2nd', '3rd', ' 4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th',
                    '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd', '24rd', '25th',
                    '26th', '27th', '28th', '29th', '30th', '31st']

    say("Today is " + month_names[month_name - 1] + " " + ordinalnames[day_name - 1] + '.')

def display_calendar(year, month):
    cal = calendar.month(int(year), int(month))
    say(f"Calendar for {calendar.month_name[int(month)]} {year}:")
    print(cal)

if __name__ == '__main__':
    print('PyCharm')

    print(wishMe())
    say("hello I m jarvis AI. Please tell me how may I help you")
    while True:

        command = take_command()

        #to open site
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"], ["stackoverflow" , "https://stackoverflow.com/"]]
        for site in sites:

            if f"Open {site[0]}".lower() in command.lower():
                say(f"opening {site[0]} mam")
                webbrowser.open(site[1])

        #to tell date
        if 'date' in command.lower():
            date()

        # identity of virtual assistant
        if 'who are you' in command.lower() or 'what can you do'  in command.lower():
            say(
                'I am  your personal assistant. I am programmed to minor tasks like opening youtube, google chrome, gmail and search wikipedia etcetra')
            print('I am  your personal assistant. I am programmed to minor tasks like opening youtube, google chrome, gmail and search wikipedia etcetra')

        # to search on wikipedia
        if "search on wikipedia".lower() in command.lower():
            say("searching in wikipedia")
            res= wikipedia.summary(command, sentences = 2)
            print(res)
            say("according to wikipedia")
            say(res)

        # to search on google
        if "search on google".lower() in command.lower():
            say("searching on google")
            say("Enter your search query: ")
            search_query = input("Enter your search query: ")
            say("you want to search about " )
            say( search_query)
            say("searching..")
            search_url = "https://www.google.com/search?q=" + "+".join(search_query.split())
            webbrowser.open_new_tab(search_url)

        # to tell the current time
        if "the time" in command:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"the time is {strfTime}")
            say(f" The time is {strfTime}")

        # to open excel
        if "open excel".lower() in command.lower():
            codepath1= "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.exe"
            os.startfile(codepath1)
            say("opening excel ")
            print("opening excel.... ")

        # to open word
        if "open word".lower() in command.lower():
            codepath2= "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.exe"
            os.startfile(codepath2)
            say("opening microsoft word ")
            print("opening microsoft word.... ")

         # to open brave browser
        if "open brave".lower() in command.lower():
            codepath3= "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
            os.startfile(codepath3)
            say("opening brave browser")
            print("opening brave browser....")

        # to open powerpoint preaentation
        if "open power".lower() in command.lower():
            codepath4= "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.exe"
            os.startfile(codepath4)
            say("opening powerpoint presentation")
            print("opening powerpoint presentation.... ")

        # to send email
        if "send email" in command.lower():
            # Assuming email credentials are defined elsewhere in your code
            send_email_assistant("chetnadixit79@gmail.com", "ujcq rweg vpvg pchr")

        # to tell the news
        if "news" in command.lower():
            say("showing today's news")
            say("here are some headlines from the todays newspaper from hindustan times, happy reading")
            webbrowser.open("wwwhindustantimes.com")

        # to locate some place on google maps
        if "where is" in command.lower():
            location = command.replace("where is", "")
            location = location.strip()  # Remove leading/trailing whitespaces
            print("Location:")
            say("opening gge maps")
            webbrowser.open("https://www.google.com/maps/place/" + location)

        # to print calendar
        if "calendar" in command.lower():
            say("Enter the month for which you want to see the calendar:")
            month = input("Enter the month (as a number): ")
            say("Enter the year for which you want to see the calendar:")
            year = input("Enter the year: ")
            display_calendar(year, month)

        # to stop running virtual assistant
        if "close it" in command.lower():
            say("goodbye ! have a nice day")
            print("goodbye ! have a nice day")
            time.sleep(2)
            sys.exit()




