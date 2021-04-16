# -*- coding: utf-8 -*-
import pyfiglet.fonts
import smtplib
import ssl
import mimetypes
import numpy as np
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tkinter import *
import tkinter as tk
import six
import os.path
import tkinter.scrolledtext as scrolledtext
from pyfiglet import figlet_format
from email.mime.application import MIMEApplication
from tkinter import filedialog
try:
    from termcolor import colored
except ImportError:
    colored = None

    #cli_print("Connecting to SMTP Server  By U

styletagsfinal = """<style>
        @media only screen and (min-width: 768px){
            .main{
                width: 50%;
            }
            .header{
                width:100%;
            }
            .inside{
                width:100%;
            }
            img{
                width:100%
            }
            .left{
                margin-left:15px;
                width:145px;
                float:right;
            }
            .text{
                font-family: nunito;
                font-size:18.5px;
                text-align:justify;
            }
        }
        @media only screen and (max-width: 768px){
            .main{
            width: 100%;
            }
            .header{
            width:100%;
            }
            .inside{
            width:100%;
            }
            img{
            width:100%
            }
            .left{
            margin-left:10px;
            width:80px;
            float:right;
            }
            .text{
            font-family: nunito;
            font-size:15px;
            text-align:justify;
            }
        }
    </style>"""

bodytagsfinal = """<body>
    <div class="main">
        <img class="header" src="https://i.ibb.co/WG52VhC/recruiters.jpg" alt="Document">

        <div class='inside'>
            <div class="left">
                <img src="https://i.ibb.co/SXkjZVJ/image-2.png
https://i.ibb.co/WW7gQZS/image-3.png" alt="Image">
                <img src="https://i.ibb.co/WW7gQZS/image-3.png" alt="Image">
            </div>

            <p class='text'> Dear Manager<br>
                Greeting from <strong>Start@KMV</strong>,
                The Placement Cell Of <strong>Keshav Mahavidyalaya, University of Delhi</strong><br><br>

                We would like to inform you about the <strong>commencement
                of the Placement Season</strong> for the session 2020-21.<br><br> The
                Students of Keshav Mahavidyalaya are a diverse group of
                exceptional individuals interested in a variety of
                opportunities. Our students have been placed in
                prestigious companies like AT Kearney, EY, TresVista,
                Deloitte, ZS, Infosys, S&P Global & Gartner among others.<br><br>

                We wish to set new records this season and thus <strong>invite
                your esteemed organization</strong> for placement and internship
                opportunities. In the light of the current situation, we are
                also open to Virtual hiring drives.<br>
                <strong>PFA</strong> : Placement Brochure<br><br>

                For further information, contact-<br>
                Pranjal Kukreja - 9711376316<br>
                Riya Himmatramka - 9289997915<br>

                Thank You<br>
                Regards<br>
                Start@KMV<br>
                The Placement Cell<br>
                Keshav Mahavidyalaya<br>
                University of Delhi</p>
        </div>
    </div>

</body>
"""


login_details = dict()
to_emails = np.zeros(3)
details = dict()

PLAINT_TEXT_EMAIL = """
    Hi there,

    This message is sent from The Placement cell of Keshav Mahavidyalaya.

    Have a good day!
"""


def emails_to_be_sent(snum, enum, mailaddr, sheetlink, subject):
    global login_details
    try:
        database = pd.read_csv(sheetlink)
        lb.insert(END, "Database Loaded successfully \n\n")
        try:
            snum = int(snum)
            enum = int(enum)
            database = database.iloc[snum-1:enum, ::]
        #                starting: ending , 169, 200
            database["HR Name"] = database["HR Name"].fillna(
                "Hiring Manager")  # "" --> hiring manager
            details = dict(zip(database["Email"], database["HR Name"]))
            to_emails = database["Email"].values
            login_details = {'smtp_server': 'smtp.gmail.com', 'smtp_protocol': 'tls',
                             'example_no': '3', 'from_email': mailaddr, 'to_email': to_emails, 'subject': subject}

        except:
            lb.insert(END, "starting and ending numbers are not correct \n\n")
    except:
        lb.insert(END, "Please enter valid URL \n\n")
    return (login_details, details)


def get_html_message(from_email, to_email, subject, detailinfo, style, body):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    print(to_email)
    # Plain-text version of content
    plain_text = """\

        Greetings from Start@KMV, The Placement Cell of Keshav 
        Mahavidyalaya, University of Delhi.

        To continue on the path of providing novel opportunities to the 
        students, we are organizing the second edition of the Virtual 
        Internship Fair on February 28, 2021.

        Last year’s fair was an extreme success with 30+ esteemed 
        organizations offering a plethora of opportunities across diverse fields
        and witnessed more than a combined 6000+ student applications.

        The aim of the event is to bring employers from a variety of industries
        — Start-ups, MNCs and Social Organisations — hiring for various 
        roles of internship and apprenticeship positions. In the past,
        companies such as Sharekhan, Zomato Feeding India, Sanguine 
        Capital, Vivo, The MoneyRoller, Grant Thornton, Deloitte, NITI Aayog,
        Wipro and many more have participated and selected students for 
        internships.

        We wish to invite your esteemed organisation for the same and look 
        forward to a mutually beneficial relationship with your organisation.

        PFA: Brochure

        For further information, contact -
        Pranjal Kukreja - 9711376316
        Riya Himmatramka - 9289997915

        Thank You
        Regards
        Start@KMV
        The Placement Cell
        Keshav Mahavidyalaya
        University of Delhi
    """
    # html version of content
    html_content = """\
    <!DOCTYPE html>
        <html lang="en">
            <head>
                styletags
                <!-- <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css" integrity="sha384-TX8t27EcRE3e/ihU7zmQxVncDAy5uIKz4rEkgIXeMed4M0jlfIDPvg6uqKI2xXr2" crossorigin="anonymous">
                jQuery library--> 
                <!-- <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script> -->

                <!--Latest compiled and minified JavaScript--> 
                <!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script> -->


                <meta charset="UTF-8">
                <!-- <link rel="stylesheet" href="css/bootstrap.css"> -->
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Document</title>
            </head>
            bodytags
        </html>
        
    """
    html_content = html_content.replace("styletags", style, 1)
    html_content = html_content.replace("bodytags", body, 1)
    html_content = html_content.replace("Manager", detailinfo[to_email])
    text_part = MIMEText(plain_text, 'plain')
    html_part = MIMEText(html_content, 'html')
    msg.add_header('Content-Type', 'text/html')
    msg.attach(text_part)
    msg.attach(html_part)
    return msg


def get_attachment_message(from_email, to_email, subject, detailinfo, style, body, filename):
    msg = get_html_message(from_email, to_email, subject,
                           detailinfo, style, body)
    # Define MIMEImage part
    # Remember to change the file path
    file_path = filename
    ctype, encoding = mimetypes.guess_type(file_path)
    maintype, subtype = ctype.split('/', 1)
    pdf = MIMEApplication(open(file_path, 'rb').read())
    pdf.add_header('Content-Disposition', 'attachment', filename=os.path.splitext(
        filename)[0].split("/")[-1] + os.path.splitext(filename)[1])
    msg.attach(pdf)

    #print(msg)
    return msg


def cli_print(string, color, font="slant", figlet=False):
    if colored:
        if not figlet:
            six.print_(colored(string, color))
        else:
            six.print_(colored(figlet_format(
                string, font=font), color))
    else:
        six.print_(string)


def send_email(email_info, detailinfo, style, body, filename, server):
    example_no = email_info.get('example_no', '')
    from_email = email_info.get('from_email', '')
    to_email = email_info.get('to_email', '')
    subject = email_info.get('subject', '')

    try:
        cli_print("Sending your email...", "green")
        if filename == '':
            for x in to_email:
                msg = get_html_message(
                    from_email, x, subject, detailinfo, style, body)
                server.send_message(msg)
                cli_print("Your email was sent to " + x, "green")
                lb.insert(tk.END, "Your email was sent to " +
                          x + " Successfully \n\n")
        else:
            for x in to_email:
                msg = get_attachment_message(
                    from_email, x, subject, detailinfo, style, body, filename)
                server.send_message(msg)
                cli_print("Your email was sent to " + x, "green")
                lb.insert(tk.END, "Your email was sent to " +
                          x + " Successfully \n\n")
    finally:
        server.quit()


def main(start, end, mailaddr, passwd, sheetlink, subject, style, body, filename, server):
    """
    Simple CLI for sending emails with Python
    """
    cli_print("Sending Email CLI", color="blue", figlet=True)
    cli_print("*** Welcome to Sending Email With Python ***", "green")
    loginreturn = emails_to_be_sent(start, end, mailaddr, sheetlink, subject)
    email_info, detailinfo = loginreturn
    send_email(email_info, detailinfo, style, body, filename, server)


#main()
if __name__ == "__main__":
    smtp_server = 'smtp.gmail.com'
    port = 587
    server = smtplib.SMTP(smtp_server, port)
    filename = ""
    server = NONE
    # pass
    startNumber = 0
    endingNumber = 0

    def login(*args):
        global server
        smtp_server = 'smtp.gmail.com'
        protocol = 'tls'
        username = txtfldem.get()
        password = txtfldpass.get()

        context = ssl.create_default_context()

        cli_print("********************************************", "green")

        try:
            if protocol == 'ssl':
                port = 465
                server = smtplib.SMTP_SSL(smtp_server, port, context=context)
                #cli_print("Connecting to SMTP Server By Using SSL...", "green")
                server.login(username, password)
            elif protocol == 'tls':
                port = 587
                server = smtplib.SMTP(smtp_server, port)
                #cli_print("Connecting to SMTP Server  By Using TLS...", "green")
                # Secure the connection with TLS
                server.starttls(context=context)
                server.login(username, password)
                lb.insert(
                    tk.END, "Your have logged in with your account successfully " + username + "\n\n")
        except Exception as e:
            cli_print(
                "Could not connect to SMTP server with exception: %s" % e, "red")
            lb.insert(
                tk.END, "Could not connect to SMTP server with exception: %s  \n\n" % e)

    def ButtonClick(*args):
        startNumber = txtfldstarting.get()
        endingNumber = txtfldending.get()
        txtem = txtfldem.get()
        txtpass = txtfldpass.get()
        sheetlink = txtfldshtlink.get()
        subject = txtfldSubject.get()
        style = Tstyle.get('1.0', 'end-1c')
        body = Tbody.get('1.0', 'end-1c')
        # lb.insert(END, style,body)
        main(startNumber, endingNumber, txtem, txtpass,
             sheetlink, subject, style, body, filename, server)

    def uploadFile(*args):
        global filename
        filename = filedialog.askopenfilename()
        print(os.path.splitext(filename)[1])
        print(os.path.splitext(filename)[0].split("/")[-1])
        lb.insert(tk.END, "selected file is " + filename + "\n\n")

    window = Tk()
    window.title("Placement Fair")
    # window.iconbitmap("icon\Graphicloads-100-Flat-Email-2.ico")
    # w= Scrollbar(window);
    # w.pack( side = RIGHT, fill = Y )
    # w.config(command=window.yview)
    # add widgets here
    """This is for email and password and login button"""
    labelemail = Label(window, text="Enter the email here",
                       fg='black', font=("karla", 14))
    labelemail.pack()
    labelemail.place(x=60, y=50)

    txtfldem = Entry(window, text="This is email",
                     bg='white', fg='black', bd=5)
    txtfldem.pack()
    txtfldem.place(x=90, y=100)

    lblpass = Label(window, text="Enter the Password",
                    fg='black', font=("karla", 14))
    lblpass.pack()
    lblpass.place(x=60, y=150)

    txtfldpass = Entry(window, text="This is Password",
                       bg='white', fg='black', bd=5)
    txtfldpass.pack()
    txtfldpass.place(x=90, y=200)

    loginButton = tk.Button(window, text="Login", fg='black')
    loginButton.pack()
    loginButton.place(x=130, y=250)

    loginButton.bind('<Button-1>', login)

    #This is for the sheet link and sheet number
    labelsheetlink = Label(
        window, text="Enter the sheet link here", fg='black', font=("karla", 14))
    labelsheetlink.pack()
    labelsheetlink.place(x=350, y=50)

    txtfldshtlink = Entry(window, text="This is sheetlink",
                          bg='white', fg='black', bd=5)
    txtfldshtlink.pack()
    txtfldshtlink.place(x=400, y=100)

    labelSubject = Label(window, text="Enter the Subject here",
                         fg='black', font=("karla", 14))
    labelSubject.pack()
    labelSubject.place(x=350, y=150)

    txtfldSubject = Entry(window, text="This is Subject",
                          bg='white', fg='black', bd=5)
    txtfldSubject.pack()
    txtfldSubject.place(x=400, y=200)

    # this is for the starting and ending numbers
    labelstarting = Label(
        window, text="Enter the Starting Number", fg='black', font=("karla", 14))
    labelstarting.pack()
    labelstarting.place(x=690, y=50)
    txtfldstarting = Entry(
        window, text="This is Starting Number", bg='white', fg='black', bd=5)
    txtfldstarting.pack()
    txtfldstarting.place(x=750, y=100)

    lblending = Label(window, text="Enter the Ending Number",
                      fg='black', font=("karla", 14))
    lblending.pack()
    lblending.place(x=690, y=150)
    txtfldending = Entry(window, text="This is Ending Number",
                         bg='white', fg='black', bd=5)
    txtfldending.pack()
    txtfldending.place(x=750, y=200)

    btn = Button(window, text="Send Mails", fg='black')
    btn.pack()
    btn.place(x=475, y=500)

    btn.bind('<Button-1>', ButtonClick)

    lb = scrolledtext.ScrolledText(window, width=40, height=10, undo=True)
    lb['font'] = ('consolas', '12')
    lb.pack()
    lb.place(x=325, y=300)

    labelstyle = Label(
        window, text="Enter the complete style tag here", fg='black', font=("karla", 14))
    labelstyle.pack()
    labelstyle.place(x=75, y=500)
    Tstyle = tk.Text(window, height=10, width=50)
    Tstyle.pack()
    Tstyle.place(x=50, y=550)
    Tstyle.insert(tk.END, styletagsfinal)

    labelbody = Label(window, text="Enter the complete body here",
                      fg='black', font=("karla", 14))
    labelbody.pack()
    labelbody.place(x=600, y=500)
    Tbody = tk.Text(window, height=10, width=50)
    Tbody.pack()
    Tbody.place(x=550, y=550)
    Tbody.insert(tk.END, bodytagsfinal)

    #packing the widget to unpack it
    Tstyle.pack()
    Tbody.pack()
    labelstyle.pack()
    labelbody.pack()
    # This will remove the widget
    Tstyle.pack_forget()
    Tbody.pack_forget()
    labelstyle.pack_forget()
    labelbody.pack_forget()

    def hide_label():
        Tstyle.pack()
        Tbody.pack()
        labelstyle.pack()
        labelbody.pack()
        # This will remove the widget
        Tstyle.pack_forget()
        Tbody.pack_forget()
        labelstyle.pack_forget()
        labelbody.pack_forget()

# Method to make Label(widget) visible
    def recover_label():
        # This will recover the widget
        Tstyle.pack()
        Tbody.pack()
        labelstyle.pack()
        labelbody.pack()
        labelstyle.place(x=75, y=550)
        Tstyle.place(x=50, y=600)
        labelbody.place(x=600, y=550)
        Tbody.place(x=550, y=600)


# hide_label() function hide the label temporarily
    B2 = Button(window, text="Hide", fg="red", command=hide_label)
    B2.pack()
    B2.place(x=850, y=300)

# recover_label() function recover the label
    B1 = Button(window, text="Show", fg="green", command=recover_label)
    B1.pack()
    B1.place(x=850, y=350)

    uploadbutton = tk.Button(window, text="Select File")
    uploadbutton.place(x=120, y=300)
    uploadbutton.bind('<Button-1>', uploadFile)

    window.geometry("1000x800")
    window.mainloop()
