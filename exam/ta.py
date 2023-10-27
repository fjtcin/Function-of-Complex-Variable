import pandas as pd
import smtplib

'''
Change these to your credentials and name
'''
your_name = "张学涵"
your_email = "fjtcin@mail.ustc.edu.cn"
your_password = ""

# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below

# Read the file
email_list = pd.read_excel("EmailList.xlsx")

# Get all the Names, Email Addreses, Subjects and Messages
all_names = email_list['Name']
all_emails = email_list['Email']
all_total = email_list['Total']
all_p1 = email_list['P1']
all_p2 = email_list['P2']
all_p3 = email_list['P3']
all_p4 = email_list['P4']
all_p5 = email_list['P5']
all_p6 = email_list['P6']

# Loop through the emails
for idx in range(len(all_emails)):
    server = smtplib.SMTP('mail.ustc.edu.cn', 25)
    server.ehlo()
    server.login(your_email, your_password)

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    subject = f'复变函数期中成绩_{name}'
    message = f'总分：{all_total[idx]}\n第一题：{all_p1[idx]}/20\n第二题：{all_p2[idx]}/20\n第三题：{all_p3[idx]}/20\n第四题：{all_p4[idx]}/20\n第五题：{all_p5[idx]}/20\n第六题：{all_p6[idx]}/10'

    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message).encode('utf-8'))

    # In the email field, you can add multiple other emails if you want
    # all of them to receive the same text
    try:
        server.sendmail(your_email, [email], full_email)
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

    # Close the smtp server
    server.close()