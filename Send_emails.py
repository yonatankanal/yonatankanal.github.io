import pandas as pd
import smtplib
from email.mime.text import MIMEText
import pprint
from email.mime.multipart import MIMEMultipart
from email import generator



def main():
    emails = make_data()
    # send_email(emails)
    make_email(emails)


def make_data():
    file_one_path = 'Files/spreadsheet_1.xlsx'
    file_two_path = 'Files/spreadsheet_2.xlsx'
    df_1 = pd.read_excel(file_one_path)
    df_cleaned_1 = df_1.dropna(subset=['Dates'])
    df_2 = pd.read_excel(file_two_path)
    df_cleaned_2 = df_2.dropna(subset=['Dates'])

    all_the_data = pd.merge(df_cleaned_1, df_cleaned_2, on=['First Name', 'Last Name'], how='inner')
    list_o_emails = []

    for index, row in all_the_data.iterrows():
        name = row['First Name']
        current = row['Practice Setting_x']
        future = row['Practice Setting_y']
        filler = input(f"Which email format for {name} (1, 2, or 3)? ")
        while filler != '1' and filler != '2' and filler != '3':
            filler  = input("***Input 1, 2, or 3*** Which email format for {name}? ")

        email = create_email(filler,name,current,future)
        move_on = input(f"\nTHAT WOULD LOOK LIKE THIS:\n-----{email}\n-----\nPress enter to move on to the next student. Press any character and then press enter to choose another format. ")
        while move_on != '':
            filler = input(f"Which email format for {name} (1, 2, or 3)? ")
            while filler != '1' and filler != '2' and filler != '3':
                filler  = input("***Input 1, 2, or 3*** Which email format for {name}? ")
            email = create_email(filler,name,current,future)
            move_on = input("Press enter to move on to the next student. Press any character and then press enter to choose another format. ")
        list_o_emails.append({'Email': email, 'Name': name})
    return list_o_emails


def create_email(filler,name,current,future):
    if filler == '1':
        body = f'''

Dear {name},

Congratulations on completing your first Level II Fieldwork at {current}! You made it! Although it may seem like you are starting over in some capacity at the {future}, remember that you have even more knowledge than you had 3 months ago. You’ve got this!

I hope this week goes well—please let me know how things are going. Remember that you can always reach out to myself or Ann with any questions or concerns.

Take care,
Avital
        '''
    
    elif filler == '2':
        body = f'''

Dear {name},

Can you believe you are already at midterm for your second Level II Fieldwork! I hope things are going well. What are some of the high points or low points you’ve experienced?

Feel free to reach out to either Ann or me with any questions or concerns. Enjoy the rest of your time at {current}.

Take care,
Avital
        '''

    elif filler == '3':
        body = f'''

Dear {name},

Congratulations on finishing your second Level II Fieldwork! How was it at {current}? You should feel very proud of your accomplishment. Have a wonderful summer and I’ll see you back here at BSP in August!

Take care,
Avital
        '''
    
    return body


def make_email(emails):
    i = 1
    for email in emails:
        message = MIMEMultipart()
        message['From'] = 'ASI14@pitt.edu'
        message['To'] = 'recipient@example.com'
        message['Subject'] = f'Email to be sent to {email['Name']}'
        body = email['Email']
        message.attach(MIMEText(body, 'plain'))

        with open(f'Emls/email{i}.eml', 'w') as eml_file:
            gen = generator.Generator(eml_file)
            gen.flatten(message)
            i += 1
    

def send_email(emails):
    for email in emails:
        subject = f'Email to be sent to {email['Name']}'
        body = email['Email']
        sender = '*PUT YOUR EMAIL ADDRESS HERE*' # ACTION REQUIRED
        recipients = ['*PUT YOUR EMAIL ADDRESS HERE*'] # ACTION REQUIRED
        password = '*PUT YOUR PASSWORD HERE*' # ACTION REQUIRED. This is your app password from Outlook, not your normal password (look online 
        # and it'll walk you through how to get this -- https://mailtrap.io/blog/outlook-smtp/)

        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = ', '.join(recipients)
        with smtplib.SMTP_SSL('smtp-mail.outlook.com', 587) as smtp_server:
            smtp_server.login(sender, password)
            smtp_server.sendmail(sender, recipients, msg.as_string())

        send_email(subject, body, sender, recipients, password)




if __name__ == "__main__":
    main()