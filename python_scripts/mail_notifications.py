import pandas as pd
import win32com.client 
import os 
from datetime import datetime

current_time = datetime.now()
date_format = current_time.strftime('%Y-%m-%d')

filepath = 'C:\\Users\\aperalesc\\Desktop\\github\\python_mails'
attachments_filepath = 'C:\\Users\\aperalesc\\Desktop\\github\\python_mails\\attachments'

filename = 'data.xlsx'

# Connect with outlook
outlook = win32com.client.Dispatch('outlook.application')

def send_mails(usermail,username):

    mail = outlook.CreateItem(0)
    mail.To = usermail
    #mail.cc = mail - You can add any mail here if you want to copy someone
    
    mail.Subject = f'Observed Invoices - {username.upper()} - [{date_format}]'

    # In this part, we will 
    # While it's not necessary to do it in html, you can add some format like bold font, italics, etc.

    texto_html = f'''
    <!DOCTYPE html>
    <html lang="es">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body>
    <p>Dear User,</p>
    <p>You currently have {total_rows} invoice(s) under review. Below is the reason for the review, and an Excel file with the details of the observed documents is attached.</p>  
    {df_html}
    <p>Best regards,<br>
    <strong>Alessandro Perales</strong></p>
    </body>
    </html>
    ''' 
    mail.HTMLBody = texto_html

    # Add the excel created as attachment
    mail.Attachments.Add(attachments_fullpath)

    # Finally, send the email
    mail.Send()

df = pd.read_excel(os.path.join(filepath,filename),
                   sheet_name='invoices',  # It's always a good practice to declare the exact sheet name, just in case
                   dtype=str) # I prefer to manually change the type of my columns

# Select all unique mails
user_mails = set(df['mail'].values)

# We're going to send mails for all unique mails in our dataframe
for user_mail in user_mails:
    
    print(f'Sending mail to {user_mail}...')
    print('')

    # First, we select the rows 
    df_notification = df[df['mail']==user_mail]

    # Then, we select. In this particular case, we have the name in the table so an iloc function is ok
    # If you want to extract more information related to this user, you'll have to call a master table previously
    user_name = df_notification['user'].iloc[0]

    # Declaring attachment variables
    attachments_filename = f'Invoices - {user_name.upper()}.xlsx'
    attachments_fullpath = os.path.join(attachments_filepath,attachments_filename) 

    # Export the Dataframe as an Excel. We're going to add as an attachment
    df_notification.to_excel(attachments_fullpath,index=False)

    # We will need these inputs for our notification
    df_html = df_notification.to_html(index=False) # The dataframe in HTML
    total_rows = len(df_notification) # The number of rows, just for information

    # Call function
    send_mails(user_mail,user_name)

    # You can delete the file after you send it if you don't want to keep it
    os.remove(attachments_fullpath) 

print('Done!')