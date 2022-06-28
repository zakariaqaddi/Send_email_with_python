import smtplib
from email.message import EmailMessage
import imghdr
import pandas as pd

email_subject = "Promo : Rechargez votre clim avec Point S"
sender_email_address = "****************"
email_smtp = "smtp.gmail.com"
email_password = "**************"

# Read the image
with open(r'C:\Users\departement-techniqu\OneDrive - Imperial pneu\Bureau\Codes\Send Emails with python\Image.jpg', 'rb') as file:
    image_data = file.read()

# Get the Emails from the excel sheet
Excel_data = pd.read_excel('Emails.xlsx', sheet_name='Emails')
Emails = Excel_data['Emails'].to_list()

for em in Emails:

    receiver_email_address = em
    # create an email message object
    message = EmailMessage()
    # configure email headers
    message['Subject'] = email_subject
    message['From'] = sender_email_address
    message['To'] = receiver_email_address
    message.set_content("""✨Profitez du forfait recharge clim chez Point S à partir de 299Dhs✨ *UN PARE-SOLEIL OFFERT POUR LES 100 PREMIERS CLIENTS A RECHARGER LA CLIMATISATION* ❄️              

     استفيدوا من عرضنا الحالي لإعادة صيانة مكيف سيارتكم😍             

    ➡️ https://www.point-s.ma/promotion/""")

    # attach image to email
    message.add_attachment(image_data, maintype='image',
                           subtype=imghdr.what(None, image_data))
    # set smtp server and port
    server = smtplib.SMTP(email_smtp, '587')
    # identify this client to the SMTP server
    server.ehlo()
    # secure the SMTP connection
    server.starttls()
    # login to email account
    server.login(sender_email_address, email_password)
    # send email
    server.send_message(message)

# close connection to server
server.quit()
