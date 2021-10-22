import imaplib
import email
import csv
import smtplib
import os

accounts_path = f"{os.path.dirname(os.path.realpath(__file__))}/email_accounts.csv"

with open(accounts_path) as f:
    account_list = list(csv.reader(f))

for index, item in enumerate(account_list):
    if index == 0:
        continue
    id, IMAP_server, email_id, pwd, port, mailbox, prev_id, forwardee = item

    ############### IMAP SSL ##############################
    with imaplib.IMAP4_SSL(host=IMAP_server, port=port) as imap_ssl:
        print("Connection Object : {}".format(imap_ssl))

        ############### Login to Mailbox ######################
        print("Logging into mailbox...")
        resp_code, response = imap_ssl.login(email_id, pwd)

        print("Response Code : {}".format(resp_code))
        print("Response      : {}\n".format(response[0].decode()))

        ############### Set Mailbox #############
        resp_code, mail_count = imap_ssl.select(mailbox=mailbox, readonly=True)

        ############### Retrieve Mail IDs for given Directory #############
        resp_code, mails = imap_ssl.search(None, "ALL")
        # print("Mail IDs : {}\n".format(mails[0].decode().split()))

        ############### Select unread email IDs #############
        mail_list = mails[0].decode('utf-8').split()
        max_id = len(mail_list)
        last_id = int(prev_id)
        max_count = min(max_id - last_id, 5)

        if max_count <= 0:
            print(f"All emails updated, no email to forward. {max_id} - {last_id}")
            print("\nClosing selected mailbox....")
            imap_ssl.close()
            continue
        else:
            unread_mails = mail_list[-max_count:]

        ############### Read each email #############
        for mail_id in unread_mails:
            print(f"================== Start of Mail [{mail_id}] ====================")
            resp_code, mail_data = imap_ssl.fetch(mail_id, '(RFC822)')  # Fetch mail data.
            message = email.message_from_bytes(mail_data[0][1])  # Construct Message from mail data
            print(message.get('From'))
            print(message.get('To'))
            print(message.get('Date'))
            print(message.get('Subject'))
            print(message.get_payload())
            print(f"================== End of Mail [{mail_id}] ====================\n")
            with smtplib.SMTP("smtp.gmail.com") as connection:
                connection.starttls()
                connection.login("mage.automailer@gmail.com", "Magical2019")
                message.replace_header('Subject', f"Forwarded from: {message.get('From')}")
                message.replace_header('From', message.get('To'))
                print(f"New Subject: {message.get('Subject')}, New sender: {message.get('From')}")
                connection.sendmail(
                    from_addr="mage.automailer@gmail.com",
                    to_addrs=forwardee,
                    msg=message.as_bytes())
                print('Message Sent!')

        ############# Close Selected Mailbox #######################
        print("\nClosing selected mailbox....")
        imap_ssl.close()
        ############# Update last read email ID and save to CSV #######################
        account_list[index][-2] = max_id


with open(accounts_path, 'w', newline='') as f:
    print("Saving new account list...")
    csv_writer = csv.writer(f)
    csv_writer.writerows(account_list)
    print("Account list updated")
