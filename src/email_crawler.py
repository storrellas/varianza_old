import poplib
import email

# Import credentials if any
try:
    from credentials import *
except ImportError as e:
    print("not found", e)

"""
server_pop = "pop.gmail.com"
server = poplib.POP3_SSL(server_pop)

server.user(user)
server.pass_(password)

resp, items, octets = server.list()
n_message = len(items)
for i in range(n_message):
    for raw in server.retr(i+1)[1]:
        print(raw)
        #msg = email.message_from_string(raw)
        #print(msg)


print(items)

server.quit()
"""
import imaplib

server_imap = 'imap.gmail.com'
mail = imaplib.IMAP4_SSL(server_imap)
mail.login(user, password)
#mail.list()
rv, mailboxes = mail.list()
if rv == 'OK':
    print("Mailboxes:")
    print(mailboxes)
mail.select('INBOX')

type, data = mail.search(None, 'ALL')
mail_ids = data[0]
id_list = mail_ids.split()
first_email_id = int(id_list[0])
latest_email_id = int(id_list[-1])

print(id_list)

# rv, data = mail.fetch(str(1), '(RFC822)')
# if rv != 'OK':
#     print("ERROR getting message", num)

#for i in range(latest_email_id, first_email_id, -1):
for i in [5]:
    typ, data = mail.fetch(str(i), '(RFC822)')

    print("NEEEEEEXXXXXXXXXXXXTTTTTTTTT")
    print("NEEEEEEXXXXXXXXXXXXTTTTTTTTT")
    print("NEEEEEEXXXXXXXXXXXXTTTTTTTTT")
    #print(data)
    for response_part in data:
        print("-- SEPARATOR --")
        print(response_part)
        """
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1])
            email_subject = msg['subject']
            email_from = msg['from']
            print('From : ' + email_from + '\n')
            print('Subject : ' + email_subject + '\n')
        """
