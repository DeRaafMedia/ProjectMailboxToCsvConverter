import mailbox
import csv
import sys

# Use "python mbox_to_csv_convert.py test.mailbox test.csv"
# where "test.mailbox" is the mbox to convert and "test.csv"
# is the name you want to give to the csv file.

mbox_file = sys.argv[1] ## Name of mbox_file
csv_file = sys.argv[2] ## Name of csv_file

# Function to determine charsets used
def determine_charsets(message):
    charsets = set({})
    for x in message.get_charsets():
        if x is not None:
            charsets.update([x])
    return charsets

# Error handeling
def error_handeling(error_message, email_message, charset):
    print()
    print(error_message)
    print('This error occurred while decoding with ',charset,' charset.')
    print('These charsets were found in the one email.',determine_charsets(email_message))
    print('This is the subject:',email_message['subject'])
    print('This is the sender:',email_message['From'])

# Get body of email message
def get_body_from_email(email_message):
    body = None

    # Walk through the parts of the email to find the text body
    if email_message.is_multipart():
        for part in email_message.walk():

            # If this is a multipart, walk through the su_part
            if part.is_multipart():

                for sub_part in part.walk():

                    if sub_part.get_content_type() == 'text/plain':

                        # Get the email body (.get_payload)
                        body = sub_part.get_payload(decode = True)

            elif part.get_content_type() == 'text/plain':

                body = part.get_payload(decode = True)

            # Some error handeling
            for x in determine_charsets(email_message):
                try:
                    body = body.decode(x)
                except UnicodeDecodeError:
                    error_handeling('UnicodeDecodeError encountered', email_message, x)
                except AttributeError:
                    error_handeling('AttributeError encountered',error_message, x)

            return body

# Create csv writer
csv_writer = csv.writer(open(csv_file, 'wb'))

# Some sort of header is handy (needs work!)
csv_writer.writerow(['Date:', 'Subject:', 'From:', 'To:', 'Contents:'])

# Walk through mbox and parse them to the csv_file
for email in mailbox.mbox(mbox_file):

    body = get_body_from_email(email)

    print body

    csv_writer.writerow([email["date"], email["subject"], email["from"], email["to"], unicode(body).encode('utf-8')])
