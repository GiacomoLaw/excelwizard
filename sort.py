import win32com.client
import os
import pandas as pd

# outlook object creation
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# search through inbox
folder = outlook.GetDefaultFolder(6)

# get latest items
emails = folder.Items
emails.Sort("[ReceivedTime]", True)

for email in emails:
    # define email subject to search for
    if email.Subject == "CSV file from Report Wizard":
        subject = email.Subject
        print(f"Subject: {subject}")

        # attachment check
        if email.Attachments.Count > 0:
            # save to temp file
            attachment = email.Attachments.Item(1)
            attachment_filename = os.path.join(os.getcwd(), attachment.FileName)
            attachment.SaveAsFile(attachment_filename)
            print(f"Saved attachment: {attachment_filename}")

            # add to dataframe with windows encoding
            df = pd.read_csv(attachment_filename, encoding='cp1252')

            # sort criteria
            df.sort_values(by=['Order No', 'Order Line', 'Date Entered'], inplace=True)

            # copy dataframe to clipboard
            df.to_clipboard(index=False, header=True, sep='\t')

            print(f'Successfully sorted and copied to the clipboard with columns.')

            os.remove(attachment_filename)
        else:
            print("No attachments found in the email.")

        break

# release outlook
del outlook
