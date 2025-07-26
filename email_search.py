import win32com.client 

def search_emails_for_string(search_string):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 just refers to the Inbox folder

    for item in inbox.Items:
        if hasattr(item, "Body") and search_string.lower() in item.Body.lower():
            print("Found in email subject:", item.Subject)

search = input("search: ")
search_emails_for_string(search)