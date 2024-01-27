import win32com.client


def main():
    # Connect to Outlook
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")

    # Access the inbox
    inbox = outlook.GetDefaultFolder(6)  # 6 is the index for Inbox

    # Access all the messages in the inbox
    messages = inbox.Items

    # Loop through the messages and print them
    for message in messages:
        print("Subject:", message.Subject)
        print("Sender:", message.Sender)
        print("Received:", message.ReceivedTime)
        print("Email: ", message.Email )
        # Prints first 200 characters of the body
        print("Body:", message.Body[:200])
        print("-" * 50)  # Separator


if __name__ == "__main__":
    main()
