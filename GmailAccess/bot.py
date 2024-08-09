from botcity.plugins.gmail import BotGmailPlugin

# Set the path of the credentials json file
credentials = r"C:\PDF\client_secret_274358634089-sphd887k3d8td2bumfq6o4qg9716llrs.apps.googleusercontent.com.json"

try:
    # Instantiate the plugin
    gmail = BotGmailPlugin(credentials, "prasannapkkumar1999@gmail.com")

    # Search for all emails with subject: Hello World
    messages = gmail.search_messages(criteria="subject:Hello World")

    # For each email found: print the subject, sender address, and text content of the email
    for msg in messages:
        print("\n----------------------------")
        print(f"Subject => {msg.subject}")
        print(f"From => {msg.from_}")
        print(f"Msg => {msg.text}")

except Exception as e:
    print(f"An error occurred: {e}")