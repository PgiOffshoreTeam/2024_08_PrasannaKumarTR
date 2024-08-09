from botcity.plugins.files import BotFilesPlugin
from botcity.web import WebBot, Browser

# Instantiating the webbot
bot = WebBot()

# Setting the default browser
bot.browser = Browser.CHROME
# Setting the WebDriver path (use raw string to avoid issues with backslashes)
bot.driver_path = r"C:\Chromedriver-64\chromedriver.exe"

# Instantiating the plugin
files = BotFilesPlugin()

# Accessing a web page through the WebBot
#bot.browse("https://filesamples.com/formats/bin")

# Waiting for a new ".bin" file to be saved in the "downloads" folder
with files.wait_for_file(
    directory_path="C:\\",
    file_extension=".jpg",
    timeout=3000):
    print("\nDownloading file...")

    # Clicking to start download
    #if bot.find(label="download_file", matching=0.97, waiting_time=10000):
        #bot.click()

# Continuing the process after waiting for the file
print("\nDownload completed, continuing the process...")

# Getting the full path of the newest ".bin" file in the "downloads" folder
file_path = files.get_last_created_file(
    directory_path="C:\\",
    file_extension=".jpg")

print(f"Downloaded file path: {file_path}")