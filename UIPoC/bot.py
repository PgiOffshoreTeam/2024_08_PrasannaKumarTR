from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server URL, the id of the task being executed,
    # the access token, and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = DesktopBot()
    bot.browse("https://mail.google.com")

    # Implement your logic here...
    if not bot.find("Compose", matching=0.97, waiting_time=10000):
        not_found("Compose")
    bot.click()

    if not bot.find("ToAddress", matching=0.97, waiting_time=10000):
        not_found("ToAddress")
    bot.click_relative(50, 10)  # Adjust the relative click position as needed
    bot.type_keys("anandgopalakrishna@gmail.com") 
    bot.type_keys("; pktr1999@gmail.com")  # Separate email addresses with a semicolon or comma

    if not bot.find("OnSubject", matching=0.97, waiting_time=10000):
        not_found("OnSubject")
    bot.click()

    bot.type_keys("Test Email - Configured using BotCity Computer Vision")  # Use bot.keyboard.paste to paste text

    if not bot.find("OnClickSend", matching=0.97, waiting_time=10000):
        not_found("OnClickSend")
    bot.click()

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finished OK."
    )

def not_found(label):
    print(f"Element not found: {label}")

if __name__ == '__main__':
    main()
