# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False
checkLogIn = False
def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = DesktopBot()
    bot.browse("https://www.automationexercise.com/")
    
    if not bot.find( "SignIn", matching=0.97, waiting_time=10000):
        not_found("SignIn")
    bot.click()
    
    if not bot.find( "Email", matching=0.97, waiting_time=10000):
        not_found("Email")
    bot.click_relative(42, 16)
    bot.type_keys("ui@gmail.com")
    
    if not bot.find( "Pass", matching=0.97, waiting_time=10000):
        not_found("Pass")
    bot.click_relative(36, 22)
    bot.type_keys("asdfghjkl")
    
    if not bot.find( "Login", matching=0.97, waiting_time=10000):
        checkLogIn = True
        not_found("Login")
    bot.click()
    
    if not bot.find( "View", matching=0.97, waiting_time=10000):
        not_found("View")
    bot.click()
    
    if not bot.find( "ItemCart", matching=0.97, waiting_time=10000):
        not_found("ItemCart")
    bot.click()
    
    if not bot.find( "ViewCart", matching=0.97, waiting_time=10000):
        not_found("ViewCart")
    bot.click()
    
    if not bot.find( "Checkout", matching=0.97, waiting_time=10000):
        not_found("Checkout")
    bot.click()
    
    if not bot.find( "OrderPlace", matching=0.97, waiting_time=10000):
        not_found("OrderPlace")
    bot.click()

    if not bot.find( "NameCard", matching=0.97, waiting_time=10000):
        not_found("NameCard")
    bot.click_relative(10, 16)
    bot.type_keys("John Doe")
    
    if not bot.find( "CardNum", matching=0.97, waiting_time=10000):
        not_found("CardNum")
    bot.click_relative(18, 14)
    bot.type_keys("1234567890123456")
    
    if not bot.find( "CVV", matching=0.97, waiting_time=10000):
        not_found("CVV")
    bot.click_relative(35, 21)
    bot.type_keys("123")
    
    if not bot.find( "ExpMon", matching=0.97, waiting_time=10000):
        not_found("ExpMon")
    bot.click_relative(32, 19)
    bot.type_keys("12")
    
    if not bot.find( "ExpYear", matching=0.97, waiting_time=10000):
        not_found("ExpYear")
    bot.click_relative(27, 16)
    bot.type_keys("2025")
    
    if not bot.find( "OrderConfirm", matching=0.97, waiting_time=10000):
        not_found("OrderConfirm")
    bot.click()
    
    if not bot.find( "Invoice", matching=0.97, waiting_time=10000):
        not_found("Invoice")
    bot.click()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )

def not_found(label):
    print(f"Element not found: {label}")
    
if __name__ == '__main__':
    main()






