from botcity.plugins.excel import BotExcelPlugin

# Instantiate the plugin
bot_excel = BotExcelPlugin()

# Read from an Excel File
bot_excel.read('read.xlsx')
# Add a row
bot_excel.add_row([0, 22])
# Sort it by columns A and B in descending order
bot_excel.sort(['A', 'B'], False)

# Print the result
print(bot_excel.as_list())
# Save it to a new file
bot_excel.write('write.xlsx')