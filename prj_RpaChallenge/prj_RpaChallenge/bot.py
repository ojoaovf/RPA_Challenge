"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install -e .`
- Use the same interpreter as the one used to install the bot (`pip install -e .`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

from botcity.web import WebBot, Browser
from selenium.webdriver.common.by import By
import pandas as pd

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *

class Bot(WebBot):
    def action(self, execution=None):
        # Uncomment to silence Maestro errors when disconnected
        # if self.maestro:
        #     self.maestro.RAISE_NOT_CONNECTED = False

        # Configure whether or not to run on headless mode
        self.headless = False

        # Uncomment to change the default Browser to Firefox
        self.browser = Browser.CHROME

        # Uncomment to set the WebDriver path
        self.driver_path = r'C:\Users\joao.ferreira\Desktop\RPA Challenge BotCity\prj_RpaChallenge\prj_RpaChallenge\chromedriver.exe'
        self.download_folder_path = r'C:\Users\joao.ferreira\Desktop\RPA Challenge BotCity\Downloads'
        var_strCaminhoPlanilhaXlsx = self.download_folder_path + '\\challenge.xlsx'

        
        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id

        # Opens the RPA Challenge website and maximize the window
        self.browse("https://rpachallenge.com/")
        self.maximize_window()

        # Downloads the .xlsx doc containing the data
        self.find_element('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a', By.XPATH).click()
        self.wait(2000)
        
        # Read the .xlsx doc and allocate it in a variable
        var_df = pd.read_excel(var_strCaminhoPlanilhaXlsx, sheet_name='Sheet1', engine='openpyxl') #df = dataframe

        # Starts the challenge
        self.find_element('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button', By.XPATH).click()

        # Creates two auxiliar variables that will be used inside the following the loop
        var_intQtdeLinhas = len(var_df["First Name"])
        var_intContador = 0
        
        # Loop responsible for populate the fields and submit them
        for _ in range(0, var_intQtdeLinhas):
            
            var_strFirstName = var_df.loc[var_intContador, 'First Name']
            var_strLastName = var_df.loc[var_intContador, 'Last Name ']
            var_strCompanyName = var_df.loc[var_intContador, 'Company Name']
            var_strRoleInCompany = var_df.loc[var_intContador, 'Role in Company']
            var_strAdress = var_df.loc[var_intContador, 'Address']
            var_strEmail = var_df.loc[var_intContador, 'Email']
            var_strPhoneNumber = var_df.loc[var_intContador, 'Phone Number']

            var_weFirstName = self.find_element('//label[contains(text(), "First Name")]/following-sibling::input', By.XPATH).send_keys(var_strFirstName)
            var_weLastName = self.find_element('//label[contains(text(), "Last Name")]/following-sibling::input', By.XPATH).send_keys(var_strLastName)
            var_weCompanyName = self.find_element('//label[contains(text(), "Company Name")]/following-sibling::input', By.XPATH).send_keys(var_strCompanyName)
            var_weRoleInCompany = self.find_element('//label[contains(text(), "Role in Company")]/following-sibling::input', By.XPATH).send_keys(var_strRoleInCompany)
            var_weAddress = self.find_element('//label[contains(text(), "Address")]/following-sibling::input', By.XPATH).send_keys(var_strAdress)
            var_weEmail = self.find_element('//label[contains(text(), "Email")]/following-sibling::input', By.XPATH).send_keys(var_strEmail)
            var_wePhoneNumber = self.find_element('//label[contains(text(), "Phone Number")]/following-sibling::input', By.XPATH).send_keys(str(var_strPhoneNumber))

            var_intContador += 1

            self.find_element('/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input', By.XPATH).click()

        # Uncomment to mark this task as finished on BotMaestro
        # self.maestro.finish_task(
        #     task_id=execution.task_id,
        #     status=AutomationTaskFinishStatus.SUCCESS,
        #     message="Task Finished OK."
        # )

        # Wait for 10 seconds before closing
        self.wait(10000)

        # Stop the browser and clean up
      #  self.stop_browser()

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
