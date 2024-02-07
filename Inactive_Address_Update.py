from selenium import webdriver
from selenium.webdriver.common.keys import Keys # Give access to keys like enter and esc key to see all search results
from selenium.webdriver.common.by import By # Used to locate elements within a website
from selenium.webdriver.chrome.options import Options # Also used to select options in dropdown menu
from selenium.webdriver.support.wait import WebDriverWait # Used for Implicit and Explicit Waits
from selenium.webdriver.support import expected_conditions as EC # Something also used for Waits
from selenium.webdriver.support.select import Select # Used for selecting options in a dropdown menu
from selenium.webdriver.common.action_chains import ActionChains # Used in clearing out username before putting in login info
from screeninfo import get_monitors
import pandas as pd # Turns Excel sheet into a dataframe
import tkinter
from tkinter import filedialog # Used to open up file dialog
import pyautogui # Used primarily for message/confirm boxes
import pymsgbox # Formatting for confirm boxes
import time

# ---------------- Monitor resolution & display default message box ----------------------
primary_mon_width = 0
primary_mon_height = 0
primary_mon_x = 0

for m in get_monitors(): # https://stackoverflow.com/questions/3129322/how-do-i-get-monitor-resolution-in-python
    # print(str(m))
    if m.is_primary == True:
        primary_mon_width = m.width
        primary_mon_height = m.height
        primary_mon_x = m.x

print('\nprimary_mon_width: ' + str(primary_mon_width))
print('primary_mon_height: ' + str(primary_mon_height))
print('primary_mon_x: ' + str(primary_mon_x))

# Changes default message box location to a third of height and width of primary display
pymsgbox.rootWindowPosition = '+' + str(int(primary_mon_width/3)) + '+' + str(int(primary_mon_height/3))

# ---------------- Options ----------------------
options = Options()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--window-position=' + str(primary_mon_x) + ',0')
options.page_load_strategy = 'normal'
driver = webdriver.Chrome(options=options)

# ---------------- Open up website ----------------------
driver.get("https://discover.highpoint.edu/manage/")
driver.maximize_window()
print('\nDriver Title: ' + driver.title)

# ------------------ FUNCTION FOR LOGGING OUT ----------------------
def Log_Out():
    print('\nLogging Out...')

    pyautogui.alert(text='Logging out...', button='OK')

    avatar = driver.find_element(By.CLASS_NAME, 'user_menu_icon')
    avatar.click()

    # EXPLICIT WAIT UNTIL LOGOUT IS CLICKABLE
    try:
        element = WebDriverWait(driver, 10)
        element.until(EC.element_to_be_clickable((By.LINK_TEXT, 'Logout')))
    except TimeoutError:
        print('Timeout Error... Quitting program')
        exit()

    logout = driver.find_element(By.PARTIAL_LINK_TEXT, 'Logout')
    logout.click()
    print('Logging Out Complete :)')

# ------------------ FUNCTION FOR MATCHING REF #'S (Used in next step) ----------------------
def Match_Ref_Num(given_ref_num):
    try:
        final_num = dataframe1[dataframe1['Ref'] == given_ref_num].index[0]
    except:
        return 'Not Matching'
    else:
        return str(final_num)

# --------------------- Sign into login -----------------------------
def Login_Attempts(attempts):
    attempt = 0
    for attempt in range(attempts): # https://stackoverflow.com/questions/2083987/how-to-retry-after-exception
        # IF TOO MANY ATTEMPTS, EXIT
        if attempt == attempts-1:
            print('ERROR: TOO MANY ATTEMPTS...')
            exit(0)
        
        username = ''
        password = ''

        username_window = pyautogui.prompt(text='Please enter your username for HPU\'s Login Page.',
                title='Enter Username')
        if str(username_window) == 'None':
            print('Cancel Button Pressed...')
            exit(0)

        else:
            print('OK Button Pressed...')
            username = str(username_window)

        password_window = pyautogui.password(text='Please enter your password for HPU\'s Login Page.',
                                            title='Enter Password',
                                            mask='#')
        if str(password_window) == 'None':
            print('Cancel Button Pressed...')
            driver.quit()
            exit(0)
        else:
            password = str(password_window)

        # EXPLICIT WAIT UNTIL LOGOUT IS CLICKABLE
        try:
            element = WebDriverWait(driver, 10)
            element.until(EC.element_to_be_clickable((By.ID, 'userNameInput')))
        except TimeoutError:
            print('Timeout Error... Quitting program')
            exit()

        search = driver.find_element(By.ID, 'userNameInput')
        ActionChains(driver).move_to_element(search).click().perform()
        ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        ActionChains(driver).key_up(Keys.CONTROL).perform()
        #search.send_keys(username, Keys.TAB, password, Keys.ENTER)
        search.send_keys('clapiz', Keys.TAB, 'Mich7890!@#$', Keys.ENTER) # TODO: CHANGE BACK TO COMMENTED CODE ABOVE

        try:
            driver.find_element(By.ID, 'footer_school')
            break
        
        except:
            continue # Retry loop

    print('Number of additional attempts needed: ' + str(attempt))
    print('Login Complete')
Login_Attempts(11) # 10 Login attempts allowed

# ------------------- Reading through excel sheet ----------------------------
# dataframe1 = pd.read_excel('Look Through.xlsx', dtype=object) # Modify this setting in VS Code: File > Preferences > Settings > Python > Data Science > Execute in File Dir
def File_Access_Attempts(attempts):
    attempt = 0
    path = ''

    for attempt in range(attempts):
        # IF TOO MANY ATTEMPTS, EXIT
        if attempt == attempts-1:
            print('ERROR: TOO MANY ATTEMPTS...')
            Log_Out()
            exit(0)

        try:
            tkinter.Tk().withdraw() # prevents an empty tkinter window from appearing
            path = filedialog.askopenfilename() # https://stackoverflow.com/questions/66663179/how-to-use-windows-file-explorer-to-select-and-return-a-directory-using-python
            dataframe1 = pd.read_excel(str(path))
            print(dataframe1)
        except:
            print('File is not valid, retrying...')
            continue
        else:
            break
        
    print('Number of additional attempts needed: ' + str(attempt))
    print('File Access Complete')
    return dataframe1
dataframe1 = File_Access_Attempts(4) # 3 File open attempts

# ----------------- Prompt window for which ref # to start on (stored as a string) ------------------------
def Ref_Search(searches):
    for search_attempt in range(searches):
        # If no more attempts remaining
        if search_attempt == searches-1:
            print('ERROR: TOO MANY ATTEMPTS...')
            exit(0)
        
        prompt_window = pyautogui.prompt(text='Type the Reference Number (Ex: 744984288) to start off with.' + 
                                        '\nTo start with first entry, leave empty and press OK.' + 
                                        '\nTo stop, press Cancel')
        if str(prompt_window) == 'None': # If CANCEL button is pressed
            print('\nPrompt Window #1: Cancel Button Pressed')
            Log_Out()
            exit(0)
        
        elif str(prompt_window) == '':
            print('\nPrompt Window #1: OK Button Pressed')
            ref_row = 0
            return ref_row

        # Checks if returned string is all numbers
        elif str(prompt_window).isdecimal():
            print('\nPrompt Window #1: OK Button Pressed\nIs a Decimal...')

            if Match_Ref_Num(prompt_window) == 'Not Matching':
                print('ref_row: No matching Reference Number')
                print('Retrying...')
                continue

            else:
                print('ref_row: ' + ref_row)
                ref_row = Match_Ref_Num(prompt_window)
                return ref_row

        else: # If something has gone wrong, just default from starting at 0
            print('\nPrompt Window = ' + str(prompt_window))
            print('Error: Incorrect Reference Number Type.')
            Log_Out()
            exit(0)
ref_row = Ref_Search(11) # 10 search attempts allowed


# -------------- For loop to go through each row and the data from each row --------------------
for index, row in dataframe1.iloc[int(ref_row):].iterrows(): # Start from ref row # to the end
    # Store Full Name and Reference Number as String
    full_name = str(row['Last Name']) + ', ' + str(row['First Name'])
    reference_num = str(row['Ref'])

    # Check for if ref # has 9 digits, if not add a 0 at the end
    length_ref = len(reference_num)
    if length_ref < 9:
        for i in range(9 - length_ref):
            reference_num = '0' + reference_num

    # ----------------- Confirm Window #2: Continue to next person or stop ---------------------
    # Confirm Window response stored as a string
    confirm_window = pyautogui.confirm(text='Press OK to Continue and Cancel to stop.' + 
                                        '\nCurrent Reference Number: ' + reference_num + 
                                        '\nCurrent Person: ' + full_name + 
                                        '\n\nContinuing in 10 seconds...',
                                        title='Continue?', buttons=['OK', 'Skip Person', 'Cancel'])
    # print('CONFIRM_WINDOW: ' + confirm_window)

    if str(confirm_window) == 'Skip Person':
        print('\nPrompt Window #2: Skip Person Button Pressed')
        continue
    if str(confirm_window) == 'Cancel':
        print('\nPrompt Window #2: Cancel Button Pressed')
        break

    if str(confirm_window) == '':
        print('\nPrompt Window #2: OK Button Pressed')

    # ----------------- Add Name to Search Bar ---------------------
    search_bar = driver.find_element(By.ID, 'qs_suggest')
    search_bar.send_keys(reference_num, Keys.ENTER)
    print('Reference Number: ' + reference_num + '\nFull Name: ' + full_name + '\nAdded to search bar...')
    
    # EXPLICIT WAIT UNTIL ID ON HOVER ROW IS CLICKABLE
    try:
        element = WebDriverWait(driver, 10)
        element.until(EC.element_to_be_clickable((By.CLASS_NAME, 'table'))) # !!! For some reason having 2 (()) after "element_to_be_clickable" was important !!!
    except:
        print('Some Error Occurred... Logging Out')
        Log_Out()

    # Press tab 9 times and then press enter
    match_bar = driver.find_element(By.ID, 'search_quick')
    match_bar.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, 
                         Keys.TAB, Keys.TAB, Keys.ENTER)

    try:
        element = WebDriverWait(driver, 10)
        element.until(EC.element_to_be_clickable((By.ID, 'part_profile_link')))
    except:
        print('Some Error Occurred... Logging Out')
        Log_Out()

# ----------------- Go into profile tab ---------------------
    profile_click = driver.find_element(By.ID, 'part_profile_link')
    profile_click.click()
    print('\nClicked Profile Button...')

    # EXPLICIT WAIT UNTIL CONTACT IS CLICKABLE
    try:
        element = WebDriverWait(driver, 10)
        element.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Contact')))
    except:
        print('Some Error Occurred... Logging Out')
        Log_Out()

    # ----------------- Go into Contact/Address ------------------------
    address_click = driver.find_element(By.PARTIAL_LINK_TEXT, 'Contact')
    address_click.click()
    print('Clicked Contact / Address Button...')
   
    # EXPLICIT WAIT UNTIL ADDRESS RESULTS TABLE IS VISIBLE (DISPLAYED ON SCREEN)
    try:
        element = WebDriverWait(driver, 10)
        element.until(EC.visibility_of_element_located((By.ID, 'part_profile_address_results')))
    except:
        print('Some Error Occurred... Logging Out')
        Log_Out()

    # --------------------------- Find Mailing Addresses ---------------------------------
    mail_addresses = driver.find_elements(By.PARTIAL_LINK_TEXT, 'Mailing') # TODO: Put in exception handling for if there are no addresses found
    mail_count = len(mail_addresses)
    print('\nCount for mailing addresses: ' + str(mail_count))
    new_address = False

# ------------------ FUNCTION FOR STALE ATTEMPTS ----------------------
    def Mail_Stale_Attempt(index, attempts, duration):
        attempt = 0
        for attempt in range(attempts): # https://stackoverflow.com/questions/2083987/how-to-retry-after-exception
            try:
                new_mail_addresses = driver.find_elements(By.PARTIAL_LINK_TEXT, 'Mailing')
                #print('\nWeb Addresses for i = ' + str(index) + ":")
                #print(new_mail_addresses)
                new_mail_addresses[index].click()

            except:
                #print('Sleeping for ' + str(duration) + ' seconds...')
                time.sleep(duration)

                # IF TOO MANY ATTEMPS, LOG OUT AND QUIT
                if attempt == attempts-1:
                    print('ERROR: TOO MANY ATTEMPTS...')
                    Log_Out()

                continue # Retry loop

            else:
                break # Get out of loop

        print('Mailing Address (i = ' + str(index) + ') SUCCESSFULLY clicked...')
        print('Number of additional attempts needed: ' + str(attempt))

# ------------------ FUNCTION FOR CLICKING INTO MAIL ADDRESS AND SAVING ----------------------
    def Mail_Click_And_Save(index):
        # TODO: ADD PROMPT WINDOW FOR CHECKING FOR INTERNET SPEED AND MAKE # OF ATTEMPTS BASED ON THAT (SLOW, MED, FAST)
        # Try attempts until not stale
        print('\nAttempting to Click Mailing Address (i = ' + str(index) + ')...')
        Mail_Stale_Attempt(index, 10, 0.2) # Parameters: (Index, # of attempts, # of seconds for each wait)

        # EXPLICIT WAIT UNTIL SAVE BUTTON IS CLICKABLE
        try:
            element = WebDriverWait(driver, 10)
            element.until(EC.element_to_be_clickable((By.CLASS_NAME, 'default')))
        except:
            print('Some Error Occurred... Logging Out')
            Log_Out()

        # Checking for any matches with excel and pairing with respective priority
        web_address = driver.find_element(By.ID, 'address_street').text # Look at street address textbox
        print('Current Address: ' + str(web_address))

        # if web_address == str(row['Address']): # If address MATCHES WITH NEW ADDRESS LINE
        #     select = Select(driver.find_element(By.NAME, 'priority'))
        #     print('Current Priority: ' + str(select.first_selected_option.text))
        #     print('Matching NEW Address Line...')
        #     global new_address
        #     new_address = True
            
        #     if str(select.first_selected_option.text) == 'Normal Priority':
        #         print('Address Line already correct priority...')
        #         driver.find_element(By.XPATH, '//button[normalize-space()="Cancel"]').click() # Presses the cancel button
        #         print('Cancel Button Pressed...')

        #     else:
        #         print('Making NORMAL Priority...')
        #         select.select_by_value('0') # (value = 0) Make NORMAL PRIORITY
        #         driver.find_element(By.CLASS_NAME, 'default').click() # Presses the save button
        #         print('Save Button Pressed...')

        # elif web_address == str(row['Original Address Line 1']): # If address MATCHES WITH ORIGINAL ADDRESS LINE
        #     select = Select(driver.find_element(By.NAME, 'priority')) # MAKE INACTIVE PRIORITY
        #     print('Current Priority: ' + str(select.first_selected_option.text))

        #     print('Matching ORIGINAL Address Line...')
        #     if str(select.first_selected_option.text) == 'Inactive':
        #         print('Address Line already correct priority...')
        #         driver.find_element(By.XPATH, '//button[normalize-space()="Cancel"]').click() # Presses the cancel button
        #         print('Cancel Button Pressed...')

        #     else:
        #         print('Making INACTIVE Priority...')
        #         select.select_by_value('-2') # (value = -2) Make INACTIVE PRIORITY
        #         driver.find_element(By.CLASS_NAME, 'default').click() # Presses the save button
        #         print('Save Button Pressed...')

        # else: # If address DOES NOT MATCH WITH ANYTHING
        #     select = Select(driver.find_element(By.NAME, 'priority')) # MAKE INACTIVE PRIORITY
        #     print('Current Priority: ' + str(select.first_selected_option.text))
        #     print('Unknown Address...')
        #     print('Making INACTIVE Priority...') 
        #     select.select_by_value('-2') # (value = -2) Make INACTIVE PRIORITY
        #     driver.find_element(By.CLASS_NAME, 'default').click() # Presses the save button

        # EXPLICIT WAIT UNTIL ADDRESS RESULTS TABLE IS VISIBLE (DISPLAYED ON SCREEN)
        try:
            element = WebDriverWait(driver, 10)
            element.until(EC.visibility_of_element_located((By.ID, 'part_profile_address_results')))
        except:
            print('Some Error Occurred... Logging Out')
            Log_Out()

# -------------- Looking through Addresses and Changing Priorities -------------------- 
    if mail_count > 1:
        for i in range(len(mail_addresses)): # https://www.codesansar.com/python-programming-examples/print-1-12-123-1234-pattern.htm
            for j in range(i+1):
                #print(str(j)) # Pattern now goes j = 0 -> 01 -> 012 (better for if values change places after a save)
                Mail_Click_And_Save(j) # Clicks into each mail
        
    elif mail_count == 1:
        Mail_Click_And_Save(0)

    else:
        print('Unknown Error: Going to next part of code')
