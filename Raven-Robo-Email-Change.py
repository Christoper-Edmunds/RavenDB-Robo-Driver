#Raven-Robo-Email-Change

#Requires basic set up by opening the terminal, assuming using VSCode with Python 3.10.7 
#CTRL+SHIFT+P to open the command pallete, type terminal then ENTER to open the terminal, then enter the command "pip install selenium", further documentation below.
#https://www.selenium.dev/documentation/webdriver/getting_started/install_library/#requirements-by-language
#now also using Pyperclip, https://pypi.org/project/pyperclip/ 


#fringe bug, any zoom settings on the page will break the bot, need a script to reset zoom

""" 
Design Doc

Basic Functionality
Robo-Drive a ravenDB frontend to automate data edit tickets using the following technologies 
Python
Selenium 

Extended Functionality
Connect to Slack API to allow for non-technical users

Initial Attempt: Changing an email automatically 

"""
 
#depreciated or non-functioning methods 
######################################    
#EditJSONUsingJscriptInjection(driver) #depreciated cus not as good as the chosen method, but might be useful for fringe issues + other functionality like batch edits
#GetPlacementCode(driver) #non-functioning as of yet
#test_eight_components() #exists purely to test that the library is functioning properly, will be removed later in development
#NavigateToRavenDBPlacement(driver) #Origional proof of concept, now being broken up into modules for dynamic behaviour 

#The below functions are on the todo list
#ReplaceContractEmailCandidate(driver):
#ReplacePersonalAddressCandidate(driver):
#ReplaceLTDAddressCandidate(driver):
#ReplacePersonalEmailCandidate(driver):
######################################
 
 
#initial Set up and navigation
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import time
import json
import pprint
import tkinter as tk
import pyperclip


#String for chrome profile data, should be changed for new users
# C:\Users\ChrisEdmunds\AppData\Local\Google\Chrome\User Data\Default
#unique path can be found by entering chrome://version/ into your browser, it will be displayed under "profile path:"

#Sets up the Robo-driver with the correct cookies and capabilities. 
options = webdriver.ChromeOptions() 
options.add_argument(r"user-data-dir=C:\Users\ChrisEdmunds\AppData\Local\Google\Chrome\User Data") #Path to your chrome profile
driver = webdriver.Chrome(service=ChromeService(executable_path=ChromeDriverManager().install()), chrome_options=options)


#Test Function to insure robo-driver library is working
def test_eight_components():
    #test doc to check the library is functioning correctly
    driver = webdriver.Chrome(service=ChromeService(executable_path=ChromeDriverManager().install()))

    driver.get("https://www.selenium.dev/selenium/web/web-form.html")

    title = driver.title
    assert title == "Web form"

    driver.implicitly_wait(0.5)

    text_box = driver.find_element(by=By.NAME, value="my-text")
    submit_button = driver.find_element(by=By.CSS_SELECTOR, value="button")

    text_box.send_keys("Selenium")
    submit_button.click()

    message = driver.find_element(by=By.ID, value="message")
    value = message.text
    assert value == "Received!"
    
    driver.implicitly_wait(10.0)

    driver.quit()
#OG proof of concept, currently depreciated    
def NavigateToRavenDBPlacement(driver):
    #loads the ravenDB website
    driver.get("https://sonovate3-ygod.ravenhq.com/studio/index.html#databases/query/index/AdjustedTransactionIndex?&database=Invoicing-Production")
    

    #checks that the html title is correct for the loading screen for ravendb
    title = driver.title
    assert title == "Raven.Studio"

    #forces the robo to wait until the title has changed to indicate that the page has loaded, has a timeout of 20 seconds
    WebDriverWait(driver, timeout=20).until(EC.title_is("Query | Raven.Studio"))
    
    #finds the box to input the placement data
    text_box = driver.find_element(By.ID, value="goToDocInput")
    
    #sends the placement data as key presses, has waits built in because sometimes the page reacts slowly, waits akin to those used for loading the page would be more resiliant
    placement_string = input("Enter Placement String: \n")
    text_box.send_keys(placement_string)
    #text_box.send_keys("placements/161257")
    time.sleep(1)
    text_box.click
    time.sleep(1)
    text_box.send_keys(Keys.ENTER)
    time.sleep(1)
    
    #if the page doesn't change, then nothing was found to enter, if it was, begins navigating the json 
    if title == "Query | Raven.Studio":
        print("No such document found")
    else:
        print("Document Found, Navigating")

        time.sleep(1)
        
        #clicks into the edit box
        ace_box = driver.find_element(By.CLASS_NAME, value="ace_content")
        ace_box.click
        

        #Forces the edit box into fullscreen and then back out again by entering shift + f11, as this loads the entire JSON file into the HTML at once
        #This makes it easier to search programatically, as you dont have to send key inputs to scroll and read lines and can instead edit the DIV's 
        ActionChains(driver).key_down(Keys.LEFT_SHIFT).perform()
        ActionChains(driver).send_keys(Keys.F11).perform()
        ActionChains(driver).key_up(Keys.LEFT_SHIFT).perform()
        ActionChains(driver).send_keys(Keys.F11).perform()  
        
        ActionChains(driver).key_down(Keys.LEFT_SHIFT).perform()
        ActionChains(driver).send_keys(Keys.F11).perform()
        ActionChains(driver).key_up(Keys.LEFT_SHIFT).perform()
        ActionChains(driver).send_keys(Keys.F11).perform()
        
        
        #Captures the contents of the ace box:  captures it from the clipboard, saves it to a variable, recognises it as a json file and converts it into a dictionary of objects,
        #This dictionary can be iterated through much more programatically, the format can then be converted back into a json format
        #The output can then be pasted back and the formatting repaired by the robo driver 
        
        #Select all (ctrl + a)
        ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
        ActionChains(driver).send_keys("a").perform()
        ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()

        #Copy (ctrl + c)
        ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
        ActionChains(driver).send_keys("c").perform()
        ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()

        #Grabs the data from the clipboard and puts it into a variable
        root = tk.Tk()
        json_as_text = root.clipboard_get()
        
        #converts the string into a python dictionary that the json library can recognise as json objects
        json_as_dict = json.loads(json_as_text)
        
        #converts back to a string as proof of ability
        json_as_formatted_string = json.dumps(json_as_dict)
        print(json_as_formatted_string)
        
        
        #Edits the value under the ContactDetails > TradingAddress > Line1 object to = Party House used for specific changes
        json_as_dict['ContactDetails']['TradingAddress']['Line1'] = "Party House"
        json_as_dict['ContactDetails']['TradingAddress']['Line2'] = "Where The Fun Happens"
        json_as_dict['ContactDetails']['TradingAddress']['Line3'] = "But the Invoices go to the office still"


        #example for replacing all instances of substring, may not work, idk we will see 
        
        #list = ['helloXXX', 'welcomeXXX', 'to999', 'SofthuntUUU']
        #replace = [list.replace('XXX', 'ZZZ') for list in list]
        #print(replace)
        
        
        #Converts this edited json back into a string
        json_as_formatted_string = json.dumps(json_as_dict)
        
        
        #clears the clipboard and adds the edited json to it 
        root.clipboard_clear()
        pyperclip.copy(json_as_formatted_string)
        
        #pastes the edited string into the ace box
        ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
        ActionChains(driver).send_keys("v").perform()
        ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()
#Origional method, a better method was devised so this is no longer used        
def EditJSONUsingJscriptInjection(driver):
            #Creates a tuple containing every ace line as a seperate list object
        div_contents = driver.find_elements(By.CLASS_NAME, value="ace_line")


        #prints the objects
        print(div_contents)
        
        
        #prints out a copy of the json file. May store this as an array to search programatically for understanding logically
        for line in div_contents:
            
            print(line.text)
            #json_as_text = json_as_text + line.text
        
        #json_as_text = json_as_text + ","    
        #json_as_dict = json.loads(json_as_text)
        #json_as_formatted_string = json.dumps(json_as_dict)
        #print(json_as_formatted_string)
        
        string_to_find = ""
        string_to_replace = ""
        
        string_to_find = input("What string are you looking to change?")
        string_to_replace = input("what would you like to replace it with?")
        
        #Constructs a JS line to replace documents on the page
        JS_String_Builder = ""
        replace_all_string = 'document.body.innerHTML = document.body.innerHTML.replaceAll("'
        JS_String_Builder = replace_all_string + string_to_find + '", ' + '"' + string_to_replace + '" )'''
                
        #Runs a line of javascript to make bulk edits to the entire HTML of the page, this allows for the HTML to be altered on the page, allowing for pretty significant json changes.
        #These strings can be constructed programatically and entered as a variable to allow for very dynamic behaviour
        #driver.execute_script('document.body.innerHTML = document.body.innerHTML.replaceAll("Line1", "Testing?")')
        driver.execute_script(JS_String_Builder)


#Gets the placement URL from the placement dashboard, using the placement number from the ticket - nonfunctioning
def GetPlacementCode(driver):
    #Im struggling to get this to work, for some reason the buttons aren't registering clicks. Not sure why. 
    driver.get("https://members.sonovate.com/") 
    time.sleep(2)
    
    initialTitle = driver.title 
    if "Sign In" in initialTitle:
        Login_Button = driver.find_element(By.ID, value="login")
        Login_Button.click
    else:
        pageURl = driver.current_url
        assert pageURl == "https://members.sonovate.com/"
        
    pageURl = driver.current_url
    assert pageURl == "https://members.sonovate.com/"
    
    time.sleep(1)
    Placements_Button = driver.find_element(By.XPATH, "//*[@data-testid='Placements']")
    WebDriverWait(driver, timeout=20).until(EC.element_to_be_clickable((By.XPATH, "//*[@data-testid='Placements']")))
    time.sleep(1)
    Placements_Button.click
#Opens and preps a candidate file using a candidate file URL provided by the user
def NavigateToRavenDBCandidate(driver):
    
    #loads the ravenDB website
    candidate_url = input("input the candidate URL")
    driver.get(candidate_url)
    

    #checks that the html title is correct for the loading screen for ravendb
    title = driver.title
    assert title == "Raven.Studio"

    #forces the robo to wait until the title has changed to indicate that the page has loaded, has a timeout of 20 seconds
    #WebDriverWait(driver, timeout=20).until(EC.title_is("Query | Raven.Studio"))
    
    
    #if the page doesn't change, then nothing was found to enter, if it was, begins navigating the json 
    if title == "Query | Raven.Studio":
        print("No such document found")
    else:
        print("Document Found, Navigating")

        time.sleep(5)
        
        #clicks into the edit box
        ace_box = driver.find_element(By.CLASS_NAME, value="ace_content")
        ace_box.click
#Extracts the contents of the Ace Content Editor Box open on the current accessed page        
def ExtractAceContents(driver):
    
    #Forces the edit box into fullscreen and then back out again by entering shift + f11, as this loads the entire JSON file into the HTML at once
    #This makes it easier to search programatically, as you dont have to send key inputs to scroll 
    ActionChains(driver).key_down(Keys.LEFT_SHIFT).perform()
    ActionChains(driver).send_keys(Keys.F11).perform()
    ActionChains(driver).key_up(Keys.LEFT_SHIFT).perform()
    ActionChains(driver).send_keys(Keys.F11).perform()  
    
    ActionChains(driver).key_down(Keys.LEFT_SHIFT).perform()
    ActionChains(driver).send_keys(Keys.F11).perform()
    ActionChains(driver).key_up(Keys.LEFT_SHIFT).perform()
    ActionChains(driver).send_keys(Keys.F11).perform()
    
    
    #Captures the contents of the ace box:  captures it from the clipboard, saves it to a variable, recognises it as a json file and converts it into a dictionary of objects,
    #This dictionary can be iterated through much more programatically, the format can then be converted back into a json format
    #The output can then be pasted back and the formatting repaired by the robo driver 
    
    #Select all (ctrl + a)
    ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
    ActionChains(driver).send_keys("a").perform()
    ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()

    #Copy (ctrl + c)
    ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
    ActionChains(driver).send_keys("c").perform()
    ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()

    #Grabs the data from the clipboard and puts it into a variable
    root = tk.Tk()
    json_as_text = root.clipboard_get()
    
    #converts the string into a python dictionary that the json library can recognise as json objects
    json_as_dict = json.loads(json_as_text)
    
    #converts back to a string as proof of ability
    json_as_formatted_string = json.dumps(json_as_dict)
    print(json_as_formatted_string)
    
    #clears the clipboard
    root.clipboard_clear()
    
    return json_as_dict
#Replaces the contents of the Ace Content Editor Box with whatever json string you feed it (json_as_formatted_string)
def InsertAceContents(driver, json_as_formatted_string):
    
    #adds the edited json to the clipboard
    
    clipboard_still_held = True
    exception_timout = 0
    while clipboard_still_held == True and exception_timout <= 30 :
        try:
            pyperclip.copy(json_as_formatted_string)
        except:
            print("clipboard access denied, retrying (this is normal)")
            time.sleep(1)
            timout += 1
            clipboard_still_held = True
        else:
            print("Clipboard Opened Again! (i hate tkinter)")
            clipboard_still_held = False
    
    #pastes the edited string into the ace box
    ActionChains(driver).key_down(Keys.LEFT_CONTROL).perform()
    ActionChains(driver).send_keys("v").perform()
    ActionChains(driver).key_up(Keys.LEFT_CONTROL).perform()
#Replaces invoice candidates email, returns json_as_formatted_string
def ReplaceInvoiceEmailCandidate(driver, json_as_dict):
    json_as_dict['Line1'] = "Party House"
    
    #Converts this edited json back into a string
    json_as_formatted_string = json.dumps(json_as_dict)
    
    return json_as_formatted_string
#Replaces contract candidates email, returns json_as_formatted_string
def ReplaceContractEmailCandidate(driver, json_as_dict):
    #asks the user to input the email they want to change it to 
    Email_For_Contracts = input("Enter the email you want to set for the candidate contract handler")
    #sets the email to what was input
    json_as_dict['LimitedCompany']['EmailForContracts'] = Email_For_Contracts
    
    #Converts this edited json back into a string
    json_as_formatted_string = json.dumps(json_as_dict)
    
    return json_as_formatted_string
#Replaces personal invoice candidates address, returns json_as_formatted_string
def ReplacePersonalAddressCandidate(driver, json_as_dict):
    pass
#Replaces limited company invoice address, returns json_as_formatted_string
def ReplaceLTDAddressCandidate(driver, json_as_dict):
    #asks the user to input the email they want to change it to 
    Line1 = input("Enter Line 1: ")
    json_as_dict['LimitedCompany']['Address']['Line1'] = Line1
    
    Line2 = input("Enter Line 2: ")
    json_as_dict['LimitedCompany']['Address']['Line2'] = Line2
    
    Line3 = input("Enter Line 3: ")
    json_as_dict['LimitedCompany']['Address']['Line3'] = Line3
    
    City = input("Enter City: ")
    json_as_dict['LimitedCompany']['Address']['City'] = City
    
    Postcode = input("Enter Postcode: ")
    json_as_dict['LimitedCompany']['Address']['Postcode'] = Postcode
    
    Country = input("Enter Country: ")
    json_as_dict['LimitedCompany']['Address']['Country'] = Country
    
    UniqueId = input("Enter Unique ID: ")
    json_as_dict['LimitedCompany']['Address']['UniqueId'] = UniqueId
    
    IsRegisteredAddress = input("Enter is registered status: ")
    json_as_dict['LimitedCompany']['Address']['IsRegisteredAddress'] = IsRegisteredAddress
    
    AsString = Line1 + " ," + Line2 + " ," + Line3 + " ," + City + " ," + Postcode + " ," +  Country #creates final string from inputs
    AsString_Fixed = AsString.replace("null", " ") #removes null values
    AsString_Fixed = AsString_Fixed.replace(" , ,", " ,") #removes nonvalued parts
    json_as_dict['LimitedCompany']['Address']['AsString'] = AsString_Fixed    
    
    AsStringIncludingNonvaluedParts = Line1 + " ," + Line2 + " ," + Line3 + " ," + City + " ," + Postcode + " ," +  Country #creates final string from inputs
    AsStringIncludingNonvaluedParts = AsStringIncludingNonvaluedParts.replace("null", " ") #removes null values
    json_as_dict['LimitedCompany']['Address']['AsStringIncludingNonvaluedParts'] = AsStringIncludingNonvaluedParts  
    #Converts this edited json back into a string
    json_as_formatted_string = json.dumps(json_as_dict)
    
    return json_as_formatted_string
#Replaces a personal candidates email, returns json_as_formatted_string
def ReplacePersonalEmailCandidate(driver, json_as_dict):
    pass    
#Presses the auto format button shortcut (alt + [)
def AutoFormatButton(driver):
    
    ActionChains(driver).key_down(Keys.LEFT_ALT).perform()
    ActionChains(driver).send_keys("[").perform()
    ActionChains(driver).key_up(Keys.LEFT_ALT).perform()
#Presses the Save button shortcut (alt + s)
def SaveButton(driver):
    
    ActionChains(driver).key_down(Keys.LEFT_ALT).perform()
    ActionChains(driver).send_keys("s").perform()
    ActionChains(driver).key_up(Keys.LEFT_ALT).perform()


#example of replacing invoice candidate https://invoicing-staging.sonovate.com:13518/studio/index.html#databases/edit?&id=candidates%2F999001153&database=billing&list=Candidates&item=0
NavigateToRavenDBCandidate(driver)
json_as_dict = ExtractAceContents(driver)
json_as_formatted_string = ReplaceContractEmailCandidate(driver, json_as_dict)
InsertAceContents(driver, json_as_formatted_string)
AutoFormatButton(driver)

#NavigateToRavenDBPlacement(driver)


#Shuts down the robo-driver. THIS IS VERY IMPORTANT, either kill the program manually or make sure this always runs, strange behaviour can happen if the robo is left running 
isInputCorrect = input("Please Check JSON to insure validity of edits, enter y if it is, and n if it is not, please dont enter bunk data there is no validation :P   ")
if isInputCorrect == "y":
    #this is where you'd press the save button
    SaveButton(driver)
    driver.quit()
elif isInputCorrect == "n":
    driver.quit()
else:
    driver.quit()