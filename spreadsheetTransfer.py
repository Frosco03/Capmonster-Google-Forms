from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from capmonster_python import RecaptchaV2Task
from configparser import ConfigParser
import pandas as pd 
import time
import requests
import json

#essential global variables
file = 'config.ini'
config = ConfigParser()
config.read(file)
sitekeyElement = None
sitekey = None


def createFormElements():
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.whsOnd.zHQkBf')) #text
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.KHxj8b.tL9Q4c')) #textarea
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.jgvuAb.ybOdnf.cGN2le.t9kgXb.llrsB')) #dropdown
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.lLfZXe.fnxRtf.cNDBpf')) #radio group
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.lLfZXe.fnxRtf.BpKDyb')) #scale group
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.Y6Myld')) #list group
	formElements.extend(browser.find_elements(By.CSS_SELECTOR, '.e12QUd')) #grid group

def sortFormElements():
	#Basic bubble sort to sort the form elements
	n = len(formElements)

	#Traverse through form elements
	for i in range(n):
		for j in range(0, n-i-1):
			#Swap if the element's y location is greater than the next element
			if formElements[j].location.get('y') > formElements[j+1].location.get('y'): 
				formElements[j], formElements[j+1] = formElements[j+1], formElements[j]

def answerFormSingle(y):
	x = 0
	for individual in formElements:
		time.sleep(1)
		if(individual.get_attribute("class") == "whsOnd zHQkBf"): #present element is a text type
			individual.send_keys(excelData.at[y,headers[x]])

		elif(individual.get_attribute("class") == "KHxj8b tL9Q4c"): #present element is a textarea type
			individual.send_keys(excelData.at[y,headers[x]])

		elif(individual.get_attribute("class") == "lLfZXe fnxRtf cNDBpf"): #present element is a radio question
			radioButtons = individual.find_elements(By.CSS_SELECTOR, '.Od2TWd')
			for radio in radioButtons:
				if str(radio.get_attribute("data-value")).lower() == str(excelData.at[y,headers[x]]).lower():
					radio.click()
					break

		elif(individual.get_attribute("class") == "jgvuAb ybOdnf cGN2le t9kgXb llrsB"): #present element is a dropdown type
			individual.click()
			time.sleep(2)
			dropChoices = individual.find_elements(By.CSS_SELECTOR,'[role="option"]')
			temp = excelData.at[y,headers[x]].replace('.','') #Convert string with numbers into either float or int
			if(temp.isdecimal()):
				conv = float(excelData.at[y,headers[x]])
				if conv.is_integer():
					excelData.at[y,headers[x]] = int(conv)
			for choice in dropChoices:
				if str(choice.get_attribute("data-value")).lower() == str(excelData.at[y,headers[x]]).lower():
					choice.click()
					break

		elif(individual.get_attribute("class") == "lLfZXe fnxRtf BpKDyb"): #present element is a scale group
			radioButtons = individual.find_elements(By.CSS_SELECTOR, '[role="radio"]')
			for radio in radioButtons:
				if str(radio.get_attribute("data-value")).lower() == str(excelData.at[y,headers[x]]).lower():
					radio.click()
					break

		elif(individual.get_attribute("class") == "Y6Myld"): #present element is a list group
			checkButtons = individual.find_elements(By.CSS_SELECTOR, '[role="checkbox"]')
			for box in checkButtons:
				if str(box.get_attribute("data-answer-value")).lower() in str(excelData.at[y,headers[x]]).lower():
					box.click()

		elif(individual.get_attribute("class") == "e12QUd"): #present element is a grid group
			radioGroupRow = individual.find_elements(By.CSS_SELECTOR, '[role="radiogroup"]')
			for radioGroup in radioGroupRow:
				radioButtons = radioGroup.find_elements(By.CSS_SELECTOR, '[role="radio"]')
				for radio in radioButtons:
					if str(radio.get_attribute("data-value")).lower() in str(excelData.at[y,headers[x]]).lower():
						radio.click()
		x+=1

def answerFormMultiple(y):
	global sitekeyElement 
	global sitekey 
	x = 0
	pageCount = int(config['web-info']['page-count'])
	for pageNo in range(1, pageCount+1):
		if(pageNo > 1):
			#Reload form elemets and organize them
			formElements.clear()
			createFormElements()
			sortFormElements()
		index = 1
		for individual in formElements:
			time.sleep(1)
			if((len(formElements) > 1) and (index < len(formElements))):
				index+=1
				continue #this code makes us continue from where we left off from the last formElements item
			if(individual.get_attribute("class") == "whsOnd zHQkBf"): #present element is a text type
				individual.send_keys(excelData.at[y,headers[x]])

			elif(individual.get_attribute("class") == "KHxj8b tL9Q4c"): #present element is a textarea type
				individual.send_keys(excelData.at[y,headers[x]])

			elif(individual.get_attribute("class") == "lLfZXe fnxRtf cNDBpf"): #present element is a radio question
				radioButtons = individual.find_elements(By.CSS_SELECTOR, '.Od2TWd')
				for radio in radioButtons:
					if radio.get_attribute("data-value").lower() == excelData.at[y,headers[x]].lower():
						radio.click()
						break

			elif(individual.get_attribute("class") == "jgvuAb ybOdnf cGN2le t9kgXb llrsB"): #present element is a dropdown type
				individual.click()
				time.sleep(2)
				dropChoices = individual.find_elements(By.CSS_SELECTOR,'[role="option"]')
				temp = excelData.at[y,headers[x]].replace('.','')
				if(temp.isdecimal()): #Convert string with numbers into either float or int
					conv = float(excelData.at[y,headers[x]])
					if conv.is_integer():
						excelData.at[y,headers[x]] = int(conv)
				for choice in dropChoices:
					if str(choice.get_attribute("data-value")).lower() == str(excelData.at[y,headers[x]]).lower():
						choice.click()
						break

			elif(individual.get_attribute("class") == "lLfZXe fnxRtf BpKDyb"): #present element is a scale group
				radioButtons = individual.find_elements(By.CSS_SELECTOR, '[role="radio"]')
				for radio in radioButtons:
					if str(radio.get_attribute("data-value")).lower() == str(excelData.at[y,headers[x]]).lower():
						radio.click()
						break

			elif(individual.get_attribute("class") == "Y6Myld"): #present element is a list group
				checkButtons = individual.find_elements(By.CSS_SELECTOR, '[role="checkbox"]')
				for box in checkButtons:
					if str(box.get_attribute("data-answer-value")).lower() in str(excelData.at[y,headers[x]]).lower():
						box.click()

			elif(individual.get_attribute("class") == "e12QUd"): #present element is a grid group
				radioGroupRow = individual.find_elements(By.CSS_SELECTOR, '[role="radiogroup"]')
				for radioGroup in radioGroupRow:
					radioButtons = radioGroup.find_elements(By.CSS_SELECTOR, '[role="radio"]')
					for radio in radioButtons:
						if str(radio.get_attribute("data-value")).lower() in str(excelData.at[y,headers[x]]).lower():
							radio.click()
			x+=1
		time.sleep(1)
		if(pageNo==1):
			browser.find_element(By.CSS_SELECTOR,'.l4V7wb.Fxmcue').click() #click the next button
		elif(pageNo < pageCount): #succeeeding pages except last page
			nextBtn = browser.find_elements(By.CSS_SELECTOR,'.l4V7wb.Fxmcue') #returns 3 web elements: back btn, next btn, clear form btn
			nextBtn[1].click() #click the next button
		else: #last page won't click a next button anymore but get the captcha key
			sitekeyElement = browser.find_element(By.ID, "recaptcha")
			sitekey = sitekeyElement.get_attribute("data-sitekey")
		time.sleep(2)

def solveCaptcha(y):
	print("Solving Captcha #" + str(y) + "...")
	capmonster = RecaptchaV2Task(config['capmonster']['capmonsterKey'])
	task_id = capmonster.create_task(url, sitekey)
	result = capmonster.join_task_result(task_id)
	print("Captcha Response:")
	print(result.get("gRecaptchaResponse"))

	browser.execute_script('var element=document.getElementById("g-recaptcha-response"); element.style.display="";')
	browser.execute_script("""document.getElementById("g-recaptcha-response").textContent = arguments[0]""", result.get("gRecaptchaResponse"))
	browser.execute_script('var element=document.getElementById("g-recaptcha-response"); element.style.display="none";')

"""
Basic Algorithm:
1. Read excel file and store in a DataFrame variable
2. Create headers variable from the dataframes
3. Initialize webdriver, load the website, and get the captcha sitekey
4. Initialize an empty list of form elements
5. Find and append all input form elements to the list
6. Organize all input form elements to match their corresponding places
7. Fill up the form based on the values from the DataFrame
8. If the next value is empty, stop the program
"""

url = config['web-info']['url']

#1. Read excel file and store in a DataFrame variable
if(config['spreadsheet']['index_id'] == 'True'):
	excelData = pd.read_excel(config['spreadsheet']['spreadsheet-dir'], index_col=0)
else:
	excelData = pd.read_excel(config['spreadsheet']['spreadsheet-dir'], index_col=None)
	excelData.index+=1
	
excelData = excelData.astype(str) #turn the whole dataframe into strings
startIndex = 1

#2. Create headers variable from the dataframes
headers = excelData.columns

#3. Initialize webdriver, load the website, and get the captcha sitekey
#browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
browser = webdriver.Chrome()

browser.get(url)

if(config['web-info']['page-count'] == '1'):
	sitekeyElement = browser.find_element(By.ID, "recaptcha")
	sitekey = sitekeyElement.get_attribute("data-sitekey")

dataCount = excelData.shape[0]

exceptCtr = startIndex #counter to count at what datatype an error occured.
try:
	for y in range(startIndex, dataCount):

		#4. Initialize an empty list of form elements
		formElements = []

		#5. Find and append all input form elements to the list
		createFormElements()

		#6. Organize all input form elements to match their corresponding places
		sortFormElements()

		#7. Fill up the form based on the values from the DataFrame
		if(int(config['web-info']['page-count']) > 1):
			answerFormMultiple(y)
			formButtons = browser.find_elements(By.CSS_SELECTOR,'.l4V7wb.Fxmcue')
			submit = formButtons[1] #submit button 
		else:
			answerFormSingle(y)
			submit = browser.find_element(By.CSS_SELECTOR,'.l4V7wb.Fxmcue') #submit button

		time.sleep(2)
		solveCaptcha(y) #solve captcha before submitting
		browser.execute_script("window.scrollTo(0, document.body.scrollHeight)") #scroll to the bottom of the page
		browser.execute_script("arguments[0].click();", submit)
		#submit.click() #submit the form
		print("Successfully encoded item number " + str(exceptCtr))
		print("\n===================================================\n")
		exceptCtr+=1
		time.sleep(1)
		browser.get(url) #submit more responses
		time.sleep(2)

except Exception as e:
	print("There has been an error while encoding item number " + str(exceptCtr) + ".\n\n")
	print("Error log for developers:\n")
	print(e)

else:
	print("Finished Encoding!")

finally:
	print("The browser will close in 15 seconds.")
	time.sleep(15)
	browser.quit()