import os
import time
import pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class AutomationBot:

	def __init__(self):
		# Initiates all variables
		self._URL = "http://www.rpachallenge.com"
		self._LOG_PATH = rf'{os.path.dirname(__file__)}\Log'
		self._SCREENSHOT_PATH = rf'{os.path.dirname(__file__)}\Screenshots'
		self._DOWNLOAD_PATH = rf'{os.path.dirname(__file__)}\Downloaded Sheet'
		self._EXCEL_FILE = rf'{self._DOWNLOAD_PATH}\challenge.xlsx'
		self._EDGE_OPTIONS = webdriver.EdgeOptions()
		self._EDGE_OPTIONS.add_experimental_option('prefs', {
			'download.default_directory': self._DOWNLOAD_PATH,
			'download.prompt_for_download': False,
			'download.directory_upgrade': True,
			'safebrowsing.enabled': True
		})
		self._EDGE_OPTIONS.add_argument('-inprivate')
		self._EDGE_OPTIONS.add_argument('--start-maximized')
		self._driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=self._EDGE_OPTIONS, service_log_path=os.devnull)

	def access_site(self):
		# Access the URL
		self._driver.get(self._URL)
	
	def download_excel(self):
		# Delete a preveously downloaded file
		False if not os.path.isfile(self._EXCEL_FILE) else os.remove(self._EXCEL_FILE)
		
		# Initiates variables
		_BUTTON_XPATH = '//a[contains(text(), "Download Excel")]'
		
		# Awaits and Clicks the button
		download_button = WebDriverWait(self._driver, 10).until(ec.element_to_be_clickable((By.XPATH, _BUTTON_XPATH)))
		download_button.click()

		# Wait file exsists
		while not os.path.isfile(self._EXCEL_FILE):
			time.sleep(0.5)
			
	def process_document(self):
		# Create the images directory
		False if os.path.isdir(self._SCREENSHOT_PATH) else os.makedirs(self._SCREENSHOT_PATH)
		
		# Initiates variables
		_START_XPATH = '//button[text() = "Start"]'
		_SUBMIT_XPATH = '//input[@value = "Submit"]'
		
		# Reads the downloaded file
		excel = pandas.read_excel(io=self._EXCEL_FILE, dtype='str')
		
		# Start the field's filling
		self._driver.find_element(By.XPATH, _START_XPATH).click()
		
		# Iterates throght rows
		for row in excel.iterrows():
			# Iterates throght comlumns
			for column in excel.columns:
				# Create the XPath to be used
				temp_xpath = f'//label[contains(text(), "{column.strip()}")]//following-sibling::input'
				
				# Fill the fields
				self._driver.find_element(By.XPATH, temp_xpath).send_keys(row[1][column])
		
			# Make a page Screenshot
			self._driver.save_screenshot(rf'{self._SCREENSHOT_PATH}\Page {row[0]}.png')
			
			# Go to the next page
			self._driver.find_element(By.XPATH, _SUBMIT_XPATH).click()
			
			# Make a Screenshot os the final page
			self._driver.save_screenshot(rf'{self._SCREENSHOT_PATH}\Page Ending.png')
			
	def close_driver(self):
		# Closes both Window and Driver instances
		self._driver.close()
		self._driver.quit()
		
	def run(self):
		# Access the URL
		self.access_site()
		
		# Downloads the Excel sheet
		self.download_excel()
		
		# Fill fields
		self.process_document()

# Create the instance
automation_bot = AutomationBot()

# Bot process
try:
	# Starts the Automation
	automation_bot.run()
except Exception as err:
	# Throw error
	print(f'Script execution failure. Details: {str(err)}')

# Finishes
automation_bot.close_driver()
