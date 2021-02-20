# This script creates screenshots from a list of URLs in an excel file.
# The Excel file should have a header and at least 2 columns,
# one with the URLs, and another a yes/no column to track if we want
# a screenshot for that row, i.e.
# nameOfExcelColumnToDecideIfScreenshotIsTaken = 'Use'
# Value in that Column if No Screenshot should be taken.
# valueOfExcelColumnIfScreenshotIsNotToBeTaken = 'No'
# These variables can be set below.

# This script uses Selenium and the Chrome driver.
# Modified from https://pythonbasics.org/selenium-screenshot/

# The Selenium module must be installed.
# pip3 install selenium

# For the Mac, download the Chrome driver from
# https://chromedriver.storage.googleapis.com/index.html?path=88.0.4324.96/
# as specified in:
# https://sites.google.com/a/chromium.org/chromedriver/downloads
# Download the driver, e.g. the chromedriver, in the same directory as this python script.

from selenium import webdriver
# To handle cookies.
from selenium.webdriver.chrome.options import Options
from time import sleep

import sys
import pandas as pd

# From: https://stackoverflow.com/questions/287871/how-to-print-colored-text-to$
class bcolors:
  HEADER = '\033[95m'
  OKBLUE = '\033[94m'
  OKCYAN = '\033[96m'
  OKGREEN = '\033[92m'
  WARNING = '\033[93m'
  FAIL = '\033[91m'
  ENDC = '\033[0m'
  BOLD = '\033[1m'
  UNDERLINE = '\033[4m'

screenshotsDirectory = '/Users/jesusdelvalle/Documents/projects/make_slides/screenshots/'
pathToChromedriver = '/Users/jesusdelvalle/Documents/projects/make_slides/chromedriver'
chromeOptionsCookieFolder = "user-data-dir=selenium"
nameOfExcelColumnWithURLs = 'URL'
# There should be a yes/no column to track if we want a screenshot for that row.
nameOfExcelColumnToDecideIfScreenshotIsTaken = 'Use'
# Value in that Column if No Screenshot should be taken.
valueOfExcelColumnIfScreenshotIsNotToBeTaken = 'No'

try:
  # Expected: Excel file with ending .xslx
  inputFilename = sys.argv[1]

  excel = pd.read_excel(inputFilename, header=0)

  # To handle cookies.
  # From: https://stackoverflow.com/questions/15058462/how-to-save-and-load-cookies-using-python-selenium-webdriver
  chrome_options = Options()
  chrome_options.add_argument(chromeOptionsCookieFolder)

  driver = webdriver.Chrome(pathToChromedriver, options=chrome_options)

  i = 0
  for index, row in excel.iterrows():

    use = row[nameOfExcelColumnToDecideIfScreenshotIsTaken]
    if use == valueOfExcelColumnIfScreenshotIsNotToBeTaken:
      continue

    url = row[nameOfExcelColumnWithURLs]
    print(row[nameOfExcelColumnWithURLs])

    driver.get(url)
    # The script should be run twice.
    # The first time you run the script, you give sleep time so the webpage
    # loads and you can click on the "Accept cookies" buttons.
    # Then run the script a second time, the pages will take the previously
    # saved cookies and you don't see the "Accept Cookies" messages anymore.
    # These Screenhots are saved and will replace the Screenshots from the
    # first time you run the script.
    sleep(5)

    url = url.replace("https://","")
    url = url.replace("/","")
    # print (url)
    savedFile = screenshotsDirectory + url + ".png"
    # print (savedFile)
    driver.get_screenshot_as_file(savedFile)

    i = i + 1
  driver.quit()
  print("Number of websites:", i)
except IOError:
  print(f"{bcolors.FAIL}File not found: {bcolors.ENDC}" + inputFilename)
except IndexError as error:
  print(f"{bcolors.FAIL}Please enter a filename:{bcolors.ENDC}", error)
except Exception as error:
  print(f"{bcolors.FAIL}Error:{bcolors.ENDC}", error)
