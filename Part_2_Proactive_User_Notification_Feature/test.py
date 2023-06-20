import os
from selenium import webdriver

# chrome_driver_path = "C:/Users/nanda/Desktop/chromedriver.exe"
# webdriver.ChromeOptions().binary_location = chrome_driver_path
# driver = webdriver.Chrome(chrome_driver_path)

driver = webdriver.Edge()

driver.get("https://www.google.com")

driver.quit()