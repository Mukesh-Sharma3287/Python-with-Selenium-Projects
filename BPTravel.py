from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime

class BPTravel:
    def __init__(self, driver_path: str, url: str):
        self.driver_path = driver_path
        self.url = url
        #self.chrome_options = Options()
        #self.chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.element_present = None
        self.locator = None
        self.timeout = 30
        self.driver = webdriver.Edge(executable_path=driver_path)
        self.driver.get(self.url)
        self.driver.maximize_window()
        self.delay = 0.005

    def login_bp_travel(self):
        try:
            # User name
            self.locator = "#username"
            self.element_present = EC.presence_of_element_located((By.CSS_SELECTOR, self.locator))
            WebDriverWait(self.driver, self.timeout).until(self.element_present)
            self.driver.find_element_by_css_selector(self.locator).send_keys("admin")
            time.sleep(self.delay)

            # Enter Password
            self.locator = "#password"
            self.element_present = EC.presence_of_element_located((By.CSS_SELECTOR, self.locator))
            WebDriverWait(self.driver, self.timeout).until(self.element_present)
            self.driver.find_element_by_css_selector(self.locator).send_keys("admin")
            time.sleep(self.delay)

            # Login
            self.locator = "//span[@onclick='login()']"
            self.element_present = EC.presence_of_element_located((By.XPATH, self.locator))
            WebDriverWait(self.driver, self.timeout).until(self.element_present)
            self.driver.find_element_by_xpath(self.locator).click()
            time.sleep(self.delay)

        except Exception as e:
            print("Unable to Login BP Travel"+str(e))
            raise Exception()

    def create_quotes(self, exl_path):
        try:
            df_quotes = pd.read_excel(exl_path)
            for i in range(df_quotes.shape[0]):

                # Click on Create Quote
                locator = "//a[@href='createquote.html']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).click()
                time.sleep(self.delay)

                # From
                locator = "//select[@id='from']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 7])
                time.sleep(self.delay)

                # To
                locator = "//select[@id='to']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 8])
                time.sleep(self.delay)

                # Departing
                locator = "//input[@id='departing']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 9].strftime('%d/%m/%Y'))
                time.sleep(self.delay)

                if pd.isna(df_quotes.iloc[i, 10]):
                    # One way
                    locator = "//input[@id='oneway']"
                    self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                    WebDriverWait(self.driver, self.timeout).until(self.element_present)
                    self.driver.find_element_by_xpath(locator).click()
                    time.sleep(self.delay)
                else:
                    # Return
                    locator = "//input[@id='returning']"
                    self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                    WebDriverWait(self.driver, self.timeout).until(self.element_present)
                    self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 10].strftime('%d/%m/%Y'))
                    time.sleep(self.delay)

                # Adult
                locator = "//select[@id='adults']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                # select = Select(self.driver.find_element_by_xpath(locator))
                # select.select_by_visible_text(int(df_quotes.iloc[i, 11]))
                self.driver.find_element_by_xpath(locator).send_keys(int(df_quotes.iloc[i, 11]))
                time.sleep(self.delay)

                # Children
                locator = "//select[@id='children']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                # select = Select(self.driver.find_element_by_xpath(locator))
                # select.select_by_visible_text(int(df_quotes.iloc[i, 12]))
                self.driver.find_element_by_xpath(locator).send_keys(int(df_quotes.iloc[i, 12]))
                time.sleep(self.delay)

                # Name
                locator = "//input[@id='name']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 1]+' '+df_quotes.iloc[i, 2]+' '+df_quotes.iloc[i, 3])
                time.sleep(self.delay)

                # Postcode
                locator = "//input[@id='postcode']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 6])
                time.sleep(self.delay)

                # Telephone
                locator = "//input[@id='telephone']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(int(df_quotes.iloc[i, 5]))
                time.sleep(self.delay)

                # Email
                locator = "//input[@id='email']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).send_keys(df_quotes.iloc[i, 4])
                time.sleep(self.delay)

                # Create Quote
                locator = "//span[@class='button' and @onclick='createQuote()']"
                self.element_present = EC.presence_of_element_located((By.XPATH, locator))
                WebDriverWait(self.driver, self.timeout).until(self.element_present)
                self.driver.find_element_by_xpath(locator).click()
                time.sleep(self.delay)

        except Exception as ex:
            raise ex

    def check_exists_by_xpath(self, xpath):
        try:
            self.driver.find_element_by_xpath(xpath)
        except NoSuchElementException:
            return False
        return True



