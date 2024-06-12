import time

import selenium
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver

from openpyxl import Workbook
import pandas as pd

from env import URL, DriverLocation

"""Scraped Daten"""
def get_data(driver):
    print('Data scraping...')
    more_elements = driver.find_elements_by_class_name('w8nwRe kyuRq')
    for list_more_element in more_elements:
        list_more_element.click()
        print("click auf das Mehr")
    
    elements = driver.find_elements_by_class_name(
        'jftiEf')
    list_data = []
    x = 0
    for data in elements:
        name = data.find_element_by_xpath(
            './/div[2]/div[2]/div[1]/button/div[1]').text

        if check_exists_by_xpath('.//div[2]/div[2]/div[1]/button/div[2]', data):
            leftover = data.find_element_by_xpath(
                './/div[2]/div[2]/div[1]/button/div[2]').text
        else: leftover = ""

        if check_exists_by_xpath('.//div[@class="MyEned"]/span[1]', data):
            kommentar = data.find_element_by_xpath(
                './/div[@class="MyEned"]/span[1]').text
        else: kommentar = ""

        rating = data.find_element_by_xpath(
            './/div/div/div[4]/div[1]/span[1]').get_attribute("aria-label")
        ago = data.find_element_by_xpath(
            './/div/div/div[4]/div[1]/span[2]').text


        x = x + 1
        print(str(x) + " scrape :)")

        list_data.append([name, leftover, ago, rating, kommentar])

    return list_data

"""ermitteln der Iterationen anhand der Rezensionsanzahl"""
def counter():
    result = driver.find_element_by_class_name('jANrlb').find_element_by_class_name('fontBodySmall').text
    result = result.replace('.', '')
    result = result.split(' ')
    result = result[0].split('\n')
    return int(int(result[0])/10)+1

"""Scrolling des Bewertungsbereiches (damit alle Bewertungen im HTML geladen werden)"""
def scrolling(counter):
    print('scrolling...')
    scrollable_div = driver.find_element_by_xpath(
        '//div[@class="XltNde tTVLSc"]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]')
    x = 0
    for _i in range(counter):
        scrolling = driver.execute_script(
            'document.getElementsByClassName("dS8AEf")[0].scrollTop = document.getElementsByClassName("dS8AEf")[0].scrollHeight',
            scrollable_div
        )
        x = x + 1
        print(str(x) + " scroll ._.")
        time.sleep(3)

"""checkt ob xpath valid"""
def check_exists_by_xpath(xpath, data):
    try:
        data.find_element_by_xpath(xpath)
    except selenium.common.exceptions.NoSuchElementException:
        return False
    return True

"""Data in .xlsx schreiben"""
def write_to_xlsx(data):
    print('In excel schreiben...')
    cols = ["name", "weiteres", "vor", "rating", "kommentar"]
    df = pd.DataFrame(data, columns=cols)
    df.to_excel('Ergebnis_Python_Webscraping.xlsx')

"""Main-Methode"""
if __name__ == "__main__":

    print('starting...')
    options = webdriver.ChromeOptions()
    #options.add_argument("--headless")  # auskommentieren, um Browser zu beobachten
    options.add_argument("--lang=en-US")
    options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
    DriverPath = DriverLocation
    driver = webdriver.Chrome(DriverPath, options=options)

    driver.get(URL)
    driver.find_element_by_xpath('/html/body/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button').click()
    time.sleep(5)
    driver.get(URL)

    counter = counter()
    scrolling(counter)

    data = get_data(driver)
    driver.close()

    write_to_xlsx(data)
    print('Fertig!')
