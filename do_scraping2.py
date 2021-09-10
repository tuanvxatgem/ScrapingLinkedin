import json
import sys
import time
import xlsxwriter
from configparser import ConfigParser

from Scraper import Scraper

# Loading of configurations
from utils import ComplexEncoder

config = ConfigParser()
config.read('config.ini')
import csv
from parsel import Selector
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


def create_search_url(title, location, *include):
    result = ""
    base_url = "http://www.google.com/search?q=+-intitle:%22profiles%22+site:linkedin.com/in/+OR+site:linkedin.com/pub/"

    quote = lambda x: "%22" + x + "%22"

    result += base_url
    result += quote(title) + "+" + quote(location)

    for word in include:
        result += "+" + quote(word)

    return result

max_page = int(config.get("search", "max_page"))
all_urls = []

options = webdriver.ChromeOptions()
# options.add_argument('headless')

# specifies the path to the chromedriver.exe
driver = webdriver.Chrome('/media/vuxuantuan/01D34498C386B760/Work From Home/Scrapper/chromedriver')

# always start from page 1
page = 1
import pandas as pd

keyword = config.get("search", "keyword")
address = config.get("search", "address")
driver.get(create_search_url(keyword, address))
sleep(10)
while True:
    sleep(60)

    # find the urls
    urls = driver.find_elements_by_class_name('g')
    # print(urls)
    urls = [url.find_element_by_tag_name('a') for url in urls]
    urls = [url.get_attribute("href") for url in urls]
    for index, link_url in enumerate(urls):
        if "https://www.linkedin.com/in" not in link_url:
            urls[index]=link_url.replace(link_url[:link_url.find(".linkedin.com/in")], "https://www")
    # print(urls)
    all_urls = all_urls + urls

    # move to the next page
    page += 1

    if page > max_page:
        print('\n end at page:' + str(page - 1))
        break

    try:
        next_page = driver.find_element_by_css_selector("a[aria-label='Page " + str(page) + "']")
        next_page.click()
    except:
        print('\n end at page:' + str(page - 1) + ' (last page)')
        break
log = open(config.get("system", "name_file_save_link"),"a+")
driver.quit()
for link_url in all_urls:
    log.write(link_url+"\n")
log.close()
# Setting the execution mode
headless_option = len(sys.argv) >= 2 and sys.argv[1].upper() == 'HEADLESS'


if len(all_urls) == 0:
    print("Please provide an input.")
    sys.exit(0)
count_sheet = 0
profile_urls = []
try:
    len_sheet = int(config.get("property_file","len_sheet"))
except:
    print("len_sheet wrong type")
while count_sheet+len_sheet-1>=len(all_urls):
    profile_urls=all_urls[count_sheet:count_sheet+len_sheet-1]

    # Launch Scraper
    s = Scraper(
        linkedin_username=config.get('linkedin', 'username'),
        linkedin_password=config.get('linkedin', 'password'),
        profiles_urls=profile_urls,
        headless=headless_option
    )

    s.start()
    print("Start scrap!")
    s.join()
    print("Scraping...")

    scraping_results = s.results

    # Generation of XLS file with profiles data
    output_file_name = config.get('profiles_data', 'output_file_name')
    if config.get('profiles_data', 'append_timestamp').upper() == 'Y':
        output_file_name_split = output_file_name.split('.')
        output_file_name = "".join(output_file_name_split[0:-1]) + "_" + str(int(time.time())) + "." + \
                        output_file_name_split[-1]

    workbook = xlsxwriter.Workbook(output_file_name)

    worksheet = workbook.add_worksheet()

    # Headers
    headers = ["Links",'Name', 'Email', 'Skills', 'Jobs']
    for h in range(len(headers)):
        worksheet.write(0, h, headers[h])

    # Content
    for i in range(len(scraping_results)):

        scraping_result = scraping_results[i]

        if scraping_result.is_error():
            data = ['Error'] * len(headers)
        else:
            p = scraping_result.profile
            data = [
                scraping_result.linkedin_url,
                p.name,
                p.email,
                ",".join(p.skills)
            ]

            for job in p.jobs:
                data.append(json.dumps(job.reprJSON(), cls=ComplexEncoder))

        for j in range(len(data)):
            worksheet.write(i + 1, j, data[j])
        print("Write done 1 sheet")
workbook.close()

print("Scraping Ended")