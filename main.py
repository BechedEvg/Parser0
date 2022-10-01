import xlsxwriter
from bs4 import BeautifulSoup
from googletrans import Translator
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
import os
import re


class DriverChrome:

    def __init__(self):
        self.driver = None
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('headless')
        self.options.add_argument("start-maximized")
        self.options.add_argument('--disable-blink-features=AutomationControlled')
        self.options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.options.add_experimental_option('useAutomationExtension', False)

    def open_browser(self):
        self.driver = webdriver.Chrome(options=self.options,
                                       executable_path=rf"{os.getcwd()}/chromedriver")

        stealth(self.driver,
                languages=["en-US", "en"],
                vendor="Google Inc.",
                platform="Win32",
                webgl_vendor="Intel Inc.",
                renderer="Intel Iris OpenGL Engine",
                fix_hairline=True,
                )

    def close_browser(self):
        self.driver.quit()


def translation(text, lang):
    translator = Translator(service_urls=["translate.googleapis.com"])
    translated_text = translator.translate(text, dest=lang)
    if isinstance(text, list):
        translated_dict = {}
        counter = 0
        for tag in translated_text:
            translated_dict[tag.text] = text[counter]
            counter += 1
        return translated_dict
    return translated_text


def get_html(url):
    browser = DriverChrome()
    browser.open_browser()
    browser.driver.get(url)
    html = browser.driver.find_element(By.XPATH, "/html")
    html = html.get_attribute("innerHTML")
    browser.close_browser()
    return html


def write_excel(lists_entries):
    workbook = xlsxwriter.Workbook('Result.xlsx')
    worksheet = workbook.add_worksheet()
    counter = 1
    for list_element in lists_entries:
        worksheet.write(f'A{counter}', list_element[0])
        worksheet.write(f'B{counter}', list_element[1])
        worksheet.write(f'C{counter}', list_element[2])
        worksheet.write(f'D{counter}', list_element[3])
        worksheet.write(f'F{counter}', list_element[4])
        counter += 1
    workbook.close()


def get_list_tags_blog():
    html_blog = get_html("https://www.siladucha.ru/blog")
    parser_tags = BeautifulSoup(html_blog, "lxml")
    elements = parser_tags.find(class_="collection-item")
    elements = elements.find_all(class_="chip")
    list_tags = []
    for element in elements:
        list_tags.append(element.text)
    return list_tags


def get_write_list(dict_tags):
    write_list = []
    for tag in dict_tags:
        html_search = get_html(f"https://www.google.com/search?q={tag}&num=20&start=0")
        parser = BeautifulSoup(html_search, "lxml")
        block_list = parser.find_all(class_="g Ww4FFb vt6azd tF2Cxc")
        if len(block_list) > 10:
            block_list = block_list[:10]
        for block in block_list:
            description = (block.find(class_=re.compile("VwiC3b yXK7lf MUxGbd yDYNvb lyLwlc")).text).replace(u'\xa0', u' ')
            url = block.find(class_="yuRUbf").find("a")['href']
            title = block.find(class_="LC20lb MBeuO DKV0Md").text
            write_list.append([dict_tags[tag], tag, title, description, url])
    return write_list


def main():
    list_tags = get_list_tags_blog()
    dict_translated_tags = translation(list_tags, 'en')
    write_list = get_write_list(dict_translated_tags)
    write_excel(write_list)


if __name__ == "__main__":
    main()
