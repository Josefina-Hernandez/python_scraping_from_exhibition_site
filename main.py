from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time

import openpyxl as xl


class Excel_Con():
    def __init__(self, output_name):
        self.output_name = output_name
        self.wb = xl.Workbook()
        self.ws = self.wb.active
        self.ws.append([
            'Index',
            'Company Name',
            'Stand',
            'Facebook Account',
            'Company Website',
            'Company Email',
            'Company Phone',
            'Address',
            'Why Vist Our Stand',
            'Brands We Present',
            'Description'
        ])
        self.wb.save(filename=self.output_name)

    def adding_dataline(self, index, dataline):
        dataline.insert(0, index)
        self.ws.append(dataline)
        self.wb.save(filename=self.output_name)

    def close_excel(self):
        self.wb.close()


class DataScraping():
    def __init__(self, driver_path):
        options = Options()
        options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
        self.driver = webdriver.Chrome(driver_path, options=options)

    def get_single_page_contents(self, url):
        self.driver.get(url=url)
        wait = WebDriverWait(self.driver, 30)
        wait.until(EC.presence_of_element_located((By.ID, 'exhibitor_details_address')))

        html = self.driver.page_source
        soup = BeautifulSoup(str(html), features='html.parser')

        soup_company_name = soup.find('h1', class_='wrap-word')
        if soup_company_name:
            company_name = soup_company_name.text
        else:
            company_name = ''
        print(company_name)

        soup_stand = soup.find('div', class_='display-stands')
        if soup_stand:
            stand_child = soup_stand.find('p')
            stand = stand_child.text
        else:
            stand = ''
        print(stand)

        soup_facebook = soup.find('a', class_='facebook-logo-color')
        if soup_facebook:
            facebook = soup_facebook['href']
        else:
            facebook = ''
        print(facebook)

        soup_website = soup.find('div', id='exhibitor_details_website')
        if soup_website:
            website_child = soup_website.find('p')
            website = website_child.text
        else:
            website = ''
        print(website)

        soup_email = soup.find('div', id='exhibitor_details_email')
        if soup_email:
            email_child = soup_email.find('p')
            email = email_child.text
        else:
            email = ''
        print(email)

        soup_phone = soup.find('div', id='exhibitor_details_phone')
        if soup_phone:
            phone_child = soup_phone.find('p')
            phone = phone_child.text
        else:
            phone = ''
        print(phone)

        soup_address = soup.find('div', id='exhibitor_details_address')
        if soup_address:
            address_child = soup_address.find('p')
            grand_children = address_child.find_all('span')
            address = ''
            for grand_child in grand_children:
                address = address + grand_child.text + '\n'

            address = address[:-1]
        else:
            address = ''
        print(address)

        soup_why = soup.find('div', id='exhibitor_details_showobjective')
        if soup_why:
            why_child = soup_why.find('p')
            why = why_child.text
        else:
            why = ''
        print(why)

        soup_brands = soup.find('div', id='exhibitor_details_brands')
        if soup_brands:
            brands_child = soup_brands.find('p')
            brands = brands_child.text
        else:
            brands = ''
        print(brands)

        soup_description = soup.find('div', id='exhibitor_details_description')
        if soup_description:
            description_child = soup_description.find('p')
            description = description_child.text
        else:
            description = ''
        print(description)

        print('\n\n')
        return [company_name, stand, facebook, website, email, phone, address, why, brands, description]

    def get_page_links(self, url):
        self.driver.get(url=url)
        time.sleep(30)
        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="company-info"]')))

        elem = self.driver.find_element_by_tag_name("body")
        for i in range(10):
            # elem.send_keys(Keys.PAGE_DOWN)
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)  # 等待页面加载
            print(i)

        html = self.driver.page_source

        soup = BeautifulSoup(str(html), features='html.parser')

        soup_company_info = soup.find_all('div', class_='company-info')

        links = []
        for each in soup_company_info:
            soup2 = BeautifulSoup(str(each), features='html.parser')
            soup_company_link = soup2.find('a')
            link = soup_company_link['href']
            links.append(link)

        counter = 0
        for link in links:
            counter += 1
            print(link)

        print('\n', counter)

        return links

    def quit_driver(self):
        self.driver.quit()


if __name__ == '__main__':
    ec = Excel_Con(output_name='manufacturing-expo-list.xlsx')
    data_scraping = DataScraping(driver_path='./chromedriver.exe')

    url = 'https://www.assemblytechexpo.com/en-gb/for-visitors/search-for-exhibitors.html?#/'

    links = data_scraping.get_page_links(url=url)

    index = 0
    for link in links:
        index += 1
        print(index, '!!!!!!!!!!!!', link)
        dataline = data_scraping.get_single_page_contents(url=link)
        ec.adding_dataline(index=index, dataline=dataline)

    ec.close_excel()
    data_scraping.quit_driver()