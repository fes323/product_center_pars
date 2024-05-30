import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


proxy = ''
options = Options()
options.add_argument(f'--proxy-server={proxy}')

# create the ChromeDriver instance with custom options
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 5)

page = 1

detail_urls = []
while True:
    time.sleep(random.randint(0, 4))

    driver.get(f'https://productcenter.ru/producers/catalog-produkty-pitaniia-45/page-{page}')
    page += 1

    try:
        cards = driver.find_elements(By.CLASS_NAME, 'cards')
        if len(cards) == 0:
            break
    except:
        break

    for card in cards:
        images = card.find_elements(By.CLASS_NAME, 'image')
        for image in images:
            a_tags = image.find_elements(By.TAG_NAME, 'a')
            for a in a_tags:
                url = a.get_attribute('href')
                detail_urls.append(url)

name_list = []
cities_list = []
address_list = []
phone_list = []
whatsApp_list = []
email_list = []
website_list = []
inn_list = []

data = {
    'Наименование': name_list,
    'Город/область': cities_list,
    'Адрес': address_list,
    'Телефон(ы)': phone_list,
    'WhatsApp': whatsApp_list,
    'E-Mail': email_list,
    'website': website_list,
    'ИНН': inn_list,
}

for detail_url in detail_urls:
    driver.get(detail_url)

    time.sleep(random.randint(0, 6))

    original_window = driver.current_window_handle
    assert len(driver.window_handles) == 1

    try:
        driver.find_element(By.CLASS_NAME, 'tab_contacts').click()
    except:
        print('[ERROR] Контакты не найдены!')
        continue

    try:
        name = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/div[1]/div[2]/h1').text
        name_list.append(name or '')
    except:
        name_list.append('')

    try:
        city = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/div[1]/div[2]/b').text
        cities_list.append(city or '')
    except:
        cities_list.append('')

    try:
        address = driver.find_element(By.XPATH, '//td[contains(@itemprop, "address")]').text
        address_list.append(address or '')
    except:
        address_list.append('')

    try:
        phone = driver.find_element(By.XPATH, '//span[contains(@itemprop, "telephone")]').text
        phone_list.append(phone or '')
    except:
        phone_list.append('')

    try:
        whatsApp = driver.find_element(By.XPATH, '//a[contains(@href, "https://wa.me/")]').text
        whatsApp_list.append(whatsApp or '')
    except:
        whatsApp_list.append('')

    try:
        email = driver.find_element(By.XPATH, '//a[contains(@href, "mailto:")]').text
        email = email.lstrip(':')
        email_list.append(email or '')
    except:
        email_list.append('')

    try:
        inn = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/div[2]/div[1]/div[6]/div[1]/table[2]/tbody/tr[3]/td[2]').text
        inn_list.append(inn or '')
    except:
        inn_list.append('')

    try:
        website_data = driver.find_element(By.ID, 'producer_link')
        website_data.click()

        wait.until(EC.number_of_windows_to_be(2))

        time.sleep(random.randint(0, 5))

        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                website_clear = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[1]/div[1]/p/a')
                website_list.append(str(website_clear.get_attribute('href')) or '')
                driver.close()
                driver.switch_to.window(original_window)
    except:
        website_list.append('')

df = pd.DataFrame(data)
with pd.ExcelWriter('catalog-produkty-pitaniia-45.xlsx') as writer:
    df.to_excel(writer)
