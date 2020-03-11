import sys
import re
import pprint
from my_functions import mysoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pytesseract
import argparse
from PIL import Image
from subprocess import check_output

asin = 'B0728BF99G'
simu_url = "https://sellercentral.amazon.com/hz/fba/profitabilitycalculator/index?l"


class AmazonFee:

    def __init__(self, asin):
        self.asin=asin

    def fetch_amazon_data(self):
        "Amazon.comから価格とタイトルを取得"
        amazon_us_url = 'https://www.amazon.com/dp/{}'.format(self.asin)
        soup = mysoup(amazon_us_url)
        try:
            price = soup.find(id='priceblock_ourprice').text
        except:
            if 'Robot' in soup.title:
                argparser = argparse.ArgumentParser()
                argparser.add_argument('path', help='Captcha file path')
                args = argparser.parse_args()
                path = args.path
                print('Resolving Captcha')
                captcha_text = self.resolve(path)
                print('Extracted Text', captcha_text)
            print('ERROR price not get.')
        price = re.sub(r'[￥$,]', '', price)
        amazon_japan_url = 'https://www.amazon.co.jp/dp/{}'.format(self.asin)
        soup = mysoup(amazon_japan_url)
        try:
            title = soup.find(id='productTitle').text
        except:
            if not 'ページが見つかりません' in soup.title:
                with open('souptext2.txt', mode='w', encoding='utf-8') as f:
                    f.write(soup.prettify())
                argparser = argparse.ArgumentParser()
                argparser.add_argument('path', help='Captcha file path')
                args = argparser.parse_args()
                path = args.path
                print('Resolving Captcha')
                captcha_text = self.resolve(path)
                print('Extracted Text', captcha_text)

        print(title.strip())
        return price

    def fetch_from_simulator(self, url, price):
        "Amazonシュミレーターから手数料計算"
        options = Options()
        # options.add_argument('--headless')
        options.add_argument('--window-size=1280,1024')
        driver = webdriver.Chrome(executable_path='C:\driver/chromedriver.exe')
#         driver = webdriver.Chrome(executable_path='/usr/local/bin/chromedriver.exe', chrome_options=options)
        driver.get(url)
        print('Get FBA simulator page.')
        input_element = driver.find_element_by_id('search-string')
        input_element.send_keys(asin)
        input_element.send_keys(Keys.RETURN)
        print('Send asin and enter.')

        wait = WebDriverWait(driver, 10)
        print('Waiting for the undisplay...', file=sys.stderr)
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div[aria-hidden]')))

        input_element=driver.find_element_by_id('afn-pricing')
        input_element.send_keys(price)
        driver.find_element_by_id('update-fees-link-announce').click()
        driver.save_screenshot('i.png')
        driver.quit()

    def resolve(self, path):
        print('Resampling the Image')
        check_output(['convert', path, '-resample', '600', path])
        return pytesseract.image_to_string(Image.open(path))

    def get_amazon_image(self):
        "Amazonの商品サムネイルを取得"


def main():
    amazon_ins = AmazonFee(asin)
    price = amazon_ins.fetch_amazon_data()
#     amazon_ins.fetch_from_simulator(simu_url, price)


if __name__ == '__main__':
    main()

