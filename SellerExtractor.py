import datetime
import re
import time
import traceback

import xlsxwriter as xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pyuser_agent
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from amazoncaptcha import AmazonCaptcha

site = 'https://www.amazon.com/gp/browse.html?node=2975359011'
#site = 'https://www.amazon.com/gp/browse.html?node=495224'
host = 'https://www.amazon.com'
webdriver_path = 'resources/chromedriver.exe'
listSellerInfo = []
allSellers = []


class Prod:
    linkProduct: ''


def initializeSelenium():
    options = Options()
    options.add_argument(f'user-agent={pyuser_agent.UA().random}')
    return webdriver.Chrome(options=options, executable_path=webdriver_path)


def finalizeSelenium(driver):
    driver.close()
    driver.quit()


def showError(e):
    pass
    # print(f'Error: {e}')
    # print(traceback.format_exc())


def validateAmazonCaptcha(driver):
    try:
        captchaElement = WebDriverWait(driver, 1).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#captchacharacters"))
        )
        captcha = AmazonCaptcha.fromdriver(driver)
        solution = captcha.solve()
        print("CAPTCHA: " + solution)

        time.sleep(1)

        captchaElement.send_keys(solution)

        WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".a-button-text"))
        ).click()
    except:
        pass

def main(site):
    driver = initializeSelenium()
    try:
        count = 0
        while True:
            try:
                driver.delete_all_cookies()
                driver.get(site)
                validateAmazonCaptcha(driver)
                if len(listSellerInfo) == 0:
                    WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='nav-global-location-popover-link']"))
                    ).click()

                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//*[@id='GLUXZipUpdateInput']"))
                    ).find_element(By.XPATH, "//*[@id='GLUXZipUpdateInput']").send_keys("62864")

                    WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='GLUXZipUpdate']/span/input"))
                    ).click()

                    time.sleep(2)

                    driver.execute_script("document.querySelector('#GLUXConfirmClose').click()")

                    time.sleep(5)

                    driver.get(site)
                    validateAmazonCaptcha(driver)

                prods = []
                for row in \
                        WebDriverWait(driver, 3).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, ".s-main-slot.s-result-list"))
                        ).find_elements(By.CSS_SELECTOR,
                                        ".s-main-slot.s-result-list div[data-component-type='s-search-result']"):
                    prod = Prod()
                    prod.linkProduct = row.find_element(By.CSS_SELECTOR, "a.a-link-normal.s-no-outline").get_attribute(
                        "href")
                    prods.append(prod)

                try:
                    site = WebDriverWait(driver, 3).until(
                        EC.visibility_of_element_located(
                            (By.CSS_SELECTOR, ".s-pagination-next"))
                    ).get_attribute("href")
                except Exception as e:
                    site = WebDriverWait(driver, 3).until(
                        EC.visibility_of_element_located(
                            (By.CSS_SELECTOR, ".a-last a"))
                    ).get_attribute("href")

                for prod in prods:
                    driver.get(prod.linkProduct)
                    validateAmazonCaptcha(driver)
                    try:
                        delivery = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, "//*[@id='glow-ingress-line2']"))
                        ).text

                        if delivery.lower() == "brazil":
                            WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[@id='nav-global-location-popover-link']"))
                            ).click()

                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//*[@id='GLUXZipUpdateInput']"))
                            ).find_element(By.XPATH, "//*[@id='GLUXZipUpdateInput']").send_keys("62864")

                            WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[@id='GLUXZipUpdate']/span/input"))
                            ).click()

                            time.sleep(2)

                            driver.execute_script("document.querySelector('#GLUXConfirmClose').click()")

                            time.sleep(5)

                            driver.get(prod.linkProduct)
                            validateAmazonCaptcha(driver)

                        sellers = []
                        try:
                            elSeller = WebDriverWait(driver, 3).until(
                                EC.visibility_of_element_located(
                                    (By.CSS_SELECTOR,
                                     "#sellerProfileTriggerId:not([target='_blank'])")
                                )
                            )
                            seller = {
                                "linkSeller": elSeller.get_attribute("href"),
                                "sellerName": elSeller.text
                            }
                            if seller['sellerName'] and seller['sellerName'].strip().lower() != "amazon warehouse" \
                                    and not any(u['sellerName'].__eq__(seller['sellerName']) for u in allSellers):
                                sellers.append(seller)
                                allSellers.append(seller)

                        except:
                            pass

                        try:
                            WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR,
                                     "#olpLinkWidget_feature_div > div.a-section.olp-link-widget > span > a"))
                            ).click()

                            time.sleep(5)

                            for row in \
                                    driver.find_elements(By.CSS_SELECTOR, "#aod-offer-soldBy a"):
                                seller = {
                                    "linkSeller": row.get_attribute("href"),
                                    "sellerName": row.text
                                }
                                if seller['sellerName'] and seller['sellerName'].strip().lower() != "amazon warehouse" \
                                        and not any(u['sellerName'].__eq__(seller['sellerName']) for u in allSellers):
                                    sellers.append(seller)
                                    allSellers.append(seller)
                        except:
                            pass

                        for seller in sellers:
                            try:
                                time.sleep(1)
                                driver.get(seller['linkSeller'])
                                validateAmazonCaptcha(driver)
                                WebDriverWait(driver, 10).until(
                                    EC.visibility_of_element_located((By.CSS_SELECTOR, "#seller-profile-container"))
                                )
                                address = driver.find_element(By.XPATH,
                                                              "//*[@id='seller-profile-container']/div[2]/div/ul/li[2]/span/ul").text
                                zipcode = ''
                                if address:
                                    regexZipCode = re.search("\d+-?\d+$|\d+-?\d+(?=[\\n,\s]?[a-z\s]+$)", address,
                                                             re.IGNORECASE)
                                    zipcode = regexZipCode.group(0).strip() if regexZipCode else ''
                                    regexState = re.search("(?<=\\n)([a-z] ?)+(?=\\n\d+-?\d+)", address, re.IGNORECASE)
                                    state = regexState.group(0).strip() if regexState else ''
                                    regexCity = re.search("(?<=\\n)([a-z]\s?){1,}(?=\\n([a-z]\s?){1,}(?=\\n\d+-?\d+))",
                                                          address, re.IGNORECASE)
                                    city = regexCity.group(0).strip() if regexCity else ''

                                sellerInfo = [
                                    driver.current_url,
                                    driver.find_element(By.XPATH,
                                                        "//*[@id='seller-profile-container']/div[2]/div/ul/li[1]/span").text.replace(
                                        "Business Name: ", ""),
                                    address.replace('\n', " "),
                                    city,
                                    state,
                                    zipcode
                                ]
                                if not listSellerInfo:
                                    count += 1
                                    sellerInfo.insert(0, count)
                                    print(sellerInfo)
                                    listSellerInfo.append(sellerInfo)
                                else:
                                    if not any(u[2].__eq__(sellerInfo[1]) for u in listSellerInfo):
                                        count += 1
                                        sellerInfo.insert(0, count)
                                        print(sellerInfo)
                                        listSellerInfo.append(sellerInfo)

                            except:
                                pass

                    except Exception as e:
                        showError(e)

                    if len(listSellerInfo) >= 500:
                        break

                if len(listSellerInfo) >= 500:
                    break

                workbook = xlsxwriter.Workbook(
                    "C:\\Users\\55149\PycharmProjects\\freelancers-projects\\Athena Digital\\Amazon USA Seller.xlsx")
                worksheet = workbook.add_worksheet()
                worksheet.write_row(0, 0, ["#", "URL", "Business Name", "Business Address", "City", "State", "Zipcode"])
                for row, data in enumerate(listSellerInfo):
                    worksheet.write_row(row + 1, 0, data)

                workbook.close()

            except Exception as e:
                showError(e)


    except Exception as e:
        showError(e)
    finally:
        finalizeSelenium(driver)


print(datetime.datetime.now())
main(site)
print(datetime.datetime.now())