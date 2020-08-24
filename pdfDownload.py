import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ActionChains
import pandas as pd
from pathlib import Path


def pdf_downloader():
    mypath = Path().absolute()

    # Defining path of the files

    pdfpath = mypath / 'All_Pdf_Files'
    download_dir = str(pdfpath)

    # date=date.today()
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs',
                                           {
                                               "download.default_directory": download_dir,
                                               "download.prompt_for_download": False,
                                               "download.directory_upgrade": True,
                                               "plugins.always_open_pdf_externally": True
                                           })
    # chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--window-size=1280,1160")

    driver = webdriver.Chrome(options=chrome_options, executable_path="chromedriver.exe")

    def waits(el):
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, el)))

    df = pd.read_excel(r"./All_Files/OMG 0826 Chemo Auth Report.xlsx", sheet_name='ApptData - 2020-08-19T144313.36')

    ##Login to site

    driver.get('https://webchartnow.com/omg/webchart.cgi')
    driver.maximize_window()
    waits('//*[@id="login_user"]')
    driver.find_element_by_xpath('//*[@id="login_user"]').send_keys('Alectka')  # F29051
    driver.find_element_by_xpath('//*[@id="login_passwd"]').send_keys('Jax$on143')  # Jehan@123
    driver.find_element_by_xpath('//*[@id="login_form"]/table/tbody/tr/td/fieldset/table/tbody/tr[3]/td/input').click()
    waits('//*[@id="patsearch"]/table/tbody/tr/td/div/div/ul')
    print("Current session is {}".format(driver.session_id))
    try:
        for i, row in df.iterrows():
            MrNo = row['MR']
            MrNoIndex = str(MrNo).find('MR')
            MrNo = MrNo[MrNoIndex:]
            MrNo = str(MrNo).split(',')[0]
            print(MrNo)
            driver.find_element_by_xpath('//*[@id="patsearch"]/table/tbody/tr/td/label[1]/input').click()
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="patsearch"]/table/tbody/tr/td/span[2]/input[1]').send_keys(MrNo)
            driver.find_element_by_xpath('//*[@id="patsearch"]/table/tbody/tr/td/span[2]/input[2]').click()

            # time.sleep(3)

            print("Program reached here 3")
            try:

                waits('//*[@id="Billing_20Documents_tab"]/a')
                time.sleep(3)
                driver.find_element_by_xpath('//*[@id="Billing_20Documents_tab"]/a').click()
                print("Program reached here 2")
                # time.sleep(5)
                waits('//*[@id="lv_lv_wc_span"]/table/tbody/tr')
                TableLen = len(driver.find_elements_by_xpath('//*[@id="lv_lv_wc_span"]/table/tbody/tr'))
                print(TableLen)
                ChemoTable = driver.find_element_by_xpath('//*[@id="lv_lv_wc_span"]/table')

                listData = []

                #Date = driver.find_element_by_xpath('//*[@id="lv_lv_wc_span"]/table/tbody/tr[1]/td[2]/a').text
                #print(Date)
                if TableLen != 0:
                    for row in range(TableLen):
                        docID = driver.find_element_by_xpath(
                            '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[1]'.format(row + 1)).text

                        docType = driver.find_element_by_xpath(
                            '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[3]'.format(row + 1)).text

                        if docType == 'Chemo Orders':
                            serviceDate = driver.find_element_by_xpath(
                                '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[2]'.format(row + 1)).text

                            break

                    for row in range(TableLen):
                        Date = driver.find_element_by_xpath(
                                '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[2]'.format(row + 1)).text

                        docType = driver.find_element_by_xpath(
                            '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[3]'.format(row + 1)).text

                        if Date == serviceDate and docType == 'Chemo Orders':
                            driver.find_element_by_xpath(
                                '//*[@id="lv_lv_wc_span"]/table/tbody/tr[{}]/td[3]/a'.format(row + 1)).click()
                            time.sleep(5)
                            # waits('//*[@id="wc_main"]/div/div[2]/a')
                            print(driver.find_element_by_xpath('//*[@id="wc_main"]/div/div[2]/a').get_attribute('href'))

                            # ActionChains(driver).click('//*[@id="wc_main"]/div/div[2]/a').perform()

                            # path = os.path.join(download_dir, MrNo)
                            # try:
                            #     os.mkdir(path)
                            # except:
                            #     os.mkdir(path+"({})".format(row))

                            # chrome_options.set_capability("prefs", "download.default_directory={}".format(path))
                            # chrome_options.set_capability("download.default_directory", path)
                            driver.get(driver.find_element_by_xpath('//*[@id="wc_main"]/div/div[2]/a').get_attribute('href'))
                            driver.back()
                            time.sleep(3)
                            # for count, filename in enumerate(os.listdir(path)):
                            src = os.path.join(download_dir, 'webchart.pdf')
                            dst = os.path.join(download_dir, MrNo + ".pdf")
                            try:
                                os.rename(src, dst)

                            except:
                                dst = os.path.join(download_dir, MrNo + "({}).pdf".format(row))
                                os.rename(src, dst)

                    # data = {
                    #     "DOC ID": docID,
                    #     "SERV DATE": serviceDate,
                    #     "DOC TYPE": docType
                    # }
                    #
                    # listData.append(data)
            except Exception as e:
                print("Couldnt Find Billing Documents")
                pass
            print("Program reached here 1")
            driver.find_element_by_xpath('//*[@id="wc_homeicon"]').click()
            # time.sleep(3)
            waits('//*[@id="patsearch"]/table/tbody/tr/td/div/div/ul')
            # df2 = pd.DataFrame(listData)
            # print(df2.to_string)
    except Exception as e:
        print("Couldnt Scrape the pdf for ", MrNo)
        print("Error is ", str(e))
    driver.close()
    # break


if __name__ == '__main__':
    # Calling function
    pdf_downloader()
