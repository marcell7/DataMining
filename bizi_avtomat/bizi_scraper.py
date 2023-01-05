import sys
import time
import argparse

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException


def init_selenium(driver_path, headless=False):
    options = webdriver.FirefoxOptions()
    if headless:
        options.headless = True
    else:
        options.headless = False
    
    driver = webdriver.Firefox(executable_path=driver_path, options=options)
    
    return driver


def login(driver, username, password):
    driver.get("https://www.bizi.si/prijava/")
    
    wait = WebDriverWait(driver, 10)
    wait.until(lambda x: x.find_element(By.ID, "ctl00_cphMain_loginBox1_UserName")).send_keys(username)
    wait.until(lambda x: x.find_element(By.ID, "ctl00_cphMain_loginBox1_Password")).send_keys(password)
    wait.until(lambda x: x.find_element(By.ID, "ctl00_cphMain_loginBox1_ButtonLogin")).click()
    
    print("Prijavljen")
    
    return driver


def get_company_data(driver, id_):
    company_data = {
        "ID": id_,
        "Čisti_prihodki_od_prodaje": "",
        "Čisti_dobiček": "",
        "Št_zaposlenih": "",
        "Dejavnost_ts_media": ""
    }
    
    driver.get(f"https://www.bizi.si/iskanje?q={id_}")
    
    wait = WebDriverWait(driver, 10)
    try:
        company = wait.until(lambda x: x.find_element(By.XPATH, "(//a[@class='b-link-company'])[1]"))
        print(f"Podjetje: {company.get_attribute('innerText')}")
        company.click()
        
    except TimeoutException:
        print("Brez zadetkov")

        return company_data
        
    try:
        company_data["Čisti_prihodki_od_prodaje"] = wait.until(lambda x: x.find_elements(By.ID, "ctl00_ctl00_cphMain_cphMainCol_CompanyFinancePreview1_repeater_ctl01_labValue"))
    except TimeoutException:
        company_data["Čisti_prihodki_od_prodaje"] = []
        
    try:    
        company_data["Čisti_dobiček"] = wait.until(lambda x: x.find_elements(By.ID, "ctl00_ctl00_cphMain_cphMainCol_CompanyFinancePreview1_repeater_ctl00_labValue"))
    except TimeoutException:
        company_data["Čisti_dobiček"] = []
        
    try:
        company_data["Št_zaposlenih"] = wait.until(lambda x: x.find_elements(By.XPATH, "(//div[@class='col-6 b-attr-value pl-1'])[4]"))
    except TimeoutException:
        company_data["Št_zaposlenih"] = []
        
    try:
        company_data["Dejavnost_ts_media"] = wait.until(lambda x: x.find_elements(By.XPATH, "(//div[@class='col-6 b-attr-value pl-1'])[6]"))
    except TimeoutException:
        company_data["Dejavnost_ts_media"] = []
    
    for k,v in company_data.items():
        if isinstance(v, list):
            if len(v) != 0:
                company_data[k] = v[0].get_attribute("innerText")
            else:
                company_data[k] = ""
     
    return company_data


def export_data(data, file_name="bizi_data.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(file_name, sheet_name="data")
    
    return df


def read_ids_from_excel(file):
    df = pd.read_excel(file, dtype=str)
    ids_ = df.values[:, 0].tolist()

    return ids_


if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog="bizi-scraper")
    parser.add_argument("username")
    parser.add_argument("password")
    parser.add_argument("ids_file")
    parser.add_argument("-s", "--sleep", default=3)
    args = parser.parse_args()
    
    
    ids_ = read_ids_from_excel(args.ids_file)
    
    data = []
    driver = init_selenium("driver/geckodriver.exe")
    driver = login(driver, username=args.username, password=args.password)
    
    try:
        for id_ in ids_:
            company_data = get_company_data(driver, id_)

            data.append(company_data)
            print(f"Pregledal {len(data)} podjetij.")
            time.sleep(args.sleep)
    
    except KeyboardInterrupt:
        print("Prekinjam...")
        export_data(data, file_name=f"bizi_data_partial{len(data)}.xlsx")
        driver.close()
        sys.exit()
        
    
    driver.close()
    export_data(data)
    print(f"Končano! Pregledanih {len(data)} podjetij.")