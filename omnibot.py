search_term = input("Enter search term: ")

import atexit
import time
import os
import pandas as pd
from datetime import datetime
from pytz import timezone
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


profile = webdriver.FirefoxProfile()
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain,text/x-csv,text/csv,application/vnd.ms-excel,application/csv,application/x-csv,text/csv,text/comma-separated-values,text/x-comma-separated-values,text/tab-separated-values,application/pdf")
# options = Options()
# options.add_argument("--headless")
# driver = webdriver.Firefox(options=options)
driver = webdriver.Firefox(executable_path="./geckodriver", firefox_profile=profile)

def close_driver():
    driver.close()

atexit.register(close_driver)

def load_page(url):
    driver.get(url)
    return driver


driver = load_page("http://www.sourcescrub.com/account/login")
driver.find_element_by_id("Username").send_keys("vrajan@petersonpartners.com")
driver.find_element_by_id("Password").send_keys("@o-c9a3vHiW7iiLKAgw3")
driver.find_element_by_xpath("/html/body/main/div/div/div/form/div[4]/button").click()

filter_applied = load_page("https://www.sourcescrub.com/company/search?query=ZV2KCdaWDAxJ34hJ1pYNDInX8QHNScqxAgzxBcoJ1R5cGUnfsQQNMUfMzMnfidVbml0ZWQgWlgzOHMnXydDb3VudHJ5xSc0MMYnNcQnUydfJ1N0YXRlJ35udWxsxRk5ywwyyww3yww2J34oJ2xhdCd%252BMzkuNzZfJ2xvbid%252BLTk4LjUpXydBcmVhcyd%252BJ05vcnRoZXJuIEFtZXJpY2FfxRHIDm4gRnJlZSBUcmFkZSBBZ3JlZW1lbnRfxx4nKeUA5TIyxGpGcm9tJ34wKSk%253D")
search_bar = WebDriverWait(filter_applied, 10).until(EC.presence_of_element_located((By.ID, "mat-input-0")))
search_bar.send_keys(search_term)
time.sleep(3)
search_bar.send_keys(u'\ue007')
# https://rollout.io/blog/get-selenium-to-wait-for-page-load/
time.sleep(3)


results_num_text = WebDriverWait(search_bar, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-site-layout/div/div/app-company-search/div/h3')))
results_num = int(results_num_text.text.split(" ")[0])
if results_num > 50:
    input(f"There are {results_num} results, and I'm not good enough at this to automatically scroll down, so if you want to get all of the results scroll all the way down manually on your browser.\nOtherwise you'll just get the first 50 results. Come back here and press enter when you're done.")

check_box = WebDriverWait(search_bar, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-site-layout/div/div/app-company-search/div/div[3]/ag-grid-angular/div/div[1]/div[2]/div[1]/div[1]/div/div[1]/app-selection--header-checkbox-cell-renderer/mat-checkbox/label/div')))
time.sleep(1)
check_box.click()

export_button = WebDriverWait(check_box, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-site-layout/div/div/app-company-search/div/div[1]/div[2]/button/span[1]")))
time.sleep(1)
export_button.click()

bell_button = WebDriverWait(export_button, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-site-layout/div/app-header/nav/div[3]/mat-toolbar/app-notifications/div/div/button/span[1]")))
time.sleep(1)
bell_button.click()

to_download = WebDriverWait(bell_button, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-site-layout/div/app-header/nav/div[3]/mat-toolbar/app-notifications/div/div[2]/div/section/section/ul[1]/li[1]")))
time.sleep(1)
to_download.click()


path_to_downloads = "/Users/nickkeating/Downloads"
files = os.listdir(path_to_downloads)
paths = [os.path.join(path_to_downloads, basename) for basename in files]
last_download = max(paths, key=os.path.getctime)

results_df = pd.read_csv(last_download, skiprows=1)
os.remove(last_download)
results_length = len(results_df)
empty_list = [""] * results_length
print(f"Excel file will have {results_length} rows.")

all_columns = results_df.columns.tolist()
columns_to_go_first = reversed(["Company Name", "Executive Title",  "Executive First Name", "Executive Last Name", "Executive Email", "Phone Number",  "Personal Note 1", "Interest/Investment", "LinkedIn Account", "Website", "City", "State"])
blank_columns = ["Personal Note 1", "Interest/Investment"]
for column in columns_to_go_first:
    if column in blank_columns:
        results_df.insert(0, column, empty_list)
    else:
        series = results_df[column].tolist()
        results_df = results_df.drop(columns=[column])
        results_df.insert(0, column, series)

now = datetime.now(tz=timezone('US/Eastern')).strftime("%m_%d_%Y_%H.%M.%S")
results_df.to_excel(f"Sourcescrub_{search_term}_{now}.xlsx", index=False)
