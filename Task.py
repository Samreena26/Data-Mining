from selenium import webdriver
import time
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

import module_locator
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

my_path = module_locator.module_path()

def setup_chrome_driver(download_path):
    """
    Set up and return a Chrome WebDriver instance with specified download preferences.

    Args:
        download_path (str): The path where files should be downloaded.

    Returns:
        WebDriver: An instance of Chrome WebDriver.
    """
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--ignore-certificate-errors")

    # Set up download preferences
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": my_path + 'Downloads',  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    # Attempt to initialize Chrome driver
    try:
        chrome_driver_path = download_path + "/chromedriver.exe"
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print(f"Error occurred: {e}. Attempting to download and install ChromeDriver.")
        chrome_driver_path = ChromeDriverManager().install()
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.maximize_window()

    return driver


driver = setup_chrome_driver(my_path)

data = []

driver.get("https://www.dallascad.org/SearchAcct.aspx")
time.sleep(2)

df = pd.read_excel(my_path + "/input1.xlsx")
parcel = df.values.tolist()

for p in parcel:
    details = driver.find_element(By.NAME, "txtAcctNum")
    details.clear()
    details.send_keys(p[0].replace('"',""))
    time.sleep(1)

    driver.find_element(By.NAME, ("Button1")).click()
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    records = soup.find('table', {"id":'SearchResults1_dgResults'}).findAll("tr")[2:][0].findAll("td")[1:][0].findAll("a")[0].get("href")
    driver.get('https://www.dallascad.org/'+records)
    time.sleep(1)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    elements = soup.find('div', {'class': "HalfCol"}).findAll("p")[0].get_text().strip().split('\n')


    c = []
    address=neighborhood=mapsco=improvement=land=marketval=""

    try:
        address = elements[1].strip()
    except:
        pass
    try:
        neighborhood = elements[4].strip()
    except:
        pass
    try:
        mapsco = elements[6].strip()
    except:
        pass
    try:
        improvement = soup.find('span', {'id': "ValueSummary1_lblImpVal"}).get_text()
    except:
        pass
    try:
        land = soup.find('span', {'id': "ValueSummary1_pnlValue_lblLandVal"}).get_text()
    except:
        pass
    try:
        marketval= soup.find('span', {'id': "ValueSummary1_pnlValue_lblTotalVal"}).get_text()
    except:
        pass

    a=[address,neighborhood,mapsco,improvement,land,marketval]
    data.append(a)
    print(a)

    df = pd.DataFrame(data, columns=['Address','Neighbourhood','Mapsco','Improvement','Land','Marketval'])
    df.to_excel(my_path + "/output1.xlsx", index=False)

    driver.back()


print("mission done")








