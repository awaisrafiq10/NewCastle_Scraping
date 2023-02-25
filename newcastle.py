# pip install webdriver-manager
import time
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

Options = webdriver.ChromeOptions()
Options.add_experimental_option("prefs", {"download.default_directory": "D:\1\Education\BZU\1\tutorial"})
Options.add_argument('--allow-running-insecure-content')
Options.add_argument('--ignore-certificate-errors')
# Options.add_argument('--ignore-certificate-errors-spki-list')
Options.add_argument('--ignore-ssl-errors')
# Options.set_capability('acceptInsecureCerts', True)
Options.add_experimental_option('excludeSwitches', ['enable-logging'])
s=Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s, options=Options)

#url launch
driver.get("https://newcastle.metastreet.co.uk/public-register?search%5Bquery%5D=NE&search%5Brecaptcha%5D%5BrecaptchaV3%5D=03AFY_a8USouRLy3obQHGWb6MMZExFHzUuGLn4D3X9iWiY1XF7khqHBztKk-qpduYOiInhH68hy317tq_a3eOLAH3XudzeZfxmLm6T1lG6-bRUy-8LmcYWEHHlV9i-ceea483tUnYye8MtRvxZL_n9BCF9RnzuQKvnamepxaKoC9Vq6nhqe0d9yNSjby3KQoVZstWtkC9YPY0DwGDsPTwtcxuPhh5uT40NheQDyw0tFthGqi4f9pYB7G7afuQUqbR3b36FhaHu3B18oFayBL1qe6SLbe0Dyfc_5KdgXiqNn-Nxszy9iLrdt1BYf4u9OzeBklab1QEyPMm_WP3iRfKg5I1KOMuc1-SVQrU0jOseWWI33Om6pqGyPuTOI_hhwHDhGIEAwrsHvpgWuFySYTqXvjVIAn2JBgpSovmHj6U1CDQtPnaBvAYjbDnBIzhYP8RMfA7RiRowJJQRUZt_hhm-QlEhuguX5nHx5m-kX1qAggHcsR5uQP78HoQ7mXQhK3X1cRNndfYR_NwAdODKQw58b3W6Vy75lmXY3YDQ&search%5Brecaptcha%5D%5BrecaptchaV2%5D=&search%5B_token%5D=255c8fa8599d33d574867c52e3.ucqc6FJCLUDQ7YDstROqqqLKVsX_VzLpbI0Gc1-cOuc.6bnyrBkGGhGXqcna83TFycW_N7aIbgu4JNVJITDXeY_Dnu26OAxKAZmC4w")
driver.maximize_window()
search = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'button')))
search.click()
driver.implicitly_wait(10)
data = {"Name":[], "Address": [], "Date": []}

def searching():
        time.sleep(2)
        driver.refresh()
        time.sleep(2)
        p = 1
        addresses = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.PARTIAL_LINK_TEXT,  'Newcastle Upon Tyne')))
        names = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'text-secondary')))

        for ad, nm in zip(addresses, names):
            # p = 1
            address = ad.text
            data['Address'].append(address)
            name = nm.text
            data['Name'].append(name)
            dat = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, f'''//*[@id="main-content"]/div[2]/div/p[{p}]''')))
            date = dat.text.replace(name, '').strip()
            data['Date'].append(date)
            p += 1

        if page <= 3:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='main-content']/div[2]/div/div/div/div/a[9]"))).click()
        elif page == 4:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='main-content']/div[2]/div/div/div/div/a[10]"))).click()
        elif 5 <= page <= 565:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='main-content']/div[2]/div/div/div/div/a[11]"))).click()
        print("PAGE COMPLETED:  ", page)

# You can set the page range you want to scrape
for page in range(1, 100):
    try:
        searching()
    except (StaleElementReferenceException, NoSuchElementException, TimeoutException) as e:
        # handle the exception
        print(f"AN EXCEPTION OCCURRED: {e}")
        input1 = input("HAVE YOU RESOLVED CAPTCHA & UPDATED THE PAGE TO FOLLOWING ITERATION | WANT TO CONTINUE?(y/n)")
        if input1 == "y":
            pass
        else:
            break
        searching()

time.sleep(1)
driver.quit()
# print(data)
df = pd.DataFrame(data)
df.to_excel("file10.xlsx", index=False)
print("YOUR DATA IS CONVERTED TO EXCEL.")
