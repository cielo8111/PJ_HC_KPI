import os
import time
import math
import smtplib
from datetime import date, timedelta
from email.mime.text import MIMEText
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Daily KPI Report for 'PJ_HC' (RPA) v1.1

# Load sensitive information from environment variables
hc_url = os.getenv('HC_URL')
hc_login_username = os.getenv('HC_LOGIN_USERNAME')
hc_login_password = os.getenv('HC_LOGIN_PASSWORD')
hc_email_address_1 = os.getenv('HC_EMAIL_ADDRESS_1')
hc_email_address_2 = os.getenv('HC_EMAIL_ADDRESS_2')
hc_email_address_3 = os.getenv('HC_EMAIL_ADDRESS_3')
hc_email_password = os.getenv('HC_EMAIL_PASSWORD')

# Check if all environment variables are set
required_variables = [hc_url, hc_login_username, hc_login_password, hc_email_address_1, hc_email_address_2, hc_email_password]
if not all(required_variables) or any(variable == "" for variable in required_variables):
    raise ValueError("One or more environment variables are not set or empty.")

def wait_and_find_element(driver, by, value):
    return WebDriverWait(driver, 20).until(EC.presence_of_element_located((by, value)))

def translate_korean_to_english(text):
    kr_to_en = {'월': 'Mon', '화': 'Tue', '수': 'Wed', '목': 'Thu', '금': 'Fri', '토': 'Sat', '일': 'Sun'}
    for kr, en in kr_to_en.items():
        if kr in text:
            return text.replace(kr, en, 1)
    return ""

try:
    # Open the browser
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Run Chrome in headless mode (without GUI)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    browser = webdriver.Chrome(options=chrome_options)
    browser.implicitly_wait(10)

    # Access the website
    browser.get(hc_url)

    # Find and input the username
    username_element = wait_and_find_element(browser, By.ID, 'A')
    username_element.send_keys(hc_login_username)

    # Find and input the password
    password_element = wait_and_find_element(browser, By.ID, 'B')
    password_element.send_keys(hc_login_password)

    # Click the login button
    login_btn_element = wait_and_find_element(browser, By.XPATH, '/html/body/div[2]/button')
    login_btn_element.click()

    # Click the KPI menu
    menu_kpi_element = wait_and_find_element(browser, By.XPATH, '//*[@id="L"]/div[28]/button')
    menu_kpi_element.click()

    # Input the start and end dates for the KPI
    previous_day_date = date.today() - timedelta(1)
    previous_day_weekday = previous_day_date.strftime("%Y-%m-%d(%a)")
    previous_day_date = previous_day_date.strftime("%Y-%m-%d")
    time.sleep(3)
    
    # Date Format (eight days ago)
    eight_days_ago = date.today() - timedelta(7)
    eight_days_ago = eight_days_ago.strftime("%Y-%m-%d")

    # Input the start date for the KPI
    element_kpi_start_date = wait_and_find_element(browser, By.ID, 'a101')
    element_kpi_start_date.send_keys(Keys.COMMAND + 'a')
    element_kpi_start_date.send_keys(eight_days_ago)

    # Input the end date for the KPI
    element_kpi_end_date = wait_and_find_element(browser, By.ID, 'a102')
    element_kpi_end_date.send_keys(Keys.COMMAND + 'a')
    element_kpi_end_date.send_keys(previous_day_weekday)

    # Click the KPI menu search button
    element_menu_kpi_search_btn = wait_and_find_element(browser, By.XPATH, '//*[@id="R1"]/table/tbody/tr[3]/td/button')
    element_menu_kpi_search_btn.click()

    # Create a list to store the KPI data
    kpi_list = [[] for _ in range(7)]

    # Loop through the categories and extract the data
    kpi_categories = [0, 1, 4, 5, 19, 20, 21, 27, 24, 25, 23]
    for i in range(0, 7):
        for j in kpi_categories:
            data = browser.find_element(By.ID, 'A0_{}_{}'.format(i, j)).text
            data_mod = translate_korean_to_english(data)
            kpi_list[i].append(data_mod if data_mod else format(math.trunc(float(data)), ','))

    # Create a dataframe to store the KPI data
    kpi_df = pd.DataFrame(kpi_list, columns=['DATE', 'NRU', 'DAU', 'NRU-DAU','PU', 'PU(IOS)', 'PU(AOS)', 'ARPPU', 'SALES(IOS)', 'SALES(AOS)', 'SALES(TOTAL)'])

    # HTML Format
    html = """
    <html>
        <head>
            <style>
                table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                th {{
                    background-color: #427bf5;
                    color: #fafafa;
                    border: 1px solid black;
                    text-align: center;
                    font-size: 12px;
                    padding: 8px;
                }}
                td {{
                    border: 1px solid black;
                    text-align: center;
                    font-size: 12px;
                    padding: 8px;
                }}
                tr:last-child td {{
                    color: #427bf5;
                    background-color: #f2e3b8; 
                }}
            </style>
        </head>
        <body>
            <h3>■Daily KPI Report for 'PJ_HC'</h3>
            <p><font size="2">
                ・JP server / last 7 days<br>
                ・SALES, ARPPU: KRW 基準
            </font></p>
            {table}
            <p><font size="1">
                ※このメールアドレスは送信専用のため、返信は受け付けておりません。<br>
                ※ご意見・ご要望は、担当者までお願いします。
            </font></p>
        </body>
    </html>
    """.format(table=kpi_df.to_html(index=False))

    # Send E-mail
    smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_server.ehlo()
    smtp_server.starttls()
    smtp_server.login(hc_email_address_1, hc_email_password)

    email_msg = MIMEText(html, 'html')
    email_msg['Subject'] = '【HC】Daily KPI Report: {}'.format(previous_day_weekday)

    smtp_server.sendmail(hc_email_address_1, [hc_email_address_2, hc_email_address_3], email_msg.as_string())
    smtp_server.quit()

    print("Daily KPI Report sent successfully!")

except Exception as e:
    print("An error occurred while generating the Daily KPI Report:", e)

finally:
    # Close the browser
    browser.quit()
