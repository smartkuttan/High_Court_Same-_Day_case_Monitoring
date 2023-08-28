import smtplib
import ssl

from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as ec
import time
import datetime
import urllib.request
import requests
import base64
from threading import Thread
from openpyxl import load_workbook
from queue import Queue
import tkinter as tk
from tkinter import messagebox
from subprocess import CREATE_NO_WINDOW
import email.message, email.policy, email.utils

startflag = True

try:
    wb = load_workbook("Credentials.xlsx")
    party_name = wb.active["A2"].value
    receiver_email = wb.active["B2"].value
    captcha_username = wb.active["C2"].value
    captcha_api_key = wb.active["D2"].value
    smtp_server = wb.active["E2"].value
    port = wb.active["F2"].value
    sender_email = wb.active["G2"].value
    sender_password = wb.active["H2"].value
    start_time = wb.active["I2"].value
    end_time = wb.active["J2"].value
    wb.close()
except:
    startflag = False

status_queue = Queue()


def kerala_case_website():
    global party_name, receiver_email, captcha_username, captcha_api_key
    options = Options()
    service = Service()
    service.creation_flags = CREATE_NO_WINDOW
    # options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--disable-gpu")
    driver = Chrome(options=options, service=service)
    wait = WebDriverWait(driver, 120)
    driver.get("https://hckinfo.kerala.gov.in/digicourt/Casedetailssearch/Statuspartyname")
    img = driver.find_element(By.TAG_NAME, "img")
    print(img.get_attribute("src"))
    urllib.request.urlretrieve(img.get_attribute("src"), "img.jpg")
    with open("img.jpg", "rb") as img:
        img_data = base64.b64encode(img.read())
    data = {
        "userid": captcha_username,
        "apikey": captcha_api_key,
        "data": img_data.decode("utf-8")
    }
    response = requests.post("https://api.apitruecaptcha.org/one/gettext", json=data)
    print(response.json())
    result = response.json()["result"]
    party = driver.find_element(By.ID, "party_name")
    party.send_keys(party_name)
    driver.find_element(By.CSS_SELECTOR, "input[onclick = 'search_by_party(3);']").click()
    now = datetime.datetime.now()
    date_str = f"{now.day:02}-{now.month:02}-{now.year:04}"
    from_date = driver.find_element(By.ID, "from_date")
    from_date.send_keys(date_str)
    # driver.execute_script(f"arguments[0].value = '{date_str}';", from_date)
    to_date = driver.find_element(By.ID, "to_date")
    to_date.send_keys(str(date_str))
    # driver.execute_script(f"arguments[0].value = '{date_str}';", to_date)
    captcha = driver.find_element(By.ID, "captcha_typed_login")
    captcha.send_keys(result)
    driver.execute_script("SearchByPartyname();")
    table = wait.until(ec.presence_of_element_located((By.TAG_NAME, "tbody")))
    rows = table.find_elements(By.TAG_NAME, "tr")
    rows.pop()
    for row in rows:
        a = row.find_element(By.TAG_NAME, "a")
        case_id = a.get_attribute("onclick")
        case_id = case_id.split("'")[1]
        print(case_id)
        status_queue.put("processing case with CNR: " + case_id)
        get_hearing_date(case_id)
    driver.quit()


def get_hearing_date(case_id):
    options = Options()
    service = Service()
    service.creation_flags = CREATE_NO_WINDOW
    # options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--disable-gpu")
    driver = Chrome(options=options, service=service)
    wait = WebDriverWait(driver, 120)
    driver.get("https://hckinfo.kerala.gov.in/digicourt/Casedetailssearch/Viewcasestatusnewtab/" + case_id)
    filing = driver.find_element(By.XPATH, '//*[@id="casedetailsview"]/div/table[1]/tbody/tr[5]/td[4]')
    now = datetime.datetime.now()
    date_str = f"{now.day:02}-{now.month:02}-{now.year:04}"
    if date_str in filing.get_attribute("innerText"):
        send_reminder(case_id, False)
    driver.quit()


def send_reminder(cnr, sup_flag):
    global party_name, receiver_email, captcha_username, captcha_api_key, sender_email, smtp_server, port, sender_password
    # options = Options()
    # service = Service()
    # service.creation_flags = CREATE_NO_WINDOW
    # # options.add_experimental_option("detach", True)
    # options.add_argument("--headless")
    # options.add_argument("--no-sandbox")
    # options.add_argument("--window-size=1280,720")
    # options.add_argument("--disable-gpu")
    now = datetime.datetime.now()
    date_str = f"{now.day:02}-{now.month:02}-{now.year:04}"
    # driver = Chrome(options=options, service=service)
    # wait = WebDriverWait(driver, 10)
     # the code also support using yahoomail service to provide alerts
    # driver.get("https://mail.yahoo.com/d/")
    # username = wait.until(ec.presence_of_element_located((By.ID, "login-username")))
    # username.send_keys("youryahoomailid@yahoo.com")
    # next_btn = driver.find_element(By.ID, "login-signin")
    # next_btn.click()
    # password = wait.until(ec.presence_of_element_located((By.ID, "login-passwd")))
    # password.send_keys("youryahoomailpassword")
    # login = driver.find_element(By.ID, "login-signin")
    # login.click()
    # compose = wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[2]/div/div[1]/nav/div/div[1]/a')))
    # compose.click()
    # to = wait.until(ec.presence_of_element_located((By.ID, "message-to-field")))
    # to.send_keys(receiver_email)
    # subject = driver.find_element(By.CSS_SELECTOR, "input[data-test-id = 'compose-subject']")
    # body = driver.find_element(By.CSS_SELECTOR, "div[data-test-id = 'rte']")
    # if sup_flag:
    #     subject.send_keys(f"No Reply - There is a case registered in the name {party_name} at supreme court")
    #     body.send_keys(f"There was a case registered in the name {party_name} today ({date_str}) at the supreme court. The dairy number for the case is: {cnr}")
    # else:
    #     subject.send_keys(f"No Reply - There is a case registered in the name {party_name} at kerala high court")
    #     body.send_keys(f"There was a case registered in the name {party_name} today ({date_str}) at the kerala high court. The CNR for the case is: {cnr}")
    # send_btn = driver.find_element(By.CSS_SELECTOR, "button[data-test-id = 'compose-send-button']")
    # send_btn.click()
    # time.sleep(2)
    # driver.quit()
    portno = int(port)  # For starttls the smtp code is from here uncomment the previous code for selenium based mail
    smtp_serv = smtp_server
    send_email = sender_email
    rec_email = [receiver_email]
    password = sender_password
    print(send_email)
    print(rec_email)
    message = email.message.EmailMessage(email.policy.SMTP)
    message['To'] = receiver_email
    message['From'] = send_email
    message['Date'] = email.utils.formatdate(localtime=True)
    message['Message-ID'] = email.utils.make_msgid()
    message['Subject'] = f'Dear {party_name} a case has been registered in your name today - Do not reply to this mail.'
    if sup_flag:
        text = f"""There is a case registered in your name today ({date_str}) at the Supreme Court with Diary no.: {cnr}"""
    else:
        text = f"""There is a case registered in your name today ({date_str}) at the Kerala High Court with CNR: {cnr}"""
    message.set_content(text)
    if portno == 465:
        with smtplib.SMTP_SSL(smtp_serv, portno, context=ssl.create_default_context()) as server:
            server.login(send_email, password)
            server.ehlo()
            server.sendmail(send_email, rec_email, message.as_string())
    else:
        with smtplib.SMTP(smtp_serv, portno) as server:
            server.login(send_email, password)
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            server.sendmail(send_email, rec_email, message.as_string())


def supreme_court_website():
    global party_name, receiver_email, captcha_username, captcha_api_key
    options = Options()
    service = Service()
    service.creation_flags = CREATE_NO_WINDOW
    # options.add_experimental_option("detach", True)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--disable-gpu")
    driver = Chrome(options=options, service=service)
    wait = WebDriverWait(driver, 30)
    driver.get("https://main.sci.gov.in/case-status")
    driver.find_element(By.XPATH, '//*[@id="tabbed-nav"]/ul[2]/li[3]/a').click()
    captcha_text = wait.until(ec.presence_of_element_located((By.CSS_SELECTOR, "#cap > font")))
    captcha_input = driver.find_element(By.ID, "ansCaptcha")
    driver.execute_script(f"arguments[0].value = '{captcha_text.get_attribute('innerText')}';", captcha_input)
    time.sleep(2)
    partyname = driver.find_element(By.ID, "partyname")
    driver.execute_script(f"arguments[0].value = '{party_name}'", partyname)
    submit = driver.find_element(By.ID, "getPartyData")
    driver.execute_script("arguments[0].click()", submit)
    table = wait.until(ec.presence_of_element_located((By.ID, "cj")))
    filing_dates = table.find_elements(By.CSS_SELECTOR, "font[color = '#FF00A5']")
    dairy_number = table.find_elements(By.CSS_SELECTOR, "font[color = 'green']")
    filing_dates.pop(0)
    dairy_number.pop(0)
    now = datetime.datetime.now()
    date_str = f"{now.day:02}-{now.month:02}-{now.year:04}"
    for i in range(len(filing_dates)):
        filing_date = filing_dates[i].get_attribute("innerText")
        dairy_no = dairy_number[i].get_attribute("innerText")
        status_queue.put("processing case with diary no.: " + dairy_no)
        print(filing_dates[i].get_attribute("innerText") + " " + dairy_number[i].get_attribute("innerText"))
        if date_str in filing_date:
            send_reminder(dairy_no, True)
    driver.quit()


def run_kerala_bot():
    try:
        status_queue.put("Kerala Case Bot Started")
        kerala_case_website()
        status_queue.put("Kerala Case Bot Finished")
    except Exception as e:
        status_queue.put(f"Kerala Case Bot Stopped Due to Error: {repr(e)}")


def run_supreme_bot():
    try:
        status_queue.put("Supreme Court Case Bot Started")
        supreme_court_website()
        status_queue.put("Supreme Court Case Bot Finished")
    except Exception as e:
        status_queue.put(f"Supreme Court Case Bot Stopped Due to Error: {repr(e)}")


def run_automation():
    global start_time, end_time
    now = datetime.datetime.now()
    hour = now.hour
    while True:
        if int(start_time) <= hour <= int(end_time):
            thread1 = Thread(target=run_kerala_bot, daemon=True)
            thread1.start()
            thread2 = Thread(target=run_supreme_bot, daemon=True)
            thread2.start()
            time.sleep(1800)
            now = datetime.datetime.now()
            hour = now.hour


def update_label():
    try:
        data = status_queue.get(timeout=3)
        label.config(text=data)
    except:
        label.config(text="Waiting")
    finally:
        root.after(1000, update_label)


root = tk.Tk()
root.title("Case Search Automation")
root.geometry("700x200")
label = tk.Label(text="")
label.pack(pady=70)
root.after(1000, update_label)
thread = Thread(target=run_automation, daemon=True)
thread.start()
if startflag:
    root.mainloop()
else:
    messagebox.showerror("Credentials error", "Check the 'Credentials.xlsx' file if it exists or not")

