import math
import threading
import time
import xlsxwriter
import requests
import smtplib
from tkinter.ttk import Progressbar
from bs4 import BeautifulSoup
from datetime import datetime
from tkinter import *
from colorama import init, Fore, Back, Style
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from email.message import EmailMessage

init(convert=True)

root = Tk()
root.minsize(480, 234)
root.title("Competition Tool - Nofi")
progress = Progressbar(root, orient=HORIZONTAL,
                       length=300, mode='determinate')

progress.pack(pady=15)

EMAIL_ADDRESS_SENDER = "competitiontoolnofi@gmail.com"
EMAIL_PASSWORD_SENDER = "nofi0000!"
EMAIL_ADDRESS_RECIEVER = "ailonlivny@gmail.com"

def retrieve_input():
    global EMAIL_ADDRESS_RECIEVER
    EMAIL_ADDRESS_RECIEVER = textBox.get("1.0","end-1c")


def percentage(part, whole):
    return 100 * float(part) / float(whole)


def init_logic_thread(event):
    logic_thread = threading.Thread(target=run)
    logic_thread.start()


def put_string_between_job_title_or_location(i_string, i_job_title_or_location):
    job_title_with_hyphen = i_job_title_or_location.split(" ")
    job_title_with_hyphen = ''.join(map(str, [word + i_string for word in job_title_with_hyphen]))
    return job_title_with_hyphen[:len(job_title_with_hyphen) - len(i_string)]


def run():
    button_1.config(text="Running")

    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    workbook = xlsxwriter.Workbook('Competition.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', f'Date :{dt_string}')
    worksheet.write('A2', 'JobTitle')
    worksheet.write('B2', 'location')
    worksheet.write('C2', 'JobID')
    worksheet.write('D2', 'Account')
    worksheet.write('E2', 'Sponsored By Pando')
    worksheet.write('F2', 'Sponsored By indeed at indeed')
    worksheet.write('G2', 'Date post by Indeed at indeed')
    worksheet.write('H2', 'Sponsored By appcast at indeed')
    worksheet.write('I2', 'Date post by appcast at indeed')
    worksheet.write('J2', 'Sponsored By indeed at Jobcase')
    worksheet.write('K2', 'Date post by Indeed at Jobcase')
    worksheet.write('L2', 'Sponsored By appcast at Jobcase')
    worksheet.write('M2', 'Date post by appcast at Jobcase')
    worksheet.write('N2', 'Sponsored By indeed at Monster')
    worksheet.write('O2', 'Date post by Indeed at Monster')
    worksheet.write('P2', 'Sponsored By appcast at Monster')
    worksheet.write('Q2', 'Date post by appcast at Monster')
    worksheet.write('R2', 'Similar Pando Campaign')
    worksheet.write('S2', 'Sourcer')
    worksheet.write('T2', 'Sponsored indication on URL')

    sponsored_by_indeed_at_indeed_col = 5
    sponsored_by_Indeed_at_indeed_date_col = 6
    sponsored_by_Appcast_at_indeed_col = 7
    sponsored_by_Appcast_at_indeed_date_col = 8
    sponsored_by_Indeed_at_Jobcase_col = 9
    sponsored_by_Indeed_at_Jobcase_date_col = 10
    sponsored_by_Appcast_at_Jobcase_col = 11
    sponsored_by_Appcast_at_Jobcase_date_col = 12
    sponsored_by_indeed_at_Monster_col = 13
    sponsored_by_indeed_at_Monster_date_col = 14
    sponsored_by_Appcast_at_Monster_col = 15
    sponsored_by_Appcast_at_Monster_date_col = 16

    Indeed_at_Indeed_count = 0
    Appcast_at_Indeed_count = 0
    Indeed_at_Monster_count = 0
    Appcast_at_Monster_count = 0
    Indeed_at_Jobcase_count = 0
    Appcast_at_Jobcase_count = 0

    low_val = 0
    excel_row = 2
    excel_col = 0
    amazon_force_URL = "https://amazon.force.com/"
    soup_amazon_force = BeautifulSoup(requests.get(amazon_force_URL).text, "html.parser")
    results_from_amazon_force = soup_amazon_force.find_all('div', attrs={'class': 'col-xs-12 col-sm-9'})
    post_count = len(results_from_amazon_force)
    # driver = webdriver.Chrome(executable_path=ChromeDriverManager("2.42").install())
    # driver.delete_all_cookies()

    for idx, res in enumerate(results_from_amazon_force):
        job_location = res.find('span').text.strip()

        if "United States" in job_location:
            job_title = res.find('a').text.strip()
            job_id = res.find('strong').text.strip()
            worksheet.write(excel_row, excel_col, job_title)
            excel_col += 1
            worksheet.write(excel_row, excel_col, job_location)
            excel_col += 1
            worksheet.write(excel_row, excel_col, job_id)

            job_title_with_pluses = put_string_between_job_title_or_location("+", job_title)
            job_location_with_pluses = put_string_between_job_title_or_location("+", job_location)

            indeed_search_job_and_location_url = f"https://www.indeed.com/jobs?q={job_title_with_pluses}&l={job_location_with_pluses}"

            job_title_with_hyphen = put_string_between_job_title_or_location('-', job_title)
            job_location_with_hyphen = put_string_between_job_title_or_location('-', job_location)

            monster_search_job_and_location_url = f"https://www.monster.com/jobs/search/?q={job_title_with_hyphen}&where={job_location_with_hyphen}"

            job_title_with_twenty_percentages = put_string_between_job_title_or_location("%20", job_title)
            job_location_with_twenty_percentages = put_string_between_job_title_or_location("%20", job_location)

            jobcase_search_job_and_location_url = f"https://www.jobcase.com/jobs/results?q={job_title_with_twenty_percentages}&l={job_location_with_twenty_percentages}&radius=25&sort_order=DEFAULT"

            while True:
                try:
                    while True:
                        try:
                            driver = webdriver.Chrome(executable_path=ChromeDriverManager("2.42").install())
                            driver.delete_all_cookies()
                            driver.get(jobcase_search_job_and_location_url)
                            time.sleep(5)  # 5 sec is the max time for the page to upload
                            break
                        except WebDriverException as e:
                            time.sleep(1)

                    jobscase_posts = driver.find_elements_by_xpath(
                        "//div[@class='JobResult__TextContainer-tptubx-18 eGTDZv']")  # Info post wrap by this div

                    is_Appcast_sponsored = False
                    is_indeed_sponsored = False

                    for jobcase_idx, jobscase_post in enumerate(jobscase_posts):
                        jobcase_time = jobscase_post.find_element_by_xpath(
                            ".//div[@class='Typography__Component-sc-1n7rekq-0 dWDyvW JobResult__JobResultTypography-tptubx-0 JobResult__DaysPosted-tptubx-7 iwYRBv']").text
                        jobcase_job_title = jobscase_post.find_element_by_xpath(
                            ".//a[@class='Link__LinkComponent-sc-1kbt8hh-0 iRnhgS']").text
                        jobcase_job_location = \
                            jobscase_post.find_element_by_xpath(
                                ".//div[@class='Typography__Component-sc-1n7rekq-0 dWDyvW JobResult__JobResultTypography-tptubx-0 JobResult__JobLocation-tptubx-6 jlTrOP']").text.split(
                                ",")[0]
                        jobcase_company = jobscase_post.find_element_by_xpath(
                            ".//div[@class='Typography__Component-sc-1n7rekq-0 lmhswq JobResult__JobResultCompanyTypography-tptubx-1 jXlTxF']").text
                        if "Amazon HVH" == jobcase_company or "Workforce" in jobcase_company:
                            if job_title in jobcase_job_title and jobcase_job_location in job_location:
                                if "Workforce" in jobcase_company:
                                    worksheet.write(excel_row, sponsored_by_Appcast_at_Jobcase_col, "Appcast")
                                    worksheet.write(excel_row, sponsored_by_Appcast_at_Jobcase_date_col,
                                                    jobcase_time)
                                    Appcast_at_Jobcase_count += 1
                                    is_Appcast_sponsored = True
                                elif "Amazon HVH" == jobcase_company:
                                    driver.find_element_by_xpath(
                                        "//a[@class='Link__LinkComponent-sc-1kbt8hh-0 iRnhgS']").click()
                                    current_window = driver.window_handles[1]
                                    window_before = driver.window_handles[0]
                                    driver.switch_to.window(current_window)
                                    driver.find_element_by_xpath(
                                        "//a[contains(text(),'Apply')]").click()

                                    try:  # End case for URL is down at landing page

                                        landing_page_window = driver.window_handles[2]
                                        driver.switch_to.window(landing_page_window)
                                        jobcast_amazon_landing_page_id = driver.find_element_by_xpath(
                                            "//p[@class='first']").text.replace('\n',
                                                                                '').replace(
                                            '\t', '')[8:17]
                                        if jobcast_amazon_landing_page_id == job_id:
                                            if "Amazon HVH" == jobcase_company:
                                                worksheet.write(excel_row, sponsored_by_Indeed_at_Jobcase_col, "Indeed")
                                                worksheet.write(excel_row, sponsored_by_Indeed_at_Jobcase_date_col,
                                                                jobcase_time)
                                                Indeed_at_Jobcase_count += 1
                                                is_indeed_sponsored = True
                                    except NoSuchElementException:
                                        pass

                                    driver.close()
                                    driver.switch_to.window(current_window)
                                    driver.close()
                                    driver.switch_to.window(window_before)

                        if jobcase_idx == 10 or (is_indeed_sponsored and is_Appcast_sponsored):
                            break
                    break
                except WebDriverException as e:
                    time.sleep(1)
                    driver = webdriver.Chrome(executable_path=ChromeDriverManager("2.42").install())
                    driver.delete_all_cookies()

            while True:
                try:
                    monster_first_page = ""
                    while monster_first_page == "":
                        try:
                            monster_first_page = requests.get(monster_search_job_and_location_url)
                        except requests.exceptions.RequestException as e:
                            time.sleep(1)

                    soup_search_jobs_at_monster = BeautifulSoup(monster_first_page.text, "html.parser")
                    posts_at_monster = soup_search_jobs_at_monster.find_all('div', attrs={'class': 'flex-row'},
                                                                            limit=10)

                    for post_at_monster in posts_at_monster:
                        monster_job_title = post_at_monster.find('a')
                        if monster_job_title:
                            monster_job_title = monster_job_title.text.strip()
                            monster_post_company = post_at_monster.find('div', attrs={'class': 'company'}).find('span',
                                                                                                                attrs={
                                                                                                                    'class': 'name'}).text.strip()
                            monster_post_date = post_at_monster.find('time').text.strip()
                            monster_job_location = \
                                post_at_monster.find('div', attrs={'class': 'location'}).find('span', attrs={
                                    'class': 'name'}).text.strip().split(",")[0]
                            monster_post_url = post_at_monster.find('h2', attrs={'class': 'title'}).find('a')["href"]
                            if "Amazon HVH" == monster_post_company or "Workforce" in monster_post_company:
                                if job_title in monster_job_title and monster_job_location in job_location:
                                    if "Workforce" in monster_post_company:
                                        worksheet.write(excel_row, sponsored_by_Appcast_at_Monster_col, "Appcast")
                                        worksheet.write(excel_row, sponsored_by_Appcast_at_Monster_date_col,
                                                        monster_post_date)
                                        Appcast_at_Monster_count += 1
                                    elif "Amazon HVH" == monster_post_company:
                                        while True:
                                            try:
                                                driver.get(monster_post_url)
                                                time.sleep(5)
                                                break
                                            except WebDriverException as e:
                                                time.sleep(1)

                                        apply_at_amazon_button = driver.find_element_by_xpath(
                                            "//button[@class='btn job-apply-button bg-primary-lt px-4']")
                                        apply_at_amazon_button.click()
                                        current_window = driver.window_handles[1]
                                        window_before = driver.window_handles[0]
                                        driver.switch_to.window(current_window)

                                        try:  # End case for URL is down at landing page
                                            monster_amazon_landing_page_id = driver.find_element_by_xpath(
                                                "//p[@class='first']").text.replace('\n',
                                                                                    '').replace(
                                                '\t', '')[8:17]

                                            if monster_amazon_landing_page_id == job_id:
                                                worksheet.write(excel_row, sponsored_by_indeed_at_Monster_col, "Indeed")
                                                worksheet.write(excel_row, sponsored_by_indeed_at_Monster_date_col,
                                                                monster_post_date)
                                                Indeed_at_Monster_count += 1

                                        except NoSuchElementException:
                                            pass

                                        driver.close()
                                        driver.switch_to.window(window_before)
                    driver.quit()
                    break
                except WebDriverException as e:
                    time.sleep(1)
                    driver = webdriver.Chrome(executable_path=ChromeDriverManager("2.42").install())
                    driver.delete_all_cookies()

            indeed_first_page = ""
            while indeed_first_page == "":
                try:
                    indeed_first_page = requests.get(indeed_search_job_and_location_url)
                except requests.exceptions.RequestException as e:
                    time.sleep(1)

            soup_search_jobs_at_indeed = BeautifulSoup(indeed_first_page.text, "html.parser")
            indeed_posts = soup_search_jobs_at_indeed.find_all('div', attrs={'data-tn-component': 'organicJob'},
                                                               limit=10)
            app_cast_posts = soup_search_jobs_at_indeed.find_all('div', attrs={'data-empn': '7058506697514818'},
                                                                 limit=10)

            for indeed_post in indeed_posts:
                indeed_time = indeed_post.find('span', attrs={'class': "date"}).text.strip()
                indeed_job_title = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"}).text.strip()
                indeed_url = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"})["href"]
                indeed_job_location = \
                    indeed_post.find("span",
                                     attrs={"class": "location accessible-contrast-color-location"}).text.strip().split(
                        ",")[0]

                if job_title in indeed_job_title and indeed_job_location in job_location:
                    indeed_sec_page = ""
                    while indeed_sec_page == "":
                        try:
                            indeed_sec_page = requests.get("http://indeed.com" + indeed_url)
                        except requests.exceptions.RequestException as e:
                            time.sleep(1)

                    soup_second_page_indeed = BeautifulSoup(indeed_sec_page.text, "html.parser")
                    indeed_third_page = soup_second_page_indeed.find("span", attrs={"id": "originalJobLinkContainer"})
                    if indeed_third_page:
                        indeed_third_page = indeed_third_page.find('a')["href"]
                    else:
                        continue

                    amazon_landing_page = ""
                    while amazon_landing_page == "":
                        try:
                            amazon_landing_page = requests.get(indeed_third_page)
                        except requests.exceptions.RequestException as e:
                            time.sleep(1)

                    soup_amazon_landing_page = BeautifulSoup(amazon_landing_page.text, "html.parser")
                    job_id_amazon = soup_amazon_landing_page.find("div", attrs={"class": "details-line"})
                    if job_id_amazon:
                        job_id_amazon = job_id_amazon.find("p").text.replace('\n', '').replace('\t', '')[7:16]
                    if job_id == job_id_amazon:
                        worksheet.write(excel_row, sponsored_by_indeed_at_indeed_col, "Indeed")
                        worksheet.write(excel_row, sponsored_by_Indeed_at_indeed_date_col, indeed_time)
                        Indeed_at_Indeed_count += 1
                        break

            for app_cast_post in app_cast_posts:
                company_name = app_cast_post.find('a', attrs={'data-tn-element': "companyName"}).text.strip()
                app_cast_job = app_cast_post.find('a', attrs={'data-tn-element': "jobTitle"}).text.strip()
                app_cast_location = app_cast_post.find("div", attrs={
                    "class": "location accessible-contrast-color-location"}).text.strip().split(",")[0]
                if ("Workforce" in company_name) and (job_title in app_cast_job) and (
                        app_cast_location in job_location):
                    appcast_date = app_cast_post.find('span', attrs={'class': "date"}).text.strip()
                    worksheet.write(excel_row, sponsored_by_Appcast_at_indeed_col, "Appcast")
                    worksheet.write(excel_row, sponsored_by_Appcast_at_indeed_date_col, appcast_date)
                    Appcast_at_Indeed_count += 1
                    break

            excel_col = 0
            excel_row += 1

        percentage_till_now = math.floor(percentage(idx, post_count))

        print(
            Fore.CYAN + f"Work process -> {percentage_till_now}%, " + Fore.YELLOW + f"Sponsored at Indeed By: Indeed = {Indeed_at_Indeed_count}, Appcast = {Appcast_at_Indeed_count}, " + Fore.RED + f"Sponsored at Monster By: Indeed = {Indeed_at_Monster_count}, Appcast = {Appcast_at_Monster_count}, " + Fore.LIGHTCYAN_EX + f"Sponsored at jobcast By Indeed= {Indeed_at_Jobcase_count}, Appcast = {Appcast_at_Jobcase_count}")

        if low_val <= percentage_till_now:
            low_val += 5
            progress['value'] = low_val
            root.update_idletasks()

    workbook.close()
    button_1.config(text="Finished")
    driver.quit()
    print(
        Fore.LIGHTCYAN_EX + "Done!" + Fore.RED + " Done!" + Fore.YELLOW + " Done!" + Fore.CYAN + " Done!")

    msg = EmailMessage()
    msg['Subject'] = "Pandologic Competition file"
    msg['From'] = EMAIL_ADDRESS_SENDER
    msg['To'] = EMAIL_ADDRESS_RECIEVER
    msg.set_content("This is a automatic message from Nofi application, competition file was created")

    with open('Competition.xlsx', 'rb') as Competition_file:
        Competition_file_to_attach = Competition_file.read()

    msg.add_attachment(Competition_file_to_attach, maintype="application", subtype="xlsx", filename='Competition.xlsx')

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ADDRESS_SENDER, EMAIL_PASSWORD_SENDER)
        smtp.send_message(msg)

label = Label(root, text="Email: ")
label.pack()
textBox = Text(root, height=1, width=35)
textBox.pack()
button_1 = Button(root, text="Start", fg="blue", height=4,
                  width=15,command=lambda: retrieve_input())
button_1.bind("<Button-1>", init_logic_thread)
button_1.pack()
root.mainloop()










