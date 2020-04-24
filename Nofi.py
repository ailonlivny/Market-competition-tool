import math
import time
import xlsxwriter
import requests
from tkinter.ttk import Progressbar
from bs4 import BeautifulSoup
from datetime import datetime
from tkinter import *
from colorama import Fore, Style

root = Tk()
root.minsize(480, 234)
root.title("Competition Tool - Nofi")
progress = Progressbar(root, orient=HORIZONTAL,
                       length=300, mode='determinate')

progress.pack(pady=15)


def percentage(part, whole):
    return 100 * float(part) / float(whole)


def run_app(event):
    button_1.config(text="Running")
    root.update()

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
    worksheet.write('F2', 'Sponsored By indeed')
    worksheet.write('G2', 'Date post by Indeed')
    worksheet.write('H2', 'Sponsored By appcast')
    worksheet.write('I2', 'Date post by appcast')
    worksheet.write('J2', 'Posting Date')
    worksheet.write('K2', 'Similar Pando Campaign')
    worksheet.write('L2', 'Sourcer')
    worksheet.write('M2', 'Sponsored indication on URL')

    indeed_posts_count = 0
    appcast_posts_count = 0
    low_val = 0
    excel_row = 2
    excel_col = 0
    sponsored_by_indeed_col = 5
    sponsored_by_appcast_col = 7
    URL = "https://amazon.force.com/"
    soup = BeautifulSoup(requests.get(URL).text, "html.parser")
    results_from_amazon_force = soup.find_all('div', attrs={'class': 'col-xs-12 col-sm-9'})
    post_count = len(results_from_amazon_force)
    post_index = 1
    for idx, res in enumerate(results_from_amazon_force):
        job_location = res.find('span').text.strip()
        if "United States" in job_location:
            job_title = res.find('a').text.strip()
            job_id_pando = res.find('strong').text.strip()
            worksheet.write(excel_row, excel_col, job_title)
            excel_col += 1
            worksheet.write(excel_row, excel_col, job_location)
            excel_col += 1
            worksheet.write(excel_row, excel_col, job_id_pando)

            job_title_with_pluses = job_title.split(" ")
            job_title_with_pluses = ''.join(map(str, [word + "+" for word in job_title_with_pluses]))
            job_title_with_pluses = job_title_with_pluses[:len(job_title_with_pluses) - 1]

            job_location_with_pluses = job_location.split(" ")
            job_location_with_pluses = ''.join(map(str, [word + "+" for word in job_location_with_pluses]))
            job_location_with_pluses = job_location_with_pluses[:len(job_location_with_pluses) - 1]

            indeed_search_job_and_location_url = f"https://www.indeed.com/jobs?q={job_title_with_pluses}&l={job_location_with_pluses}"

            indeed_first_page = ""
            while indeed_first_page == "":
                try:
                    indeed_first_page = requests.get(indeed_search_job_and_location_url)
                except requests.exceptions.RequestException as e:
                    time.sleep(1)

            soup = BeautifulSoup(indeed_first_page.text, "html.parser")
            indeed_posts = soup.find_all('div', attrs={'data-tn-component': 'organicJob'}, limit=10)
            app_cast_posts = soup.find_all('div', attrs={'data-empn': '7058506697514818'}, limit=10)

            for indeed_post in indeed_posts:
                root.update()
                indeed_time = indeed_post.find('span', attrs={'class': "date"}).text.strip()
                indeed_job = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"})
                indeed_url = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"})["href"]
                indeed_location = \
                    indeed_post.find("span",
                                     attrs={"class": "location accessible-contrast-color-location"}).text.strip().split(
                        ",")[0]

                if job_title in indeed_job.text.strip() and indeed_location in job_location:
                    indeed_sec_page = ""
                    while indeed_sec_page == "":
                        try:
                            indeed_sec_page = requests.get("http://indeed.com" + indeed_url)
                        except requests.exceptions.RequestException as e:
                            time.sleep(1)

                    soup2 = BeautifulSoup(indeed_sec_page.text, "html.parser")
                    indeed_third_page = soup2.find("span", attrs={"id": "originalJobLinkContainer"})
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

                    soup3 = BeautifulSoup(amazon_landing_page.text, "html.parser")
                    job_id_amazon = soup3.find("div", attrs={"class": "details-line"})
                    if job_id_amazon:
                        job_id_amazon = job_id_amazon.find("p").text.replace('\n', '').replace('\t', '')[7:16]
                    if job_id_pando == job_id_amazon:
                        worksheet.write(excel_row, sponsored_by_indeed_col, "indeed")
                        worksheet.write(excel_row, sponsored_by_indeed_col + 1, indeed_time)
                        indeed_posts_count += 1
                        break

            for app_cast_post in app_cast_posts:
                company_name = app_cast_post.find('a', attrs={'data-tn-element': "companyName"}).text.strip()
                app_cast_job = app_cast_post.find('a', attrs={'data-tn-element': "jobTitle"}).text.strip()
                app_cast_location = app_cast_post.find("div", attrs={
                    "class": "location accessible-contrast-color-location"}).text.strip().split(",")[0]
                if ("Workforce" in company_name) and (job_title in app_cast_job) and (
                        app_cast_location in job_location):
                    appcast_time = app_cast_post.find('span', attrs={'class': "date"}).text.strip()
                    worksheet.write(excel_row, sponsored_by_appcast_col, "Appcast")
                    worksheet.write(excel_row, sponsored_by_appcast_col + 1, appcast_time)
                    appcast_posts_count += 1
                    break

            root.update()
            excel_col = 0
            excel_row += 1

        percentage_till_now = math.floor(percentage(idx, post_count))
        print(
            Fore.BLUE + f"Work process -> {percentage_till_now}%, " + Fore.YELLOW + f"Indeed = {indeed_posts_count}, " + Fore.RED + f"Appcast = {appcast_posts_count}")

        if low_val <= percentage_till_now:
            low_val += 10
            progress['value'] = low_val
            root.update_idletasks()

    workbook.close()
    button_1.config(text="Finished")
    root.update()


button_1 = Button(root, text="Start", fg="blue", height=4,
                  width=15)
button_1.bind("<Button-1>", run_app)
button_1.pack()
root.mainloop()
