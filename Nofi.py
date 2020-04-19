# import requests
# from bs4 import BeautifulSoup
# import xlsxwriter
# from datetime import datetime
#
# now = datetime.now()
# dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
# workbook = xlsxwriter.Workbook('Competition.xlsx')
# worksheet = workbook.add_worksheet()
# worksheet.write('A1', f'Date :{dt_string}')
# worksheet.write('A2', 'JobTitle')
# worksheet.write('B2', 'location')
# worksheet.write('C2', 'JobID')
# worksheet.write('D2', 'Account')
# worksheet.write('E2', 'Sponsored By Pando')
# worksheet.write('F2', 'Sponsored By indeed')
# worksheet.write('G2', 'Sponsored By appcast')
# worksheet.write('H2', 'Posting Date')
# worksheet.write('I2', 'Similar Pando Campaign')
# worksheet.write('J2', 'Sourcer')
# worksheet.write('K2', 'Sponsored indication on URL')
#
# def get_US_Proxy():
#     proxy_site_URL = "https://www.us-proxy.org/"
#     soup_proxy_site = BeautifulSoup(requests.get(proxy_site_URL).text, "html.parser")
#     proxy_str = soup_proxy_site.find_all('td')
#     proxy_host = proxy_str[0].text
#     proxy_port = proxy_str[1].text
#     return {"https":f'{proxy_host}:{proxy_port}'}
#
#
# excel_row = 2
# excel_col = 0
# sponsored_by_indeed_col = 5
# sponsored_by_appcast_col = 6
# URL = "https://amazon.force.com/"
# soup = BeautifulSoup(requests.get(URL).text, "html.parser")
# results_from_amazon_force = soup.find_all('div', attrs={'class': 'col-xs-12 col-sm-9'})
#
# for res in results_from_amazon_force:
#     job_location = res.find('span').text.strip()
#     if "United States" in job_location:
#         job_title = res.find('a').text.strip()
#         job_id_pando = res.find('strong').text.strip()
#         worksheet.write(excel_row, excel_col, job_title)
#         excel_col += 1
#         worksheet.write(excel_row, excel_col, job_location)
#         excel_col += 1
#         worksheet.write(excel_row, excel_col, job_id_pando)
#
#         job_title_with_pluses = job_title.split(" ")
#         job_title_with_pluses = ''.join(map(str, [word + "+" for word in job_title_with_pluses]))
#         job_title_with_pluses = job_title_with_pluses[:len(job_title_with_pluses) - 1]
#
#         job_location_with_pluses = job_location.split(" ")
#         job_location_with_pluses = ''.join(map(str, [word + "+" for word in job_location_with_pluses]))
#         job_location_with_pluses = job_location_with_pluses[:len(job_location_with_pluses) - 1]
#
#         indeed_search_job_and_location_url = f"https://www.indeed.com/jobs?q={job_title_with_pluses}&l={job_location_with_pluses}&radius=125"
#
#         indeed_first_page = ""
#         while indeed_first_page == "":
#             try:
#                 proxy = {"https":"154.197.133.52:3128"}
#                 print(proxy)
#                 indeed_first_page = requests.get(indeed_search_job_and_location_url, proxies=proxy)
#             except requests.exceptions.RequestException as e:
#                 pass
#         soup = BeautifulSoup(indeed_first_page.text, "html.parser")
#         indeed_posts = soup.find_all('div', attrs={'data-tn-component': 'organicJob'}, limit=10)
#         app_cast_posts = soup.find_all('div', attrs={'data-empn': '7058506697514818'}, limit=10)
#
#         for indeed_post in indeed_posts:
#             indeed_job = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"}).text.strip()
#             indeed_url = indeed_post.find('a', attrs={'data-tn-element': "jobTitle"})["href"]
#             indeed_location = \
#                 indeed_post.find("span",
#                                  attrs={"class": "location accessible-contrast-color-location"}).text.strip().split(
#                     ",")[0]
#
#             if job_title in indeed_job and indeed_location in job_location:
#                 indeed_sec_page = ""
#                 while indeed_sec_page == "":
#                     try:
#                         indeed_sec_page = requests.get("http://indeed.com" + indeed_url)
#                     except requests.exceptions.RequestException as e:
#                         pass
#
#                 soup2 = BeautifulSoup(indeed_sec_page.text, "html.parser")
#                 indeed_third_page = soup2.find("span", attrs={"id": "originalJobLinkContainer"}).find('a')["href"]
#
#                 amazon_landing_page = ""
#                 while amazon_landing_page == "":
#                     try:
#                         amazon_landing_page = requests.get(indeed_third_page)
#                     except requests.exceptions.RequestException as e:
#                         pass
#
#                 soup3 = BeautifulSoup(amazon_landing_page.text, "html.parser")
#                 job_id_amazon = soup3.find("div", attrs={"class": "details-line"})
#                 if job_id_amazon:
#                     job_id_amazon = job_id_amazon.find("p").text.replace('\n', '').replace('\t', '')[7:16]
#                 if job_id_pando == job_id_amazon:
#                     worksheet.write(excel_row, sponsored_by_indeed_col, "indeed")
#                     break
#
#         for app_cast_post in app_cast_posts:
#             company_name = app_cast_post.find('a', attrs={'data-tn-element': "companyName"}).text.strip()
#             app_cast_job = app_cast_post.find('a', attrs={'data-tn-element': "jobTitle"}).text.strip()
#             app_cast_location = app_cast_post.find("div", attrs={
#                 "class": "location accessible-contrast-color-location"}).text.strip().split(",")[0]
#             if job_title in app_cast_job and app_cast_location in job_location:
#                 worksheet.write(excel_row, sponsored_by_appcast_col, "appcast")
#                 break
#
#         print(f"Row number:{excel_row}")
#         excel_col = 0
#         excel_row += 1
#
# workbook.close()
import math
from tkinter.ttk import Progressbar

import requests
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import datetime
from tkinter import *

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
    worksheet.write('F2', 'Sponsored By Others')
    worksheet.write('G2', 'Posting Date')
    worksheet.write('H2', 'Similar Pando Campaign')
    worksheet.write('I2', 'Sourcer')
    worksheet.write('J2', 'Sponsored indication on URL')
    low_val = 0
    excel_row = 2
    excel_col = 0
    sponsored_by_others_col = 5
    URL = "https://amazon.force.com/"
    soup = BeautifulSoup(requests.get(URL).text, "html.parser")
    results_from_amazon_force = soup.find_all('div', attrs={'class': 'col-xs-12 col-sm-9'})
    post_count = len(results_from_amazon_force)
    post_index = 1
    for res in results_from_amazon_force:
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
                    pass

            soup = BeautifulSoup(indeed_first_page.text, "html.parser")
            indeed_posts = soup.find_all('div', attrs={'data-tn-component': 'organicJob'}, limit=10)

            for indeed_post in indeed_posts:
                root.update()
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
                            pass

                    soup2 = BeautifulSoup(indeed_sec_page.text, "html.parser")
                    indeed_third_page = soup2.find("span", attrs={"id": "originalJobLinkContainer"}).find('a')["href"]

                    amazon_landing_page = ""
                    while amazon_landing_page == "":
                        try:
                            amazon_landing_page = requests.get(indeed_third_page)
                        except requests.exceptions.RequestException as e:
                            pass

                    soup3 = BeautifulSoup(amazon_landing_page.text, "html.parser")
                    job_id_amazon = soup3.find("div", attrs={"class": "details-line"})
                    if job_id_amazon:
                        job_id_amazon = job_id_amazon.find("p").text.replace('\n', '').replace('\t', '')[7:16]
                    if job_id_pando == job_id_amazon:
                        worksheet.write(excel_row, sponsored_by_others_col, "indeed")
                        break

            root.update()
            print(f"Row number:{excel_row}")
            excel_col = 0
            excel_row += 1

        percentage_till_now = math.floor(percentage(post_index, post_count))

        if low_val <= percentage_till_now:
            low_val += 10
            progress['value'] = low_val
            root.update_idletasks()

        post_index += 1
    workbook.close()
    button_1.config(text="Finished")
    root.update()


button_1 = Button(root, text="Start", fg="blue", height=4,
                  width=15)
button_1.bind("<Button-1>", run_app)
# button_1.bind("<Button-1>", bar)
button_1.pack()
root.mainloop()
