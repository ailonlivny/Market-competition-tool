# now = datetime.now()
# dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
# workbook = xlsxwriter.Workbook('Competition.xlsx')
# worksheet = workbook.add_worksheet()
# file = open("test.txt", "r")
# worksheet.write('A1', f'Date :{dt_string}')
# worksheet.write('A2', 'JobTitle')
# worksheet.write('B2', 'location')
# worksheet.write('C2', 'JobID')
# worksheet.write('D2', 'Account')
# worksheet.write('E2', 'Sponsored  By Pando')
# worksheet.write('F2', 'Sponsored  By Others')
#
# row = 2
# column = 0
# count = 1
#
# for line in file:
#     if count % 4 == 0:
#         count += 1
#         continue
#     if count % 5 == 0:
#         count = 1
#         continue
#     if count % 3 == 0:
#         worksheet.write(row, column, line[8:])
#     else:
#         worksheet.write(row, column, line)
#
#     count += 1
#     column += 1
#     if column % 3 == 0:
#         column = 0
#         row += 1
#
# workbook.close()

import requests
import bs4
from bs4 import BeautifulSoup
import pandas as pd
import time

# def extract_job_title_from_result(soup):
#     jobs = []
#     for div in soup.find_all(name="div", attrs={"class": "row"}):
#         for a in div.find_all(name="a", attrs={"data-tn-element": "jobTitle"}):
#             jobs.append(a["title"])
#     return jobs
#
#
# def extract_company_from_result(soup):
#     companies = []
#     for div in soup.find_all(name="div", attrs={"class": "row"}):
#         company = div.find_all(name="span", attrs={"class": "company"})
#         if len(company) > 0:
#             for b in company:
#                 companies.append(b.text.strip())
#         else:
#             sec_try = div.find_all(name="span", attrs={"class": "result-link-source"})
#             for span in sec_try:
#                 companies.append(span.text.strip())
#         return companies
#
#
# URL = "https://www.indeed.com/jobs?q=data+scientist+%2420%2C000&l=New+York&start=10"
# page = requests.get(URL)
# soup = BeautifulSoup(page.text, "html.parser")
# # print(extract_job_title_from_result(soup))
# print(extract_company_from_result(soup))


# URL = "https://www.indeed.com/jobs?q=Warehouse+Team+Member&l=Whitestown+IN"
# page = requests.get(URL)
# soup = BeautifulSoup(page.text, "html.parser")
#
# results = soup.find_all('div', attrs={'data-tn-component': 'organicJob'})
#
# for x in results:
#     company = x.find('span', attrs={"class":"company"})
#     if company:
#         print('company:', company.text.strip())
#
#     job = x.find('a', attrs={'data-tn-element': "jobTitle"})
#     print('job:', job.text.strip())
#
#     print('url:', job.text.strip())
#     # salary = x.find('nobr')
#     # if salary:
#     #     print('salary:', salary.text.strip())
#
#     print('----------')


import requests
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import datetime

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

excel_row = 2
excel_col = 0
sponsored_by_others_col = 5
URL = "https://amazon.force.com/"
soup = BeautifulSoup(requests.get(URL).text, "html.parser")
results_from_amazon_force = soup.find_all('div', attrs={'class': 'col-xs-12 col-sm-9'})

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

        print(f"Row number:{excel_row}")
        excel_col = 0
        excel_row += 1

workbook.close()
