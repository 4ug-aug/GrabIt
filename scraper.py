#! /usr/env/bin python3
'''
A script that scrapes pages from TheHub for startups and posts them to an excel spreadsheet
'''
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import pandas
import bs4
import logging as log
import openpyxl
import time
log.basicConfig(level=log.INFO)

key = 'AIzaSyAOTDlg3RXx6nXT-B-pbCOf1MY3-0W-7QU'
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

titles = []

for i in range(15):
    # time.sleep(0.5)
    pagenr = 1+i
    hub = "https://thehub.io/startups?countryCodes=DK&page={}".format(str(pagenr))

    try:
        r = requests.get(hub)
        log.info('request sent')
    except requests.exceptions.RequestException as e:
        log.info(e)
        raise ConnectionError

    content = bs4.BeautifulSoup(r.text, 'html.parser')
    titlesRaw = content.select('.card-title')
    log.info(len(titlesRaw))
    # log.info(titlesRaw)

    for i in range(len(titlesRaw)):
        titles.append(titlesRaw[i].getText())

# log.info(titles)
# log.info(len(titles))

'TODO: Insert in excel via openpyxl'

sheet = client.open("TheHubData").sheet1
# wb = openpyxl.load_workbook('TheHubStartups.xlsx')
# ark = wb['Ark1']

for i in range(len(titles)):
    time.sleep(1.5)
    cellnumb = i+1
    sheet.update_cell(cellnumb, 1, titles[i])
    log.info('Cell Number: {} updated with startup: {}'.format(cellnumb,titles[i]))
log.info("Found: {} Startups from TheHub".format(str(len(titles))))