#! python3
'''
A script that scrapes pages from TheHub for startups and posts them to an excel spreadsheet
'''
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import bs4
import logging as log
import time
log.basicConfig(level=log.INFO)

key = 'AIzaSyAOTDlg3RXx6nXT-B-pbCOf1MY3-0W-7QU'
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('C:\\Users\\hi\\OneDrive\\Desktop\\main\\client_secret.json', scope)
client = gspread.authorize(creds)

titles = []
running = True

while running:
    def main():
        for i in range(188):
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

        sheet = client.open("TheHubData").sheet1

        index = []

        col = 1
        k = 1
        cellnumb = 1
        g = 0
        for i in range(len(titles)):
            time.sleep(1.5)
            newNumb = k
            index.append(titles[i])
            if titles[i] not in index:
                sheet.update_cell(newNumb, 4, titles[i])
                log.info("New Startup found: {}, Placed in cell number: {}".format((titles[i]), newNumb))
                k += 1
            else:
                sheet.update_cell(cellnumb+g, col, titles[i])
                log.info('Cell Number: {} updated with startup: {}'.format(g,titles[i]))
                g +=1
            if g % 200 == 0:
                g = 0
                col += 1
        log.info("Found: {} Startups from TheHub".format(str(len(titles))))
    main()
