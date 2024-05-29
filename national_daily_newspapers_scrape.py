#!/usr/bin/env python
# coding: utf-8

## import
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import duckdb
from googleapiclient.discovery import build
from google.oauth2 import service_account
import time
import win32com.client
from pretty_html_table import build_table
import random

## scrape

# newspapers
prompts = ["The Financial Express", "The Daily Star", "The Business Standard", "Prothom Alo"]
sources = ["thefinancialexpress.com", "thedailystar.net", "tbsnews.net", "prothomalo.com"]

# accumulators
start_time = time.time()
df_acc = pd.DataFrame()
timings = []

# preference
options = webdriver.ChromeOptions()
options.add_argument("ignore-certificate-errors")

# open window
driver = webdriver.Chrome(service=Service(), options=options)
driver.maximize_window()

# iterate
source_count = len(sources)
for j in range(0, source_count):
    
    # link
    link = "https://www.google.com/"
    driver.get(link)

    # search
    print("Fetching articles from: " + sources[j])
    elem = driver.find_element(By.CLASS_NAME, "gLFyf")
    elem.send_keys(prompts[j] + " consumer goods FMCG news\n")

    # scroll
    last_height = driver.execute_script("return document.body.scrollHeight")
    while(1):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height: break
        last_height = new_height

    # soup
    soup_init = BeautifulSoup(driver.page_source, "html.parser")
    soup = soup_init.find_all("div", attrs={"class": "N54PNb BToiNc cvP2Ce"})

    # scrape
    headline = []
    publish_date = []
    excerpt = []
    path = []
    url = []
    pos_in_search = []
    report_date = []
    news_count = len(soup)
    for i in range(0, news_count):

        # headline
        try: val = soup[i].find("h3", attrs={"class": "LC20lb MBeuO DKV0Md"}).get_text()
        except: val = None
        headline.append(val)

        # publication date
        try: val = soup[i].find("span", attrs={"class": "lhLbod gEBHYd"}).get_text()
        except: val = None
        publish_date.append(val)

        # excerpt
        val = soup[i].find("div", attrs={"class": "kb0PBd cvP2Ce"}).get_text()
        val = val.split(publish_date[i])[1] if publish_date[i] is not None else val
        excerpt.append(val)

        # path
        try: val = soup[i].find("cite", attrs={"class": "qLRx3b tjvcx GvPZzd cHaqb"}).get_text()
        except: val = None
        path.append(val)
        
        # url
        try: val = soup[i].find("a", attrs={"jsname": "UWckNb"})["href"]
        except: val = None
        url.append(val)

        # position
        pos_in_search.append(i + 1)
        
        # timing
        timing = str(time.strftime('%Y-%m-%d %H:%M:%S'))
        report_date.append(timing)

    # accumulate 
    df = pd.DataFrame()
    df['headline'] = headline
    df['publish_date'] = [p[0:-3] if p is not None else p for p in publish_date]
    df['excerpt'] = excerpt
    df['path'] = path
    df['url'] = url
    df['position'] = pos_in_search
    df['newspaper'] = sources[j]
    df['report_date'] = report_date
    df = duckdb.query('''select * from df where path like '%''' + sources[j] +  '''%' ''').df()
    df_acc = df_acc.append(df, ignore_index=True)
    timings.append(timing)
    
# close window
driver.close()

## GSheet

# credentials
SERVICE_ACCOUNT_FILE = "read-write-to-gsheet-apis-1-04f16c652b1e.json"
SAMPLE_SPREADSHEET_ID = "1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# APIs
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build("sheets", "v4", credentials=creds)
sheet = service.spreadsheets()

# extract
values = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='News!A1:I').execute().get('values', [])
df_acc_prev = pd.DataFrame(values[1:], columns = values[0])

# transform
qry = '''
-- old
select headline, publish_date, excerpt, path, url, position, newspaper, 0 if_new, report_date
from df_acc_prev
union all
-- new
select 
    headline, publish_date, excerpt, path, url, position, newspaper, 
    case 
        when publish_date like '%23' then 1
        when publish_date like '%২৩' then 1
        when publish_date like '%ago' then 1
        when publish_date like '%আগে' then 1
        else 0
    end if_new, 
    report_date
from df_acc
where url not in(select url from df_acc_prev)
'''
df_acc_pres = duckdb.query(qry).df()

# load
sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='News').execute()
sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='News!A1', valueInputOption='USER_ENTERED', body={'values': [df_acc_pres.columns.values.tolist()] + df_acc_pres.fillna('').values.tolist()}).execute()

## novelty

# new articles
df_acc_new = duckdb.query('''select * from df_acc_pres where if_new=1''').df()
new_heads = df_acc_new['headline'].tolist()
new_links = df_acc_new['url'].tolist()
new_dates = df_acc_new['publish_date'].tolist()
new_len = df_acc_new.shape[0]

# latest pull
timing_df = pd.DataFrame()
timing_df['source'] = sources
timing_df['timing'] = timings

# store
if new_len > 0:
    with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/newspaper_fmcg_scrapings.xlsx") as writer:
        df_acc_pres.to_excel(writer, sheet_name="All Results", index=False)

## email

# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# summary
qry = '''
select newspaper, total_articles, new_articles, last_report_time, latest_report_time
from 
    (select newspaper, count(url) total_articles, count(case when if_new=1 then url else null end) new_articles
    from df_acc_pres
    group by 1
    ) tbl1 

    inner join

    (select newspaper, max(report_date) last_report_time
    from df_acc_prev
    group by 1
    ) tbl2 using(newspaper)
    
    inner join

    (select source newspaper, timing latest_report_time
    from timing_df
    ) tbl3 using(newspaper)
'''
summ_df = duckdb.query(qry).df()

# subject, recipients
newmail.Subject = 'Newspaper Scrapings - FMCG'
# newmail.To = 'shithi.maitra@unilever.com'
newmail.To = 'zoya.rashid@unilever.com'
newmail.CC = 'mehedi.asif@unilever.com; samsuddoha.nayeem@unilever.com; sudipta.saha@unilever.com; asif.rezwan@unilever.com'

# body
newmail.HTMLbody = f'''
Dear concern,<br><br>
Please find recent FMCG articles from the popular English dailies attached. New articles are reported on <a href="https://teams.microsoft.com/l/channel/19%3ae8f5c9a9e7374b51840112d6280374af%40thread.tacv2/FMCG%2520News?groupId=1b8eee70-e11c-419e-966f-d830a968c87a&tenantId=f66fae02-5d36-495b-bfe0-78a6ff9f8e6e">Teams FMCG News</a>. Here are statistics from the latest pull: 
''' + build_table(summ_df, random.choice(['green_light', 'red_light', 'blue_light', 'grey_light', 'orange_light']), font_size='12px', text_align='left') + '''
More newspapers, online news portals or even Bangla dailies can be incorporated on demand. This is an auto-generated email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, CSE<br>
Unilever BD Ltd.<br>
'''

# attachment(s) 
folder = "C:/Users/Shithi.Maitra/Downloads/"
filename = folder + "newspaper_fmcg_scrapings.xlsx"
newmail.Attachments.Add(filename)

# send
if new_len > 0: newmail.Send()

## MSTeams

# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# report
new = "⚠ The following " + str(new_len) + " article(s) are newly found, as of " + timing
for i in range(0, new_len): new = new + '''<br>&nbsp;&nbsp;&nbsp;• <a href="''' + new_links[i] + '''">''' + new_heads[i] + '''</a> [''' + new_dates[i] + ''']''' 
    
# Teams
newmail.Subject = "New FMCG Articles!"
newmail.To = "FMCG News - Auto Monitoring <062c1c6b.Unilever.onmicrosoft.com@emea.teams.ms>"
newmail.HTMLbody = new + "<br><br>"
if new_len > 0: newmail.Send()

## stats
display(df_acc_pres.head())
print("Articles in result: " + str(df_acc_pres.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))
