#!/usr/bin/env python
# coding: utf-8

# In[2]:


## Daraz Comments

# import
import pandas as pd
import duckdb
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from bs4 import BeautifulSoup
import time

# accumulators
start_time = time.time()
df_acc = pd.DataFrame()
df_cmt = pd.DataFrame()

# preferences
option = Options()
option.add_argument("--disable-notifications")

# open window
driver = webdriver.Chrome("chromedriver", options = option)
driver.maximize_window()

# implicit wait
driver.implicitly_wait(5)

# scroll down
def scroll_down():
    scroll_pause_time = 5
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(scroll_pause_time)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height: break
        last_height = new_height
        
# scroll smooth
def scroll_smooth():
    y = 500
    for timer in range(0, 10):
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, " + str(y) + ")")
        y = y + 500
        
# portfolio
brands = ['Lifebuoy', 'Lux', 'Dove']
def if_right_brand(sku):
    for b in brands:
        if b.lower() in sku.lower(): return b
    return None
        
# link
pg = 0
while(1): 
    pg = pg + 1
    link = "https://www.daraz.com.bd/unilever-bangladesh/?from=wangpu&lang=en&langFlag=en&page=" + str(pg) + "&pageTypeId=2&q=All-Products"
    driver.get(link)

    # scroll
    scroll_down()

    # soup
    soup_init = BeautifulSoup(driver.page_source, "html.parser")
    soup_init = soup_init.find_all("div", attrs = {"class": "gridItem--Yd0sa"})

    # page
    sku_count = len(soup_init)
    if sku_count == 0: break 
    print("Scraping from page: " + str(pg))
    
    # scrape
    sku = []
    brand = []
    stars_avg = []
    stars_5 = []
    stars_4 = []
    stars_3 = []
    stars_2 = []
    stars_1 = []
    top_comments = []
    comment_ratings = []
    comment_dates = []
    queries = []
    query_dates = []
    report_time = []
    df_cmt_brands = []
    df_cmt_reviews = []
    df_cmt_ratings = []
    for i in range(0, sku_count):
        
        # validity
        basepack = soup_init[i].find("div", attrs = {"class": "title--wFj93"}).get_text()
        b = if_right_brand(basepack)
        if b is None: continue
        
        # sku, brand
        sku.append(basepack)
        brand.append(b)
        
        # click
        elems = driver.find_elements(By.CLASS_NAME, "image--WOyuZ ")
        ActionChains(driver).move_to_element(elems[i]).click().perform()
        
        # scroll
        scroll_smooth()
        
        # soup
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # top comments with ratings and dates
        soup1 = soup.find_all("div", attrs = {"class": "item"})
        # dates
        date = ""
        for s in soup1: 
            val = s.find("span", attrs = {"class": "title right"}).get_text()
            date = date + "\n- " + val
            df_cmt_brands.append(b)
        comment_dates.append(date[1:])
        # ratings
        rating = ""
        for s in soup1: 
            val = str(len(s.find_all("img", attrs = {"class": "star", "src": "//laz-img-cdn.alicdn.com/tfs/TB19ZvEgfDH8KJjy1XcXXcpdXXa-64-64.png"})))
            df_cmt_ratings.append(val)
            rating = rating + "\n- " + val
        comment_ratings.append(rating[1:])
        # reviews
        review = ""
        for s in soup1:
            val = s.find("div", attrs = {"class": "item-content"}).find("div", attrs={"class": "content"}).get_text().replace("\n", " ")
            df_cmt_reviews.append(val)
            review = review + "\n- " + val
        top_comments.append(review[1:])
        
        # ratings
        soup2 = soup.find_all("span", attrs = {"class": "percent"})
        for j in range(0, 5):    
            if j == 0: stars_5.append(soup2[j].get_text())
            if j == 1: stars_4.append(soup2[j].get_text())
            if j == 2: stars_3.append(soup2[j].get_text())
            if j == 3: stars_2.append(soup2[j].get_text())
            if j == 4: stars_1.append(soup2[j].get_text())
        soup3 = soup.find_all("span", attrs = {"class": "score-average"})
        for s in soup3: stars_avg.append(s.get_text())
            
        # queries with dates
        soup4 = soup.find_all("li", attrs = {"class": "qna-item"})
        # dates
        date = ""
        for s in soup4: date = date + "\n- " + s.find("div", attrs = {"class": "qna-meta"}).get_text().split(" - ")[1]
        query_dates.append(date[1:])
        # queries
        query = ""
        for s in soup4: query = query + "\n- " + s.find("div", attrs = {"class": "qna-content"}).get_text()
        queries.append(query[1:])
        
        # record
        report_time.append(time.strftime("%d-%b-%y, %I:%M %p"))
        
        # back
        driver.back()
        
    # accumulate
    df = pd.DataFrame()
    df["sku"] = sku
    df["brand"] = brand
    df["stars_avg"] = stars_avg
    df["stars_5"] = stars_5
    df["stars_4"] = stars_4
    df["stars_3"] = stars_3
    df["stars_2"] = stars_2
    df["stars_1"] = stars_1
    df["top_comments"] = top_comments
    df["comment_ratings"] = comment_ratings
    df["comment_dates"] = comment_dates
    df["queries"] = queries
    df["query_dates"] = query_dates
    df["report_time"] = report_time
    df_acc = df_acc.append(df)
    df_acc = duckdb.query('''select distinct * from df_acc''').df()
    df = pd.DataFrame()
    df["review"] = df_cmt_reviews
    df["rating"] = df_cmt_ratings
    df["brand"] = df_cmt_brands
    df_cmt = df_cmt.append(df)
    df_cmt = duckdb.query('''select distinct * from df_cmt''').df()

# close window
driver.close()

# stats
display(df_acc.head(5))
print("Listings in result: " + str(df_acc.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))


# In[6]:


## ChatGPT Responses

# certificate: https://community.openai.com/t/ssl-certificate-verify-failed/32442/38

# import
import openai
import os
import re

# setup
os.environ['REQUESTS_CA_BUNDLE'] = 'Zscaler Root CA.crt'
openai.api_key = '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'

# prompt
def prompt_gpt(message):
    response = openai.Completion.create(engine = 'text-davinci-003', prompt = message, max_tokens = 300, temperature = 0.5, n = 1, stop = None, frequency_penalty = 0, presence_penalty = 0)
    return response.choices[0].text.strip()

# subpar reviews
qry = '''
select brand, count(*) subpar_reviews, left(concat('- ', string_agg(review order by random(), '\n- ')), 2000) reviews
from df_cmt
where rating < 5
group by 1
'''
actions_df = duckdb.query(qry).df()
reviews = actions_df['reviews'].tolist()

# ChatGPT recommendations
gpt_actions = []
for r in reviews: 
    prompt = "Propose 5 actions to improve business from the following reviews left by customers of an eCommerce platform. Assign serial numbers to the actions and limit each to 15 words.\n\n" + r
    action = "1. " + prompt_gpt(prompt).split("1. ")[-1]
    print(action)
    gpt_actions.append(action)
actions_df['chatgpt_recommendations'] = gpt_actions
actions_df = actions_df[['brand', 'subpar_reviews', 'chatgpt_recommendations']]


# In[7]:


## GSheet

# import
from googleapiclient.discovery import build
from google.oauth2 import service_account

# credentials
SERVICE_ACCOUNT_FILE = "read-write-to-gsheet-apis-1-04f16c652b1e.json"
SAMPLE_SPREADSHEET_ID = "1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# APIs
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes = SCOPES)
service = build("sheets", "v4", credentials = creds)
sheet = service.spreadsheets()

# update
resultClear = sheet.values().clear(spreadsheetId = SAMPLE_SPREADSHEET_ID, range = "Reviews").execute()
request = sheet.values().update(spreadsheetId = SAMPLE_SPREADSHEET_ID, range = "'Reviews'!A1", valueInputOption = "USER_ENTERED", body = {"values": [df_acc.columns.values.tolist()] + df_acc.fillna("").values.tolist()}).execute()
request = sheet.values().update(spreadsheetId = SAMPLE_SPREADSHEET_ID, range = "'Reviews'!P1", valueInputOption = "USER_ENTERED", body = {"values": [actions_df.columns.values.tolist()] + actions_df.fillna("").values.tolist()}).execute()


# In[8]:


## Donut Charts

# import
import matplotlib.pyplot as plt
from colour import Color
import numpy


# In[9]:


## Donut - Numbers

# data
sizes = numpy.array([
    numpy.log(pd.to_numeric(df_acc['stars_5']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_4']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_3']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_2']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_1']).sum())
])
labels = ['5 stars', '4 stars', '3 stars', '2 stars', '1 star']
colors = list(Color("green").range_to(Color("red"), 5))
colors = [str(c) for c in colors]
explode = (0.01, 0.01, 0.01, 0.01, 0.01)

# numeric
def absolute_value(val):
    a = round(numpy.exp((val/100)*sizes.sum()))
    return a

# chart
plt.pie(x=sizes, 
        colors=colors, 
        radius=1, 
        autopct=absolute_value, 
        pctdistance=0.78, 
        explode=explode,
        textprops={'fontsize': 7, 'color': 'white', 'weight': 'bold'}
       )
centre_circle = plt.Circle((0, 0), 0.6, fc='white')
leg = plt.legend(labels, loc="center", prop={'size': 7})
leg.get_frame().set_alpha(0)

# resize
fig = plt.gcf()
fig.set_size_inches(3.25, 3.25)
fig.gca().add_artist(centre_circle)

# save
plt.savefig('review_donut_num.jpeg', bbox_inches='tight', dpi=300)

# show
plt.show()


# In[10]:


## Donut - Percent

# data
sizes = numpy.array([
    numpy.log(pd.to_numeric(df_acc['stars_5']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_4']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_3']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_2']).sum()), 
    numpy.log(pd.to_numeric(df_acc['stars_1']).sum())
])
labels = ['5 stars', '4 stars', '3 stars', '2 stars', '1 star']
colors = list(Color("green").range_to(Color("red"), 5))
colors = [str(c) for c in colors]
explode = (0.01, 0.01, 0.01, 0.01, 0.01)

# pct
def absolute_value(val):
    a = round(numpy.exp((val/100)*sizes.sum()))/sum([round(numpy.exp(s)) for s in sizes])    
    return "{0:.2%}".format(a)

# chart
plt.pie(x=sizes, 
        colors=colors, 
        radius=1, 
        autopct=absolute_value, 
        pctdistance=0.78, 
        explode=explode,
        textprops={'fontsize': 7, 'color': 'white', 'weight': 'bold'}
       )
centre_circle = plt.Circle((0, 0), 0.6, fc='white')
leg = plt.legend(labels, loc="center", prop={'size': 7})
leg.get_frame().set_alpha(0)

# resize
fig = plt.gcf()
fig.set_size_inches(3.25, 3.25)
fig.gca().add_artist(centre_circle)

# save
plt.savefig('review_donut_pct.jpeg', bbox_inches='tight', dpi=300)

# show
plt.show()


# In[13]:


## Auto Email

# import
import win32com.client

# object
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = "Consumers' Voice | ChatGPT's Actions"
# newmail.To = 'shithi.maitra@unilever.com'
newmail.To = 'mehedi.asif@unilever.com'
newmail.CC = 'asifur.rahman@unilever.com; hasnain-karim.rabib@unilever.com; asif.rezwan@unilever.com'

# UBL cloud
donut1_html = f'<img src="cid: MyId1" style="border: 0px solid; padding: 0px; background-color: white; display: block" width="250" height="250"><figcaption><center><b>Fig-01:</b> Distribution of ratings (pct)</center></figcaption>'

# non UBL cloud
donut2_html = f'<img src="cid: MyId2" style="border: 0px solid; padding: 0px; background-color: white; display: block" width="250" height="250"><figcaption><center><b>Fig-02:</b> Distribution of ratings (num)</center></figcaption>'

# body
df_acc_qry = duckdb.query('''select * from df_acc where queries!='' ''').df()
newmail.HTMLbody = f'''
Dear concern,<br><br>
Analyzing end-consumers' comments/queries may help design better Customer Service. Hence, reviews and ratings from <a href="https://www.daraz.com.bd/shop/unilever-bangladesh/">Daraz UBL</a> platform have been examined, as found in the attachment.
<br>
<table style="margin-left: auto; margin-right: auto">
    <tr>
        <td>''' + donut1_html + '''</td>
        <td>''' + donut2_html + '''</td>
    </tr>
</table>
<br>
Given below are examples from queries on <i>''' + df_acc_qry['sku'].tolist()[0] + '''</i>:<br>''' + df_acc_qry['queries'].tolist()[0].replace("\n", "<br>") + '''
<br><br>
Reviews left by customers have been prompted to <b>ChatGPT</b> via integrating its API to <i>engine: text-davinci-003</i>, for recommendations. Given below are actions proposed for <i>''' + actions_df["brand"].tolist()[0] + '''</i>:<br><br>''' + actions_df["chatgpt_recommendations"].tolist()[0].replace("\n", "<br>") + '''
<br><br>
For the time being, data has been scraped for the following portfolio: <i>''' + ", ".join(brands) + '''</i>. I convey gratitude to <a href="mailto: asifur.rahman@unilever.com">@Rahman, Asifur</a> Bhaiya for the idea. This is an auto email via <i>win32com</i>.
<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, CSE<br>
Unilever BD Ltd.<br>
'''

# attachment
folder = r"C:\Users\Shithi.Maitra\Unilever Codes\Scraping Scripts"
filename = folder + r"\Eagle Eye - Reviews & Recommends.xlsx"
newmail.Attachments.Add(filename)
newmail.Attachments.Add(folder + r"\review_donut_num.jpeg").PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId2")
newmail.Attachments.Add(folder + r"\review_donut_pct.jpeg").PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

# send
newmail.Send()


# In[ ]:




