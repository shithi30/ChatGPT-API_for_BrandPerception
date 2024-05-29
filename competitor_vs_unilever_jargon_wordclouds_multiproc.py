#!/usr/bin/env python
# coding: utf-8

# import
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import pandas as pd
import duckdb
import multiprocessing
from tabulate import tabulate
from wordcloud import WordCloud, STOPWORDS
import time
import win32com.client
from pretty_html_table import build_table
import base64
import random

## Chaldal ##
def scrape_chaldal_process(keywords): 
    
    # accumulators
    brands = ['Boost', 'Clear Shampoo', 'Clear Men Shampoo', 'Clear Hijab', 'Simple', 'Pepsodent', 'Brylcreem', 'Bru', 'St. Ives', 'Horlicks', 'Sunsilk', 'Lux', 'Ponds', "Pond's", 'Closeup', 'Cif', 'Dove', 'Maltova', 'Domex', 'Clinic', 'Tresemme', 'Tresemmé', 'GlucoMax', 'Knorr', 'Glow & Lovely', 'Glow & Handsome', 'Wheel', 'Axe', 'Pureit', 'Lifebuoy', 'Surf Excel', 'Vaseline', 'Vim', 'Rin']
    df_acc_local = pd.DataFrame()
    lock = multiprocessing.Lock()

    # open window
    driver = webdriver.Chrome('chromedriver', options=[])
    driver.maximize_window()

    for k in keywords:
        # url
        url = "https://chaldal.com/search/" + k
        driver.get(url)
        
        # scroll
        SCROLL_PAUSE_TIME = 5
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(SCROLL_PAUSE_TIME)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height

        # soup
        soup_init = BeautifulSoup(driver.page_source, 'html.parser')
        soup = soup_init.find_all("div", attrs={"class": "imageWrapper"})
        
        # scrape
        skus = []
        quants = []
        prices = []
        prices_if_discounted = []
        for s in soup:
            # sku
            try: val = s.find("div", attrs={"class": "name"}).get_text()
            except: val = None
            skus.append(val)
            # quantity
            try: val = s.find("div", attrs={"class": "subText"}).get_text()
            except: val = None
            quants.append(val)
            # price
            try: val = float(s.find("div", attrs={"class": "price"}).get_text().split()[1].replace(',', ''))
            except: val = None
            prices.append(val)
            # discount
            try: val = float(s.find("div", attrs={"class": "discountedPrice"}).get_text().split()[1].replace(',', ''))
            except: val = None
            prices_if_discounted.append(val)
        
        # accumulate
        df = pd.DataFrame()
        df['sku'] = skus
        df['keyword'] = k
        df['quantity'] = quants
        df['price'] = prices
        df['price_if_discounted'] = prices_if_discounted
        
        # relevant data
        qry = '''
        select *
        from
            (select *, row_number() over() pos_in_pg
            from df
            ) tbl1 
        where replace(sku, ' ', '') ilike ''' + "'%" + k.replace(" ", "") + '''%';
        '''
        df = duckdb.query(qry).df()
        rel_idx = df['pos_in_pg'].tolist()
        len_rel_idx = len(rel_idx)
        
        # description
        bnr = 3
        try: driver.find_element(By.CLASS_NAME, "important-banner")
        except: bnr = 2
        descs = []
        report_times = []
        for i in range(0, len_rel_idx): 
            report_times.append(time.strftime('%Y-%m-%d %H:%M:%S'))
            descs.append("ERROR")
            try:
                # move
                path = '//*[@id="page"]/div/div[6]/section/div/div/div/div/section/div['+str(bnr)+']/div[2]/div['+str(rel_idx[i])+']/div/div'
                elem = driver.find_element(By.XPATH, path)
                mov = ActionChains(driver).move_to_element(elem)
                mov.perform()
                # details
                path1 = '//*[@id="page"]/div/div[6]/section/div/div/div/div/section/div['+str(bnr)+']/div[2]/div['+str(rel_idx[i])+']/div/div/div[5]/span/a'
                path2 = '//*[@id="page"]/div/div[6]/section/div/div/div/div/section/div['+str(bnr)+']/div[2]/div['+str(rel_idx[i])+']/div/div/div[6]/span/a'
                try: elem = driver.find_element(By.XPATH, path1)
                except: elem = driver.find_element(By.XPATH, path2)
                elem.click()
                # content
                path = '//*[@id="page"]/div/div[6]/section/div/div/div/div/section/div['+str(bnr)+']/div[2]/div['+str(rel_idx[i])+']/div/div[2]/div/div/article/section[2]/div[5]'
                elem = driver.find_element(By.XPATH, path)
                descs[i] = elem.text.replace("\n", " ")
                # close
                path = '//*[@id="page"]/div/div[6]/section/div/div/div/div/section/div['+str(bnr)+']/div[2]/div['+str(rel_idx[i])+']/div/div[2]/div/button'
                elem = driver.find_element(By.XPATH, path)
                elem.click()
            except: pass
            
        # conglomerate
        if_ubls = []
        skus = [skus[i-1] for i in rel_idx]
        for i in range(0, len_rel_idx):
            if_ubls.append(None)
            for b in brands:
                if b.lower() + ' ' in skus[i].lower():
                    if_ubls[i] = True
                    break     
        
        # accumulate
        df['if_unilever'] = if_ubls
        df['description'] = descs
        df['report_time'] = report_times
        df_acc_local = df_acc_local.append(df)
        
        # progress
        lock.acquire()
        print("Data fetched for keyword: " + k)
        lock.release()
        
    # close window
    driver.close()
    
    # return
    return df_acc_local
    
def scrape_chaldal(folder):
    
    # accumulators
    start_time = time.time()
    df_acc = pd.DataFrame()
    keywords = ['conditioner', 'handwash', 'bodywash', 'facewash', 'lotion', 'cream', 'toothpaste', 'dishwash', 'toilet clean', 'soup', 'shampoo', 'health drink', 'detergent', 'moisturizer', 'soap', 'petroleum jelly', 'hair oil', 'germ kill']
    process_count = 3
    keywords_chunks = [keywords[i::process_count] for i in range(process_count)]
    
    # processes
    pool = multiprocessing.Pool(process_count)
    dfs_acc = pool.map(scrape_chaldal_process, keywords_chunks)
    pool.close()
    pool.join()
    
    # csv
    for i in range(0, process_count): df_acc = df_acc.append(dfs_acc[i])
    filename = folder + r"\chaldal_unilever_keywords_data.csv"    
    df_acc.to_csv(filename, index=False)           
    
    # stats
    print("\nTotal SKUs found: " + str(df_acc.shape[0]))
    elapsed_time = str(round((time.time() - start_time) / 60.00, 2))
    print("Elapsed time to run processes (mins): " + elapsed_time)
    
    # return
    return df_acc

def send_viz(scraped_df, folder):
    
    # summary
    qry = '''
    select
        'Chaldal' platform,
        count(sku) "SKUs", 
        count(case when if_unilever=true then sku else null end) "UBL SKUs", 
        count(case when if_unilever=true then sku else null end)*1.00/count(sku) "SoS",
        count(case when if_unilever=true and pos_in_pg<11 then sku else null end)*1.00/(count(distinct keyword)*10) "top-10 SoS", 
        count(case when if_unilever=true and pos_in_pg<11 and length(description)=0 then sku else null end) "top-10 SoS missing description", 
        count(case when if_unilever=true and length(description)=0 then sku else null end) "UBL SKUs missing description", 
        count(case when description='ERROR' then sku else null end) "description errors"
    from scraped_df; 
    '''
    anls_df = duckdb.query(qry).df()
    print(tabulate(anls_df, headers='keys', tablefmt='psql', showindex=False))
    
    # email
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)

    # subject, recipients
    newmail.Subject = 'Chaldal Keyword SoS & Wordcloud'
    newmail.To = 'shithi.maitra@unilever.com'
    # newmail.To = 'mehedi.asif@unilever.com'
    # newmail.CC = 'mehedi.asif@unilever.com; zakeea.husain@unilever.com; rakaanjum.unilever@gmail.com; nazmussajid.ubl@gmail.com'
    
    # UBL cloud
    qry = ''' select description from scraped_df where if_unilever=true; '''
    df = duckdb.query(qry).df()
    text = " ".join(d for d in df['description'].fillna("").tolist())
    word_cloud = WordCloud(width=300, height=200, random_state=1, background_color="white", colormap="ocean", collocations=False, stopwords=STOPWORDS).generate(text)
    word_cloud.to_file(folder + r"\ubl_cloud.png")
    cloud1_html = f'<img src="cid: MyId1" style="border: 1px solid; padding: 5px; background-color: white; display: block" width="96%" height="100%"><figcaption><b>Fig-01:</b> UBL wordcloud</figcaption>'

    # non UBL cloud
    qry = ''' select description from scraped_df where if_unilever is null; '''
    df = duckdb.query(qry).df()
    text = " ".join(d for d in df['description'].fillna("").tolist())
    word_cloud = WordCloud(width=300, height=200, random_state=1, background_color="white", colormap="Set2", collocations=False, stopwords=STOPWORDS).generate(text)
    word_cloud.to_file(folder + r"\nonubl_cloud.png")
    cloud2_html = f'<img src="cid: MyId2" style="border: 1px solid; padding: 5px; background-color: white; display: block" width="96%" height="100%"><figcaption><b>Fig-02:</b> non-UBL wordcloud</figcaption>'
    
    # body
    newmail.HTMLbody = f'''
    Dear concern,<br><br>
    Today's <i>Chaldal</i> UBL share of search (SoS) and descriptions for selected keywords have been scraped. A brief statistics of the process is given below:
    ''' + build_table(anls_df, 'blue_light') + '''
    <table style="margin-left: auto; margin-right: auto">
        <tr>
            <td>''' + cloud1_html + '''</td>
            <td>''' + cloud2_html + '''</td>
        </tr>
    </table>
    <br>
    Two wordclouds based on UBL and non-UBL descriptions are presented above. Please find the full dataset attached. Note, this email was auto generated using <i>win32com</i>.<br><br>
    Thanks,<br>
    Shithi Maitra<br>
    Asst. Manager, Cust. Service Excellence<br>
    Unilever BD Ltd.<br>
    '''
    
    # attachment
    filename = folder + r"\chaldal_unilever_keywords_data.csv"
    newmail.Attachments.Add(filename)
    newmail.Attachments.Add(folder + r"\ubl_cloud.png").PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
    newmail.Attachments.Add(folder + r"\nonubl_cloud.png").PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId2")
    
    # send
    newmail.Send()

# call
if __name__ == "__main__":
    folder = r"C:\Users\Shithi.Maitra\Unilever Codes\Scraping Scripts\Chaldal Stocks"
    scraped_df = scrape_chaldal(folder)
    send_viz(scraped_df, folder)

# run
# "C:\Users\Shithi.Maitra\Downloads\chaldal_keywords_multi_process.py"
