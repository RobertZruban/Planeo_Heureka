import pprint
import random
import requests
import urllib.request
from selenium import webdriver
import time
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.keys import Keys  
from bs4 import BeautifulSoup
import pandas as pd
import pyodbc
import sqlalchemy
import pandas as pd
from datetime import datetime
import win32com.client
from collections import Counter
from datetime import date
import os
import math

########################################################################################################################################
###Browser opening and options ######
browser = webdriver.Chrome(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\chromedrive\chromedriver.exe') 
########################################################################################################################################
######websites#####
starting_url = 'https://www.planeo.sk/katalog/3000003-akcie.html?page='
planeo_website = "https://www.planeo.sk"   
########################################################################################################################################
browser.get(starting_url) 
browser.maximize_window()
time.sleep(2)
e = browser.find_element("id", "CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll")
e.click()
df = pd.DataFrame(columns = ["id","Popis","Nazov","cena","bezna_cena","link"])

html_source = browser.page_source  
soup = BeautifulSoup(html_source,'html.parser')

all_links = []
count = 0
today = date.today()
today = date.today().strftime("%Y-%m-%d")

try:
    for x in range(1,50):
        url = starting_url + str(x)
        browser.get(url) 
        html_source = browser.page_source  
        soup = BeautifulSoup(html_source,'html.parser')
        links_per_page = [(a['href']) for a in soup.find_all('div',{'id' :'product-list'})[0].find_all('a', href=True) if "katalog" in a['href']]
        links_per_page = list(set(links_per_page))
        [all_links.append(x) for x in links_per_page]
except:
    pass

for x in all_links:
    try:
        price_2 = ""
        price_3 = ""
        bezna_cena = ""
        price = ""
        pokus_main =""
        count += 1
        print(f"Remaining websites to crawl: {len(all_links)-count}")
        len(all_links)
        url =  planeo_website + str(x)
        browser.get(url) 
        html_source = browser.page_source  
        soup = BeautifulSoup(html_source,'html.parser')
        
        if len(soup.find_all('div', {'class' : 'price'})) == 1:
            price = soup.find_all('div',{'class' :'price'})[0].get_text()
            price = price[0:price.find('€')].replace('\n','').replace('\xa0','').replace('€Cena s DPH','').replace(',','.')
        elif len(soup.find_all('div', {'class' : 'price'})) == 2:
            price = soup.find_all('div',{'class' :'price'})[1].get_text()
            price = price[0:price.find('€')].replace('\n','').replace('\xa0','').replace('€Cena s DPH','').replace(',','.')
        else:
            price = ""
        
        try:
            pokus_main = soup.find_all('div', {'class' : 'stamp-icon no-price custom-stamp'})[0]
            price_2 = str(pokus_main)[str(pokus_main).find('_AKCNACENA')+1:str(pokus_main).find('png')-1]
            price_2 = price_2.replace('AKCNACENA_','')
            price_3 = str(pokus_main)[str(pokus_main).find('_akcna_cena_')+1:str(pokus_main).find('png')-1]
            price_3 = price_3.replace('akcna_cena_','')
        except:
            pass
        
        id_ = soup.find_all('div',{'class' :'posa r0 t0 fz90p c-text'})[0].get_text()
        id_ = id_[id_.find(':')+2:]
        
        try: 
            bezna_cena = soup.find_all('dd',{'class' :'moc'})[0].get_text()
            bezna_cena = bezna_cena[0:bezna_cena.find('€')]
            bezna_cena = bezna_cena.replace('\n','').replace('\xa0','').replace(',','.')
        except:
            pass
        
        try:
            if float(price_2) > 0:
                price = price_2
                bezna_cena = soup.find_all('div',{'class' :'price'})[1].get_text()
                bezna_cena = bezna_cena[0:price.find('€')].replace('\n','').replace('\xa0','').replace('€Cena s DPH','').replace(',','.')
        except:
            pass
        
        try:
            if float(price_3) > 0:
                price = price_3
                bezna_cena = soup.find_all('div',{'class' :'price'})[1].get_text()
                bezna_cena = bezna_cena[0:price.find('€')].replace('\n','').replace('\xa0','').replace('€Cena s DPH','').replace(',','.')
        except:
            pass

        name_1 = soup.find_all('span', {'class' : 'type'})[1].get_text()
        name_2 = soup.find_all('h1')[0].get_text().replace('\n','').replace('\t','')
        
                      
        dictionary = {"id" : id_,"Popis":name_1,"Nazov":name_2,"cena" : price ,"bezna_cena" :bezna_cena, "link" : url}
        df = df.append(dictionary, ignore_index=True, sort=False)
    except:
        print('error')
        dictionary = {"id" : "-","Popis":"-","Nazov":"-","cena" : "-" ,"bezna_cena" : "-","link" : url}
        df = df.append(dictionary, ignore_index=True, sort=False)
       
browser.quit()

df['cena'] = df['cena'].apply(pd.to_numeric, errors='coerce')
df['bezna_cena'] = df['bezna_cena'].apply(pd.to_numeric, errors='coerce')
df.loc[df['bezna_cena'].isnull(),'bezna_cena'] = df['cena']
df['zlava'] = round((df['cena'] / df['bezna_cena']) - 1,2)
df = df.sort_values(by='zlava', ascending=True, na_position='last')
df['Date'] = [today] * len(df['zlava'])
df = df.fillna(value=0)




conn = pyodbc.connect(
"Driver={SQL Server};"
"Server=DESKTOP-F86F289;"
"Database =Planeo;"
"Trusted_Connection=yes;")
cursor = conn.cursor()

cursor.execute('''
                Delete from Planeo.dbo.Planeo WHERE Date in (SELECT Convert(DateTime, DATEDIFF(DAY, 0, GETDATE())))

               ''')

conn.commit()


for row in df.itertuples():
    cursor.execute('''
        INSERT INTO Planeo.dbo.Planeo (id, Popis, Nazov,cena, bezna_cena,link,zlava,Date)
        VALUES (?,?,?,?,?,?,?,?)
        ''',
        row.id, 
        row.Popis,          
        row.Nazov,
        row.cena,
        row.bezna_cena,
        row.link,
        row.zlava,
        today,       
        )
conn.commit()


str(today)
df.to_excel(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_' + str(today) +'.xlsx')
time.sleep(10)
######################################################################################
#-----------------#importing#-----------------#
######################################################################################
df = pd.read_excel(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_' + str(today) + '.xlsx')
df = df.sort_values(by='zlava', ascending=True, na_position='last')
df = df[df['cena'] > 10]
df = df.reset_index(drop=True)
df.drop(['Unnamed: 0', 'id','Date'], axis=1,inplace=True )
df.rename(columns={"zlava": "zlava_in_perc"}, inplace=True)
df['zlava_in_perc'] = round(df['zlava_in_perc']*100)



html1 = df.head(20).to_html()
os.startfile("outlook")
ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
list_of_email = ['zrubanrobert@gmail.com']

# Function to get unique values
receiver = 'zrubanrobert@gmail.com'
#cc = 'puskjana@gmail.com'

newmail.Subject = f'Planeo Zlavy: {today}' 
newmail.To = receiver
#newmail.CC = cc
newmail.HTMLBody = html1
#str(today)
newmail.Attachments.Add(Source=r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_' + str(today)  + '.xlsx')
newmail.Display()
newmail.Send()

#for x in range(0 ,10):
     #jebo = (x," : ",  df['Popis'][x], ':' , {df['Nazov'][x]}, ' je zlacnený o ', round(df['zlava'][x]*100,3),'%' , 'z pôvodnej ceny', df['bezna_cena'][x], 'eur', ' na ', df['cena'][x],'eur','. Link na tento produkt je ',  {df['link'][x]} ) 
     #jebo = str(jebo)
     #jebo = jebo.replace(',','').replace("'",'').replace("(",'').replace(")",'')
     #mail.append(jebo)
    #jebo = (df['Popis'][x], ' so znacenim ' , df['Nazov'][x], ' je zlacneny o ', round(df['zlava'][x]*100,3),'%' , 'z beznej ceny', df['bezna_cena'][x], 'eur', ' na cenu ', df['cena'][x],'eur',' Link na tento produkt je ',  df['link'][x] )
    #mail.append(jebo)
    #print({df['Popis'][x]}, ' so znacenim ' , {df['Nazov'][x]}, ' je zlacneny o ', round(df['zlava'][x]*100,3),'%' , 'z beznej ceny', {df['bezna_cena'][x]}, 'eur', ' na cenu ', {df['cena'][x]},'eur',' Link na tento produkt je ',  {df['link'][x] })
#[print({df['Popis'][x]}, ' so znacenim ' , {df['Nazov'][x]}, ' je zlacneny o ', round(df['zlava'][x]*100,3),'%' , 'z beznej ceny', {df['bezna_cena'][x]}, 'eur', ' na cenu ', {df['cena'][x]},'eur',' Link na tento produkt je ',  {df['link'][x] }, '\n') for x in range(0 ,10) ]   
#final_text = "Posielam najviac zlacnené produkty na dnes zo stránky www.Planeo.sk \n\n" + mail[0] + "\n" + mail[1] + "\n" + mail[2] + "\n" + mail[3] + "\n" + mail[4] + "\n" + mail[5]+  "\n"  + mail[6] + "\n" + mail[7] + "\n"  + mail[8] + "\n" + mail[9] 

#####Imports######
import pprint
import random
import requests
import urllib.request
from selenium import webdriver
import time
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.keys import Keys  
from bs4 import BeautifulSoup
import pandas as pd
import pyodbc
import sqlalchemy
import pandas as pd
from datetime import datetime
import win32com.client
from collections import Counter
from datetime import date
import win32com.client as win32
from datetime import datetime, timedelta
conn = pyodbc.connect(
"Driver={SQL Server};"
"Server=DESKTOP-F86F289;"
"Database =Planeo;"
"Trusted_Connection=yes;")

df = pd.DataFrame()
df = pd.read_sql("SELECT * FROM Planeo.dbo.Planeo WHERE id IN (SELECT a.id FROM Planeo.dbo.Planeo a JOIN Planeo.dbo.Planeo b on b.id = a.id AND b.cena <> a.cena)", conn)

#today = date.today().strftime("%Y-%m-%d")
yesterday = datetime.now() - timedelta(1)
yesterday = datetime.strftime(yesterday, '%Y-%m-%d')
df = df[(df.Date == today) | (df.Date == yesterday)]
df = df.sort_values(by=['id', 'Date'], ascending=False)
df['change_from_yesterday_%'] = round((df.groupby(['id'])['cena'].pct_change(-1))*100,2)
df['zlava'] = round(df['zlava'] *100,2)

def date(row):  
    if row['Date'].strftime("%Y-%m-%d") == today:
        return 'Today'
    elif row['Date'].strftime("%Y-%m-%d") == yesterday:
        return 'Yesterday'    
    else:
        return 'Older than yesterday'

df['Date'] = df.apply(lambda row: date(row), axis=1)
df.rename(columns={"zlava": "zlava_in_perc", 'Popis' : 'item_name' }, inplace=True)
df.drop(['id'], axis=1,inplace=True )
df = df.reset_index(drop=True)
df = df[df['cena'] > 10]
df = df.sort_values(['Date','change_from_yesterday_%'], ascending = [True,True])
df = df[['item_name','Nazov','Date','cena','bezna_cena', 'zlava_in_perc','change_from_yesterday_%','link']]

receiver = 'zrubanrobert@gmail.com'
list_of_email = ['zrubanrobert@gmail.com']
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = receiver
mail.Subject = f'Planeo zlavnene ceny ktore sa zmenili oproti vcerajsku' 

html1 = df.to_html()
mail.HTMLBody = html1
mail.Display()
mail.Send()

#cc = 'puskjana@gmail.com'

from datetime import datetime
from selenium.common.exceptions import TimeoutException
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.keys import Keys  
from bs4 import BeautifulSoup
import pandas as pd
import pprint
pp = pprint.PrettyPrinter(indent=4)
import random
import requests
import urllib.request
import time
from bs4 import BeautifulSoup
import pyodbc
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.by import By
conn = pyodbc.connect(
"Driver={SQL Server};"
"Server=DESKTOP-F86F289;"
"Database =Planeo;"
"Trusted_Connection=yes;")
cursor = conn.cursor()
df = pd.DataFrame()
nazov_array = pd.read_sql("SELECT * FROM Planeo.dbo.Planeo", conn)
nazov_array = nazov_array['Nazov']
browser = webdriver.Chrome(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\chromedrive\chromedriver.exe')
url_main = 'https://www.heureka.sk/'
browser.get(url_main)
i = -1
for x in (list(nazov_array.unique())):
    try:
        print(len(list(nazov_array.unique()))-i)
        i +=1
        browser.get(url_main)
        #time.sleep(2)
        browser.find_element(By.CLASS_NAME, "c-search__input").send_keys(x)
        browser.find_element(By.CLASS_NAME, "c-search__input").send_keys(Keys.RETURN)
        browser.find_element(By.CLASS_NAME, "c-product__link").send_keys(Keys.RETURN)
        url = browser.current_url
        html_source = browser.page_source  
        soup = BeautifulSoup(html_source,'html.parser')
        alts = [x['alt'] for x in soup.find_all('img', {'class' : "c-offer__shop-logo e-image-with-fallback"}, alt = True)]
        price = [x.get_text().replace('\xa0','').replace('€','').replace(',','.') for x in soup.find_all('span', {'class' : "c-offer__price u-extra-bold u-delta"})]
        item_name = [x] * len(price)
        item_name_downloaded = soup.find_all('h1',{'class' :'e-heading c-product-info__name u-color-grey-700 u-bold u-gamma'})[0].get_text()
        item_name_downloaded = [item_name_downloaded] * len(price)
        df2 = pd.DataFrame(data = [item_name,item_name_downloaded,alts,price])
        df2 = df2.T
        df = pd.concat([df, df2])
    except:
        pass
df = df.reset_index(drop=True)
df.rename(columns = {0:'item_name',1:'item_name_downloaded',2:'alts',3:'price'}, inplace = True)
df['price'].apply(pd.to_numeric, errors='coerce')
#df.to_excel(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_Heureka.xlsx')

cursor.execute('''
                DELETE FROM Planeo.dbo.Heureka 
               ''')

conn.commit()


for row in df.itertuples():
    cursor.execute('''
        INSERT INTO Planeo.dbo.Heureka (item_name, item_name_downloaded, alts,price)
        VALUES (?,?,?,?)
        ''',
        row.item_name, 
        row.item_name_downloaded,
        row.alts,
        row.price
        )
conn.commit()

#####Imports######
import pprint
import random
import requests
import urllib.request
from selenium import webdriver
import time
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.keys import Keys  
from bs4 import BeautifulSoup
import pandas as pd
import pyodbc
import sqlalchemy
import pandas as pd
from datetime import datetime
import win32com.client
from collections import Counter
from datetime import date
import pprint
import random
import requests
import urllib.request
from selenium import webdriver
import time
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.common.keys import Keys  
from bs4 import BeautifulSoup
import pandas as pd
import pyodbc
import sqlalchemy
import pandas as pd
from datetime import datetime
import win32com.client
from collections import Counter
from datetime import date
import os
import math
from datetime import datetime, timedelta


import win32com.client as win32

conn = pyodbc.connect(
"Driver={SQL Server};"
"Server=DESKTOP-F86F289;"
"Database =Planeo;"
"Trusted_Connection=yes;")

today = date.today()

df = pd.DataFrame()
df = pd.read_sql("SELECT a.popis,a.Nazov, b.item_name_downloaded, a.cena,b.price, b.alts from (SELECT * FROM Planeo.dbo.Planeo WHERE Date in (SELECT Convert(DateTime, DATEDIFF(DAY, 0, GETDATE())))) as a inner Join Planeo.dbo.Heureka as b on a.Nazov = b.item_name",conn)
df = df[df['item_name_downloaded'].isin(list(df[df['alts'] == 'planeo.sk']['item_name_downloaded'].unique()))]
df = df.drop_duplicates()
df = df[df['alts'] != 'planeo.sk']
df = df[df['price'] == df.groupby('item_name_downloaded')['price'].transform('min')]
df = df.reset_index(drop=True)


def categorise(row):  
    if row['item_name_downloaded'].strip().lower() == row['Nazov'].strip().lower():
        return 'Yes'
    if row['item_name_downloaded'].strip().lower() in row['Nazov'].strip().lower():
        return 'Maybe'
    else:
        return 'No'
df['comparison'] = df.apply(lambda row: categorise(row), axis=1)
df = df.sort_values(by='comparison', ascending=False)
df['zlava'] = round((df['cena'] / df['price']) - 1,2)
df.rename(columns={'popis':'Category', 'Nazov': 'item_name',"zlava": "cheaper_by_%", "alts" : "second_best_price", "cena" : "price_planeo", "price" : "best_price_heureka", "comparison" : "item_is_same"}, inplace=True)
df['cheaper_by_%'] = round(df['cheaper_by_%']*100)
df = df[['Category','item_name', 'item_name_downloaded', 'item_is_same', 'second_best_price', 'best_price_heureka','price_planeo','cheaper_by_%']]
df = df.sort_values(by=['item_is_same', 'cheaper_by_%'],  ascending=[False,True])
df = df.drop_duplicates(subset=['item_name_downloaded'])
df = df.reset_index(drop=True)
df.to_excel(r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_Heureka' + str(today) +'.xlsx')
df = df.head(20)
os.startfile("outlook")
receiver = 'zrubanrobert@gmail.com'
list_of_email = ['zrubanrobert@gmail.com']
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = receiver
mail.Subject = f'Planeo zceny oproti Heureke' 
html1 = df.to_html()
mail.HTMLBody = html1
mail.Attachments.Add(Source=r'C:\Users\roboz.DESKTOP-F86F289\Desktop\Planeo\Planeo_Heureka' + str(today) +'.xlsx')
mail.Display()
mail.Send()