import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import lxml



options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),chrome_options=options)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install())) #To hide the chrome window from popping up while scraping, the above line could be uncommented and this line is to be commented out

start_time = datetime.now()
'''Declaration of the columns that will be exported in excel file'''
columns_list = ["url","product_title","size","price", "variation_id"]
result = pd.DataFrame(columns = columns_list)



Links = pd.read_excel(r'path to input file')
Links_list = Links['Link'].tolist()



'''Function to scroll down to the end of the page, if required, to collect product information at the bottom'''
def scroll_down(driver):
    start = time.time()
    initialScroll = 0
    finalScroll = 1000
    while True:
        driver.execute_script(f"window.scrollTo({initialScroll}, {finalScroll})")
                        
        initialScroll = finalScroll
        finalScroll += 1000

        time.sleep(1)
        end = time.time()
                        
        if round(end - start) > 2:
            break

'''Function to extract info'''
def do_task(iteration,driver):

    url = Links.iloc[iteration,0] 
    driver.get(url)
    scroll_down(driver) 
    src = driver.page_source
    soup = BeautifulSoup(src, 'lxml') 
    
    
    product_title = soup.find("h1",{"class":"product-title"}).get_text().split('\n')[3].strip()
    

    price = soup.find("div",{"class":"price--main"}).find("span",{"class":"money"}).get_text().split('\n')[1] 


    if "x" not in url.split('-')[-1] and "?" in url.split('-')[-1]: 
        
        size = soup.find("select",{"id":"data-product-option-0"}).findAll("option")[0].get_text().split('\n')[1].strip() 
        variant_id = url.split('?')[-1].split('=')[-1] 
    else:
        size = url.split('-')[-1] 
        variant_id = 'Not Available'
    
   
  
    output = [url,product_title,size,price,variant_id]
    print(output)
    
    return url, output


for url,iteration in zip(Links_list,range(len(Links))):
    print('on the {} link (total: {})'.format(iteration + 1, len(Links))) 
     
    ''' If the function above throws an error, for some unavialable pages, the exception is handled below''' 
    try:
        output = do_task(iteration,driver)
        result.loc[len(result)] = output
    except:
        product_title = size = price = variant_id ="Not Available"
        output = [url, product_title,size,price,variant_id] #if any link in the input file directs to a unavailable page, report the link and unavailibity of the required data
        result.loc[len(result)] = output
        
        continue
   
    

'''Exporting result to target folder'''
end_time = datetime.now()
timestr = time.strftime("%d.%m.%Y-%H%M%S")
print('Duration: {}'.format(end_time - start_time))

result.to_excel('results_file.xlsx', index=False)