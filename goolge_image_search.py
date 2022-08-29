import undetected_chromedriver.v2 as uc
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd

def save_to_excel(hrefs):
    print("Saving into excel file---")
    save_file_name = input("Save Excel File Name (ex: results.xlsx) : ")

    df = pd.DataFrame({'Image Source URL' : hrefs})

    writer = pd.ExcelWriter(save_file_name,engine='xlsxwriter')

    df.to_excel(writer,sheet_name='Results',index=False)

    writer.save()

    print("Successfully saved in " + save_file_name)

if(__name__ == '__main__'):
    options = uc.ChromeOptions()
    options.add_argument("--headless")
    driver = uc.Chrome(options=options)

    driver.delete_all_cookies()

    driver.get('https://www.google.ru/imghp?hl=en&ogbl')

    sleep(2)

    driver.find_element('xpath','/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[3]/div[3]/img').click()

    sleep(2)

    image_url = input('Please insert Image URL : ')

    c_wiz = driver.find_element('xpath','//*[@id="ow7"]/div[3]/c-wiz')
    url_input = c_wiz.find_element(By.TAG_NAME,'input')

    url_input.send_keys(image_url)

    sleep(2)

    url_input.find_element('xpath','following-sibling::div').click()

    print("Searching results----")

    sleep(5)

    div_result = driver.find_element('xpath','//*[@id="yDmH0d"]/div[3]/c-wiz/div/c-wiz/div/div[2]/div/div/div')
    
    a_tags = div_result.find_elements(By.TAG_NAME,'a')
    
    hrefs = []
    for a_tag in a_tags:
        hrefs.append(a_tag.get_attribute('href'))
        print(a_tag.get_attribute('href'))
    
    
    print(str(len(hrefs)) + ' Results found' )
    save_to_excel(hrefs)


