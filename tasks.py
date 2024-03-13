from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.Excel.Application import Application
from RPA.Excel.Files import Files
import RPA.Tables as table 
import json
import time
import datetime
import re
from dateutil.relativedelta import relativedelta
from datetime import date


with open("config_file.json","r") as f:
    data = json.load(f)

browser = Selenium()
exFile = Files()
list = []

exFile.create_workbook("output","xlsx","data")

                       
def open_web_page():
    browser.open_available_browser(data['url'])
    browser.maximize_browser_window()

def main():
    try:
        open_web_page()
        search_phrase()
    finally:
        browser.close_browser()


def search_phrase():
    browser.click_button_when_visible(data['cookieas_button'])
    browser.click_element_if_visible(data['search_button'])
    browser.press_keys(None,data['search_phrase'])
    browser.press_keys(None,"ENTER")
    filters_section()
    time.sleep(2)
    filter_date_range()
    extract_data()
    time.sleep(1)

def filters_section():
    browser.reload_page()
    browser.click_button(data['section'])
    if data['senction_filter_active'] == "yes":
        for k,v in data['senction_filters'].items():
            if k in data['senction_filter']:
                time.sleep(1)
                browser.click_element(v)
    else:
        browser.click_element(data['any'])
    browser.click_button(data['section'])

def filter_date_range():
    position1 = data['startDate']
    position2 = data['endDate']
    if data['data_range_filter_active'] == "yes":
        browser.click_button(data['data_range'])
        for k,v in data['data_range_filters'].items():
            if k in data['data_range_filter']:
                time.sleep(1)
                browser.click_element(v)
                if k == "specific_date":
                    searchDate = filter_month()
                    time.sleep(2)
                    browser.click_element(position1)
                    browser.input_text(position1,searchDate)
                    browser.input_text(position2,date.today().strftime(f"%m/%d/%Y"))
                    browser.press_keys(position2, "ENTER")

def filter_month():
    today = datetime.date.today()
    result = ""

    if data['number_of_months'] < 2:
        first = today.replace(day=1)
        result = first.strftime(f"%m/%d/%Y")
    else:
        current_date = date.today()
        past_date = current_date - relativedelta(months=data['number_of_months'])
        past_date = past_date.replace(day=1)
        result = past_date.strftime(f"%m/%d/%Y")
    return result

def extract_data():
    total_matches = len(browser.get_webelements("tag:li"))
    end_matches = range(1,total_matches)

    title = extract_title(total_matches,end_matches)
    date = extract_date(total_matches,end_matches)
    description = extract_description(total_matches,end_matches)
    img = extract_img_photo(title,end_matches)
    complete_excel(title,date,description,total_matches,img)

def extract_title(total_matches,end_matches):
    title_list = []
    for i in end_matches:
        title_data = browser.get_webelements(f"//*[@id='site-content']/div/div[2]/div[1]/ol/li[{i}]/div/div/div/a/h4")
     
        if len(title_data) == 0:
            title = "0"
        else:
            title = title_data
        
        for match in title:
            if match == "0":
                title_list.append("Empty")
            else:
                title_list.append(match.text)

    return  title_list  

def extract_date(total_matches,end_matches):
    date_list = []

    for i in end_matches:
        date_data = browser.get_webelements(f"//*[@id='site-content']/div/div[2]/div[1]/ol/li[{i}]/div/span")

        if len(date_data) == 0:
            date = "0"
        else:
            date = date_data
        
        for match in date:
            if match == "0":
                date_list.append("Empty")
            else:
                date_list.append(match.text)

    return  date_list  

def extract_description(total_matches,end_matches):
    description_list = []

    for i in end_matches:
        description_data = browser.get_webelements(f"//*[@id='site-content']/div/div[2]/div[1]/ol/li[{i}]/div/div/div/a/p[1]")

        if len(description_data) == 0:
            description = "0"
        else:
            description = description_data
        
        for match in description:
            if match == "0":
                description_list.append("Empty")
            else:
                description_list.append(match.text)

    return  description_list 


def complete_excel(title,date,description,total_matches,img):
    max_range = total_matches - 1
    exFile.set_cell_value(1,"A","Title")
    exFile.set_cell_value(1,"B","Date")
    exFile.set_cell_value(1,"C","Description")
    exFile.set_cell_value(1,"D","Image Path")
    exFile.set_cell_value(1,"D","Count Phrases")


    for i in range(max_range):
        exFile.set_cell_value(i + 2,"A",title[i])
        exFile.set_cell_value(i + 2,"B",date[i])
        exFile.set_cell_value(i + 2,"C",description[i])
        exFile.set_cell_value(i + 2,"D",img[i])
        exFile.set_cell_value(i + 2,"D", title[i].count(data['search_phrase']) + description[i].count(data['search_phrase']))



    
    exFile.save_workbook("output/data.xlsx")

def extract_img_photo(title,end_matches):
    img_list = []
    for i in end_matches:
        name = re.sub('[^a-zA-Z0-9 \n\.]', '', title[i -1])
        img_data = browser.get_webelements(f"//*[@id='site-content']/div/div[2]/div[1]/ol/li[{i}]/div/div/figure/div/img")
        if len(img_data) == 0:
            img_list.append(f"No Image")
            img = "0"
        else:
            browser.screenshot(f"//*[@id='site-content']/div/div[2]/div[1]/ol/li[{i}]/div/div/figure/div/img", f"output/{name}{i}.png")
            img_list.append( f"output/{name}{i}.png")

    return  img_list       


if __name__ == "__main__":
    main()



