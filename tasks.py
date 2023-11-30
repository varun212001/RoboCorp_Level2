from robocorp.tasks import task

from robocorp import browser
from RPA import HTTP
from RPA import PDF
from RPA.Archive import Archive
from RPA import Tables
from RPA.Browser import Selenium
import time

browser = Selenium.Selenium()

@task
def rpa_order_robots_from_robotsparebin():
    '''automation of manual task, ordering robot in robotsparebin'''
    try:
        download_csv_from_intranet_website()
        open_intranet_website_and_navigate_specific_page_through_browser()
        iterate_filling_process(get_data_from_downloaded_csv())
        zip_folders_in_output()
    finally:
        close_interanet_website()

def open_intranet_website_and_navigate_specific_page_through_browser():
    '''open intranet website through browser'''
    browser.open_available_browser('https://robotsparebinindustries.com/#/robot-order')

def download_csv_from_intranet_website():
    '''download csv from intranet website'''
    http = HTTP.HTTP()
    http.download('https://robotsparebinindustries.com/orders.csv',overwrite=True)

def get_data_from_downloaded_csv():
    '''get data from downloaded csv file and store it into a table'''
    table = Tables.Tables()
    extracted_table= table.read_table_from_csv('orders.csv')
    return extracted_table

def filling_process(row):
    '''filling the order form with given data'''
    close_alert()
    browser.execute_javascript(f"document.getElementById('head').value = {row['Head']};")
    browser.select_radio_button('body',row['Body'])
    browser.input_text("//input[@placeholder='Enter the part number for the legs']",row['Legs'])
    browser.input_text('address',row['Address'])

    browser.wait_until_element_is_enabled('id: preview',timeout=10)
    try:
        browser.execute_javascript("document.getElementById('preview').click();")
    except Exception as e:
        print(e)
        browser.execute_javascript("document.getElementById('preview').click();")
    time.sleep(3)

    screenshoot_particular_element(row['Order number'])
    place_order()
    #browser.wait_until_element_is_enabled('id: receipt',timeout=10)
    create_pdf_with_bill(row['Order number'])
    modify_pdf_by_embedding_image(row['Order number'])

def iterate_filling_process(table):
    '''iterating over each row to fill the form'''
    for each_row in table:
        filling_process(each_row)
        start_fill_and_order_again()
        time.sleep(1)

def start_fill_and_order_again():
    '''start filling process again after completing an orders'''
    browser.execute_javascript("document.getElementById('order-another').click();")

def close_alert():
    '''close that annoying alert'''
    try:
        browser.click_element("//button[contains(text(), 'OK')]")
    except Exception as e:
        print(e)
        browser.execute_javascript("document.getElementById('OK').click();")

def screenshoot_particular_element(orderNumber):
    '''take screenshot of particular element'''
    browser.screenshot('id: robot-preview-image',f'output/robots/{orderNumber}.png')

def place_order():
    '''place order of robot which we filled'''
    max_retries = 10
    for attempt in range(max_retries):
        browser.execute_javascript("document.getElementById('order').click();")
        if find_alerts():
            print(f"Attempt: {attempt+1} failed with error!")
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
        else:
            break

def find_alerts():
    '''find alerts which causes ERROR and leads to stopping our RPA'''
    alert_locator = "xpath://div[contains(@class, 'alert')]"
    alert_elements = browser.find_elements(alert_locator)
    for alert in alert_elements:
        if 'alert-danger' in alert.get_attribute('class'):
            print(f'Alert found: {alert.text}')
            return True
    print(alert_elements,len(alert_elements))

def create_pdf_with_bill(orderNumber):
    '''create pdf with bill details'''
    pdf = PDF.PDF()
    pdf.html_to_pdf(get_bill_data(),f'output/bills/{orderNumber}.pdf')

def get_bill_data():
    '''get bill data from the desplayed bill in the page'''
    max_retries = 3
    for attempt in range(max_retries):
        try:
            bill_as_html = browser.get_element_attribute('id: receipt','innerHTML')
        except Exception as e:
            print(f"Attempt {attempt+1} failed with error: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
    return bill_as_html

def modify_pdf_by_embedding_image(orderNumber):
    '''modify pdf by embedding image'''
    pdf = PDF.PDF()
    pdf.add_watermark_image_to_pdf(image_path=f'output/robots/{orderNumber}.png',source_path=f'output/bills/{orderNumber}.pdf',output_path=f'output/bills/{orderNumber}.pdf')

def zip_folders_in_output():
    '''zip all the files in bills, robots folder in output folder'''
    archive = Archive()
    try:
        archive.archive_folder_with_zip('output/robots','output/robots.zip', recursive=True)
        archive.archive_folder_with_zip('output/bills','output/bills.zip', recursive=True)
        
    except Exception as e:
        print(e)

def close_interanet_website():
    '''closing the opened browser, interanet website'''
    browser.close_browser()