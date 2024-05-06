from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil

@task
def order_robots():
    """Download the cvs with the info of robots to order, read the csv file and complete the web form to order all the robots"""
    browser.configure(
        slowmo=100,
        # browser_engine="chrome",
    )
    open_the_intranet_website()
    download_csv_file()
    read_csv_file()
    create_zip_file()
    clean_up()

def open_the_intranet_website():
    """Navigates to the robot order web page"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def accept_intranet_conditions():
    """Click 'OK' on the Conditions POP-UP"""
    page = browser.page()
    page.click("text=OK")

def download_csv_file():
    """Download CSV file that will be used as input for the robot order"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_csv_file():
    """Reads the downloaded CSV file"""
    
    table = Tables()
    csvData = table.read_table_from_csv("orders.csv", header=True)

    for robot in csvData:
        accept_intranet_conditions()
        fill_and_submit_order_form(robot)
        order_number = str(robot['Order number'])
        pdf_path = export_as_pdf(order_number)
        image_path = take_screenshot(order_number)
        add_screenshot_to_pdf(pdf_path, image_path)
        order_another_robot()


def fill_and_submit_order_form(robot):
    """Fill the form using the input parameter and submit the robot order"""
    page = browser.page()

    page.select_option("#head", str(robot['Head']))
    page.click(f"xpath=//input[@type='radio'][@value='{str(robot['Body'])}']")
    page.fill("xpath=//input[@type='number']", str(robot['Legs']))
    page.fill("#address", str(robot['Address']))

    page.click("xpath=//button[@id='order']")

    validate_order_creation()


def validate_order_creation():
    """Validation if the click Orders button worked or needs to be re-clicked"""
    page = browser.page()
    order_another_exists = page.query_selector("#order-another")
    if not order_another_exists:
        page.click("xpath=//button[@id='order']")
    # try:
    #     page.is_enabled("xpath=//button[@id='order-another']", strict= True, timeout=5000)
    # except:
    #     page.click("xpath=//button[@id='order']")
        
def export_as_pdf(order_number):
    """Export the receipt to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = f'output/pdfs/receipt_robot_{order_number}.pdf'
    pdf.html_to_pdf(sales_results_html, pdf_path)
    return pdf_path

def take_screenshot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    image_file = f'output/screenshots/screenshot_robot_{order_number}.png'
    page.screenshot(path = image_file)
    return image_file

def add_screenshot_to_pdf(pdf_path, image_path):
    """Append one file to the pdf"""
    pdf = PDF()

    list_of_files = [
            f'{image_path}:align=center',
        ]

    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document = f"{pdf_path}",
        append = True
    )

def order_another_robot():
    """Click on Order Another Robot"""
    page = browser.page()
    page.click("xpath=//button[@id='order-another']")

def create_zip_file():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("output/pdfs", "output/receipts.zip")

def clean_up():
    """Cleans up the folders where receipts and screenshots are saved."""
    shutil.rmtree("output/pdfs")
    shutil.rmtree("output/screenshots")
    shutil.rmtree("output/receipts.zip")