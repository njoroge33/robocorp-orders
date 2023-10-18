from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    get_orders()
    read_excel_data()
    archive_receipts()
    


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_excel_data():
    tables = Tables()

    # Read data from a CSV file into a table
    csv_file = "orders.csv"
    table = tables.read_table_from_csv(csv_file)

    # Iterate through the rows in the table
    for row in table:
        # Access individual columns within the row
        fill_the_form(row)
        

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def fill_the_form(data):
    page = browser.page()
    close_annoying_modal()

    page.select_option("#head", data["Head"])
    page.click("#id-body-"+data["Body"])
    page.fill("form>div:nth-child(3)>input", data["Legs"])
    page.fill("#address", data["Address"])
    page.click("text=Preview")
    page.click("#order")

    while not page.query_selector("#order-another"):
        page.click("#order")

    store_receipt_as_pdf(data["Order number"])
    screenshot_robot(data["Order number"])

    # embed_screenshot_to_receipt(f"output/{data['Order number']}.png", f"output/{data['Order number']}.pdf")

    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    page = browser.page()

    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_location = f"output/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_location)

def screenshot_robot(order_number):
    page = browser.page()

    pic = page.query_selector("#robot-preview-image")
    screenshot = f"output/{order_number}.png"
    pic.screenshot(path=screenshot)

    embed_screenshot_to_receipt(f"output/{order_number}.png", f"output/{order_number}.pdf")

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def archive_receipts():
    folder = Archive()
    folder.archive_folder_with_zip('output', 'output/orders.zip', include='*.pdf')

