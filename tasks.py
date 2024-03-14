from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil

pdf = PDF()
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
    download_order_file()
    customers = create_table()
    get_orders(customers)
    make_order()
    archive_receipts()
    # clean_up()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_order_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def store_receipt_as_pdf(order_number,order_receipt_html):
    pdf_path = pdf.html_to_pdf(order_receipt_html, "output/receipts/"+f"{order_number}.pdf")
    return pdf_path

def screenshot_robot(order_number):
    print(order_number)
    page = browser.page()
    screenshot_path = page.screenshot(path="output/screenshots/"+f"{order_number}.png", full_page=False)
    return screenshot_path

def embed_screenshot_to_receipt(order_number):
    list_of_files = [f'output/screenshots/{order_number}.png',
                     f'output/receipts/{order_number}.pdf']
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f"output/receipts/{order_number}.pdf"
    )

def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def clean_up():
    """Cleans up the folders where receipts and screenshots are saved."""
    shutil.rmtree("./output/receipts")
    shutil.rmtree("./output/screenshots")

def get_orders(customers):
    print(customers)
    page = browser.page()
    for order in customers:
        close_annoying_modal()
        page.select_option("#head", str(order['Head']))
        page.click("#id-body-"+str(order['Body']))
        legs =str(order['Legs'])
        page.fill('//input[@class="form-control"]',legs)
        page.fill("#address" , order['Address'])
        page.click("button:text('preview')")
        browser.configure(slowmo=100,)
        page.click("button:text('order')")
        order_receipt_html = page.locator("#order-completion").inner_html()
        start_index = order_receipt_html.find('RSB-ROBO-ORDER-')
        end_index = order_receipt_html.find('</p>', start_index)
        order_number = order_receipt_html[start_index:end_index]
        print("Order Number:", order_number)
        print("order_receipt_html: ",order_receipt_html)
        store_receipt_as_pdf(order_number=order_number,order_receipt_html=order_receipt_html)
        screenshot_robot(order_number)
        embed_screenshot_to_receipt(order_number=order_number)   
        page.click("button:text('Order another robot')")

def make_order():
    page = browser.page()
    page.click("button:text('ORDER')")

def create_table():
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number","Head", "Body", "Legs", "Address"]
    )
    customers = library.group_table_by_column(orders, "Head")
    orders = []
    for customer in customers:
        for row in customer:
            order = {
                "Head": int(row["Head"]),
                "Body": int(row["Body"]),
                "Legs": int(row["Legs"]),
                "Address": row["Address"]
            }
            orders.append(order)
    return orders