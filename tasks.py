from robocorp.tasks import task
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from robocorp import browser
from zipfile import ZipFile
import os
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
        slowmo=10,
    )
    open_robot_order_website()
    orders = get_orders()
    
    for r in orders:
        close_annoying_modal()
        fill_the_form(r)
        save_preview(r["Order number"])
        click_order()
        pdfpath = store_receipt_as_pdf(r["Order number"])
        imagepath = screenshot_robot(r["Order number"])
        embed_screenshot_to_receipt(imagepath,pdfpath)
        click_another_robot()
    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", header=True)
    return orders


def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def fill_the_form(row):
    print(row)
    page = browser.page()
    page.select_option("#head",row["Head"])
    page.set_checked("#id-body-"+str(row["Body"]), True) 
    page.fill(".form-control:nth-child(3)",row["Legs"])
    page.fill("#address",row["Address"])
    
def save_preview(order_number):
    page = browser.page()
    page.click("text=Preview")
    page.locator("#robot-preview-image").screenshot(path="output/screenshots/"+ order_number +".png")
    page = browser.page()



def click_order():
    
    page = browser.page()
    orderbtn = page.locator("#order")
    while orderbtn.is_visible():
        orderbtn.click()

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt = page.locator("#receipt").inner_html()
    path = "output/receipts/robot"+ order_number +".pdf" 
    pdf = PDF()
    pdf.html_to_pdf(receipt, path)
    return path

def screenshot_robot(order_number):
    page = browser.page()
    pathimage = "output/receipts/receipt"+ order_number +".png" 
    page.locator("#robot-preview").screenshot(path=pathimage)
    return pathimage


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf([screenshot],pdf_file,True)



def click_another_robot():
    page = browser.page()
    nextorderbtn = page.locator("#order-another")
    nextorderbtn.click()

def archive_receipts():
    pathfiles = "output/receipts/"
    outputfile ="output/receipts.zip"
    with ZipFile(outputfile, 'w') as zip:
        for file in os.listdir(path = pathfiles):
            zip.write(pathfiles + file, arcname=file)


