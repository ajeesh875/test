from selenium import webdriver
from sharepoint_link_checker import SharePointLinkChecker
from outlook_email_sender import OutlookEmailSender
from selenium.webdriver.edge.options import Options as EdgeOptions

def check_url_status(urls, driver):
    status_codes = []
    for url in urls:
        driver.get(url)
        status_codes.append(str(driver.execute_script("return window.performance.getEntries()[6].responsestatus")))
    return status_codes

if __name__ == "__main__":
    sharepoint_url = "https://lloydsbanking.sharepoint.com/sites/interactiveprojects/Shared%20Documents/Consumable/consumable.aspx"
    msedge_driver_path = "msedgedriver.exe"  # Provide the path to the Microsoft Edge WebDriver executable

    link_checker = SharePointLinkChecker(msedge_driver_path)
    link_urls = link_checker.check_sharepoint_links(sharepoint_url)

    if link_urls:
        driver = webdriver.Edge(executable_path=msedge_driver_path, options=EdgeOptions())
        try:
            body = check_url_status(link_urls, driver)
            sender = OutlookEmailSender()
            sender.send_email(body)
        finally:
            driver.quit()
