# SharePoint Link Checker and Outlook Email Sender

This Python project provides a solution to check SharePoint links and send the results through an Outlook email. It is implemented using object-oriented programming (OOP) and includes three classes:

1. `SharePointLinkChecker`: This class checks SharePoint links and extracts URLs from button elements.

2. `OutlookEmailSender`: This class sends an email via Outlook with the given body.

3. `WebsiteChecker`: This class verifies if a website is up and running based on its status code.

## Prerequisites

- Python 3.x
- Microsoft Edge WebDriver (msedgedriver.exe) for Selenium (Ensure the driver version matches your Microsoft Edge browser version)

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your_username/sharepoint-link-checker.git

sharepoint_url = "https://lloydsbanking.sharepoint.com/sites/interactiveprojects/Shared%20Documents/Consumable/consumable.aspx"
msedge_driver_path = "msedgedriver.exe"  # Provide the path to the Microsoft Edge WebDriver executable

2. Run the main script
   python main.py
Classes
SharePointLinkChecker

This class checks SharePoint links and extracts URLs from button elements. It utilizes Selenium to interact with the SharePoint page and retrieve the URLs.
OutlookEmailSender

This class sends an email via Outlook with the given body. It uses the win32com library to interact with Outlook and send the email.
WebsiteChecker

This class verifies if a website is up and running based on its status code. It uses the requests library to send HTTP requests and obtain status codes.
