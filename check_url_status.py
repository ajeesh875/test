from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By

class SharePointLinkChecker:
    def __init__(self, msedge_driver_path):
        self.msedge_driver_path = msedge_driver_path
        self.options = EdgeOptions()
        self.options.use_chromium = True
        self.options.add_argument("--no-sandbox")

    def _get_link_urls(self, wrapper):
        buttons = wrapper.find_elements(By.TAG_NAME, 'button')
        link_urls = []
        for button in buttons:
            onclick_attribute = button.get_attribute('onclick')
            if 'http' in onclick_attribute:
                link_urls.append(onclick_attribute.split("'")[1])
            else:
                print("No link found in the button onclick attribute.")
        return link_urls

    def check_sharepoint_links(self, sharepoint_url):
        service = EdgeService(executable_path=self.msedge_driver_path)
        driver = webdriver.Edge(service=service, options=self.options)

        try:
            driver.get(sharepoint_url)
            button_wrappers = driver.find_elements(By.TAG_NAME, 'buttonwrappers')

            link_urls = []
            for wrapper in button_wrappers:
                link_urls.extend(self._get_link_urls(wrapper))
                
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            driver.quit()
            # The link_urls list will not be returned here.
