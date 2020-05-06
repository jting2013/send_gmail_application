from selenium.webdriver.support.ui import WebDriverWait
import time

class ClearCache:
    def __init__(self, driver, logging):
        self.driver = driver
        self.logging = logging

    def click_clear_browsing_button(self):
        """Find the "CLEAR BROWSING BUTTON" on the Chrome settings page."""
        try:
            self.driver.execute_script("return document.querySelector('settings-ui').shadowRoot.querySelector('settings-main').shadowRoot.querySelector('settings-basic-page').shadowRoot.querySelector('settings-section > settings-privacy-page').shadowRoot.querySelector('settings-clear-browsing-data-dialog').shadowRoot.querySelector('#clearBrowsingDataDialog').querySelector('#clearBrowsingDataConfirm')").click()
            return True
        except:
            return False

    def click_cache_checkbox(self):
        try:
            self.driver.execute_script(
            "return document.querySelector('settings-ui').shadowRoot.querySelector('settings-main').shadowRoot.querySelector('settings-basic-page').shadowRoot.querySelector('settings-section > settings-privacy-page').shadowRoot.querySelector('settings-clear-browsing-data-dialog').shadowRoot.querySelector('#clearBrowsingDataDialog').querySelector('#cacheCheckboxBasic')").click()
            return True
        except:
            return False

    def clear_cache(self, timeout=60):
        """Clear the cookies and cache for the ChromeDriver instance."""
        # navigate to the settings page
        self.driver.get('chrome://settings/clearBrowserData')
        # wait for the button to appear
        wait = WebDriverWait(self.driver, timeout)
        time.sleep(2)

        if not self.click_clear_browsing_button():
            self.click_cache_checkbox()
            time.sleep(2)
            self.click_clear_browsing_button()
