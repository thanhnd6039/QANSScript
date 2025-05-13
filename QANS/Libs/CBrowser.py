from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions

class CBrowser(object):
    def get_chrome_options(self):
        options = ChromeOptions()
        prefs = {'download.default_directory':'C:\\RobotFramework\\Output\\',
                 "download.prompt_for_download": False,
                 "safebrowsing.enabled": True,
                 'safebrowsing.disable_download_protection': True,
                 "profile.default_content_settings.popups": 0,
                 }
        options.add_experimental_option('detach', True)
        options.add_experimental_option('prefs', prefs)
        return options
    def get_firefox_options(self):
        options = FirefoxOptions()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.dir", 'C:\\RobotFramework\\Output\\')
        return options