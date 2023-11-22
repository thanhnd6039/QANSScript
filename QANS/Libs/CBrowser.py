from selenium import webdriver

class CBrowser(object):
    def get_chrome_options(self):
        options = webdriver.ChromeOptions()
        prefs = {'download.default_directory':'C:\\RobotFramework\\Downloads\\'}
        options.add_experimental_option('detach', True)
        options.add_experimental_option('prefs', prefs)
        return options