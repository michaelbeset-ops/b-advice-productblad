# -*- coding: utf-8 -*-
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
import time


class WebDriver:
    def __init__(self, driver_path: str, starting_url: str):
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)

        # Probeer eerst webdriver-manager (downloadt automatisch de juiste versie)
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
        except Exception:
            # Fallback: gebruik meegeleverde chromedriver.exe
            service = Service(executable_path=driver_path)

        # Geen try/except hier zodat fouten zichtbaar zijn in de GUI
        self.web_driver = webdriver.Chrome(service=service, options=options)
        self.web_driver.get(starting_url)
        self._wait_for_page_load()

    def _wait_for_page_load(self, timeout: int = 30):
        try:
            WebDriverWait(self.web_driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except Exception:
            pass

    def new_tab(self, url: str, extra_wait: float = 1.0):
        self.web_driver.execute_script('window.open();')
        self.web_driver.switch_to.window(self.web_driver.window_handles[-1])
        self.web_driver.get(url)
        self._wait_for_page_load()
        if extra_wait > 0:
            time.sleep(extra_wait)

    def new_search(self, item_id: str, search: str):
        try:
            search_bar = self.web_driver.find_element(By.ID, item_id)
            search_bar.clear()
            search_bar.send_keys(search)
            search_bar.send_keys(Keys.RETURN)
        except Exception as e:
            print(f'Could not find specified web item: {item_id} - {e}')

    def __del__(self):
        try:
            self.web_driver.quit()
        except Exception:
            pass
