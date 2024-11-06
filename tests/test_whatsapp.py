import time
import pytest
import pandas as pd
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import logging
import os
from pytest_html import extras
from io import StringIO
import conftest

log_stream = StringIO()
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', stream=log_stream)

chrome_profile_directory = "C:/Users/shonz/AppData/Local/Google/Chrome/User Data/Default"

@pytest.fixture(scope="class")
def setup(request):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument(f"user-data-dir={chrome_profile_directory}")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--ignore-certificate-errors")
    service_obj = Service("C:/Users/shonz/Desktop/selenium/chromedriver2/chromedriver.exe")
    driver = webdriver.Chrome(service=service_obj, options=chrome_options)
    driver.implicitly_wait(5)
    driver.get("https://web.whatsapp.com/")
    time.sleep(10)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "(//button[@type='button'])[2]")))
    request.cls.driver = driver
    yield driver
    driver.quit()

@pytest.mark.usefixtures("setup")
class TestWhatsApp:

    @pytest.mark.parametrize("term",conftest.terms )
    def test_check_text(self, term, request):
        request.node.name = f"test_check_text_{term}"
        time.sleep(20)
        wait = WebDriverWait(self.driver, 10)
        not_called = wait.until(EC.presence_of_element_located((By.XPATH, "(//button[@type='button'])[2]")))
        assert not_called is not None, "האלמנט לא נמצא!"
        if not_called.get_attribute("aria-pressed") == "false":
            logging.info("האלמנט נמצא, לוחצים עליו.")
            not_called.click()
        else:
            logging.info("האלמנט נמצא, כבר לחוץ.")

        screenshot_path = f'C:/Users/shonz/PycharmProjects/pythonProject/tests/screenshots/screenshot_{term}.png'
        search_field = None

        try:
            logging.info(f"התחלת חיפוש עבור המילה: {term}")
            search_field = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.selectable-text.copyable-text.x15bjb6t.x1n2onr6'))
            )
            search_field.click()
            search_field.send_keys(term)
            time.sleep(6)
            logging.info(f"חיפוש המילה: {term} הושלם.")
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@role='listitem']"))
            )

            div_elements = self.driver.find_elements(By.XPATH, "//div[@role='listitem']")

            os.makedirs(os.path.dirname(screenshot_path), exist_ok=True)
            self.driver.save_screenshot(screenshot_path)
            logging.info(f"צילום מסך של תוצאות החיפוש נשמר ב-{screenshot_path}")
            self.add_screenshot_to_report(request, screenshot_path)

            if div_elements:
                logging.info(f"נמצאו תוצאות עבור המילה: {term}. שומרים את התוצאות לאקסל...")
                div_list = [div.text.split("\n") for div in div_elements]
                df = pd.DataFrame(div_list).dropna()
                df.to_excel(f'C:/Users/shonz/PycharmProjects/EXCLE_{term}.xlsx', index=False, header=False)
                self.save_results(div_list)
                logging.info(f"התוצאות עבור המילה: {term} נשמרו בהצלחה!")

            self.clear_search_field(search_field)
            logging.info("הטקסט הקודם נמחק.")

        except Exception as e:
            logging.warning(f"לא נמצאו תוצאות עבור המילה: {term}. שגיאה: {str(e)}")
            if search_field:
                self.clear_search_field(search_field)
            self.add_screenshot_to_report(request, screenshot_path)

    def clear_search_field(self, search_field):
        search_field.click()
        search_field.send_keys(Keys.CONTROL + "a")
        search_field.send_keys(Keys.BACKSPACE)
        time.sleep(6)

    def save_results(self, rows):
        with open('results.json', 'w') as f:
            json.dump(rows, f)
        logging.info("תוצאות נשמרו לקובץ results.json.")

    def add_screenshot_to_report(self, request, screenshot_path):
        try:
            # בדוק אם הנתיב קיים ואם לא, צור אותו
            if not os.path.exists(screenshot_path):
                logging.warning(f"צילום מסך לא נמצא בנתיב: {screenshot_path}")
                return

            # הוסף את התמונה לדוח כקישור שניתן להקליק עליו להגדלה
            extra_image = extras.html(
                f'<a href="{screenshot_path}" target="_blank">'
                f'<img src="{screenshot_path}" alt="screenshot" style="max-width:300px; max-height:150px;"/></a>'
            )

            if not hasattr(request.node, 'extra'):
                request.node.extra = []
            request.node.extra.append(extra_image)

            logging.info("צילום המסך נוסף לדוח עם קישור להגדלה.")
        except Exception as e:
            logging.error(f"נכשלה הוספת צילום המסך לדוח: {str(e)}")

@pytest.hookimpl(tryfirst=True)
def pytest_runtest_makereport(item, call):
    if call.when == "call":
        # הוספת כל הלוגים לדוח
        log_output = log_stream.getvalue()
        if log_output:
            extra_logs = extras.html(f"<div><pre>{log_output}</pre></div>")
            item.extra = getattr(item, 'extra', []) + [extra_logs]
        log_stream.truncate(0)
        log_stream.seek(0)
