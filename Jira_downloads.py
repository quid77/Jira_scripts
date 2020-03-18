from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pathlib
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
from docx.shared import Inches, Pt
import glob
import re
import win32com.client as win32
from win32com.client import constants
import docx
import unittest


download_path = str(pathlib.Path(__file__).parent.absolute()) + "\\Downloads"
user_login = ""
user_password = ""

epics_dictionary = {}

class JiraTestsDownload(unittest.TestCase):

    @classmethod
    def setUpClass(self):  # setUpClass runs once for ALL tests
        options = webdriver.ChromeOptions()
        preferences = {"download.default_directory": download_path, "download.prompt_for_download": "false",
                       "safebrowsing.enabled": "false", 'profile.default_content_setting_values.automatic_downloads': 1}
        options.add_experimental_option("prefs", preferences)
        self.driver = webdriver.Chrome(options=options)

    def test_1_login_to_app(self):
        driver = self.driver
        driver.implicitly_wait(10)
        # user_login = input("Username: ")
        # user_password = input("Password: ")
        driver.get("https://jira.softserveinc.com/projects/SSMWE?selectedItem=com.thed.zephyr.je:zephyr-tests-page")
        login_element = driver.find_element_by_id("login-form-username")
        password_element = driver.find_element_by_id("login-form-password")
        login_element.send_keys(user_login, Keys.TAB)
        password_element.click()
        password_element.send_keys(user_password, Keys.ENTER)

    def test_2_order_tests(self):
        driver = self.driver
        wait = WebDriverWait(driver, 10)
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Unscheduled")))
        driver.find_element_by_link_text("Unscheduled").click()
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "order-options")))
        driver.find_element_by_class_name("order-options").click()
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "check-list-field-container")))
        search = driver.find_element_by_id("order-by-options-input")
        search.send_keys("Updated")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='updated-1']/label"))) # wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@title='Updated']")))
        search.send_keys(Keys.ENTER)
        time.sleep(3)

    def test_3_download_tests(self):
        driver = self.driver
        wait = WebDriverWait(driver, 2)
        while True:
            tests_list = driver.find_elements_by_class_name("splitview-issue-link")
            for test_element in tests_list:
                test_element.click()
                for x in range(0, 10):
                    try:
                        test_element_key = test_element.find_element_by_xpath(".//span[@class='issue-link-key']").text
                        wait.until(EC.presence_of_element_located((By.XPATH, "//a[@id='key-val']".format(test_element_key))))
                        JiraTestsDownload.add_test_to_epic(self, test_element_key)
                        break
                    except (StaleElementReferenceException, TimeoutException):
                        time.sleep(0.2)
                        continue
                for x in range(0, 10):
                    try:
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='viewissue-export']")))
                        driver.find_element_by_xpath("//*[@id='viewissue-export']").click()
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='jira.issueviews:issue-word']")))
                        driver.find_element_by_xpath("//*[@id='jira.issueviews:issue-word']").click()
                        break
                    except (StaleElementReferenceException, TimeoutException):
                        time.sleep(0.2)
                        continue
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='icon icon-next']")))
                driver.find_element_by_xpath("//a[@class='icon icon-next']").click()
                time.sleep(3)
            except NoSuchElementException:
                break
        print("Downloading done")

    def add_test_to_epic(self, test_element_key):
        driver = self.driver
        wait = WebDriverWait(driver, 1)
        driver.implicitly_wait(0)
        epic_name = ""
        try:
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-fieldtype='gh-epic-link']//a")))
                epic_name = driver.find_element_by_xpath("//div[@data-fieldtype='gh-epic-link']//a").text
            except (NoSuchElementException, StaleElementReferenceException):
                wait.until(EC.presence_of_element_located((By.XPATH, "(//a[@class='lozenge'])[2]")))
                epic_name = driver.find_element_by_xpath("(//a[@class='lozenge'])[2]").text
        except (StaleElementReferenceException, TimeoutException, NoSuchElementException):
            pass
        if epic_name.strip():
            epics_dictionary.setdefault(epic_name, []).append(test_element_key)
            print(epics_dictionary)

#
# def rename_files():
#     os.chdir(download_path)
#     for filename in os.listdir(download_path):
#         if filename.startswith("SSMWE"):
#             os.rename(filename, filename.replace("SSMWE-", "Copy-"))


def create_dir_hierarchy():
    if not os.path.exists(download_path + "\\DocFiles"):
        os.makedirs(download_path + "\\DocFiles")
    if not os.path.exists(download_path + "\\DocxFiles"):
        os.makedirs(download_path + "\\DocxFiles")
    if not os.path.exists(download_path + "\\TestTemplates"):
        os.makedirs(download_path + "\\TestTemplates")
    if not os.path.exists(download_path + "\\Directories"):
        os.makedirs(download_path + "\\Directories")


def move_doc_files():
    # Create list of paths to .doc files
    for filename in os.listdir(download_path):
        if filename.endswith(".doc") and filename.startswith("SSMWE"):
            if os.path.exists(download_path + "\\DocFiles\\" + filename):
                os.remove(download_path + "\\DocFiles\\" + filename)
            shutil.move(download_path + "\\" + filename, download_path + "\\DocFiles")


# this function isn't standalone, use save_to_docx instead
def save_as_docx(name):  # Github conversion solution
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(download_path + "\\DocFiles\\" + name)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(download_path + "\\DocxFiles\\" + name)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)


def save_to_docx():  # chceck obs if it doesnt work
    file_names = os.listdir(download_path + "\\DocFiles")
    for name in file_names:
        save_as_docx(name)


def remove_doc_files():  # NOW IM CHECKING FLOW, IS THIS NECESSARY?
    os.chdir(download_path + "\\DocxFiles")
    filenames = os.listdir(download_path + "\\DocxFiles")
    for file in filenames:
        if file.endswith(".doc"):
            os.remove(file)

    #  Remove existing RWS filled template files
    os.chdir(download_path + "\\TestTemplates")
    for filename in os.listdir(download_path + "\\TestTemplates"):
        if filename.startswith("RWS"):
            os.remove(filename)
    time.sleep(2)



def read_docx_files():
    files = glob.glob(download_path + "\\DocxFiles\\Copy*")
    os.chdir(download_path)
    for file in files:
        docx_handler = docx.Document(file)
        docx_tables = docx_handler.tables
        title = docx_tables[0].rows[0].cells[0].text
        jira_test_id = title.split()[0]
        if "Zephyr" in docx_tables[2].rows[2].cells[0].text:
            zephyr_tests = docx_tables[2].rows[2].cells[1]
        elif "Zephyr" in docx_tables[2].rows[3].cells[0].text:
            zephyr_tests = docx_tables[2].rows[3].cells[1]
        else:
            break
        zephyr_tests_table = zephyr_tests.tables

        zephyr_rows = zephyr_tests_table[0].rows  # get row id's
        zephyr_rows = zephyr_rows[1:]  # remove first cell from all rows (e.g. "Test Step", "Test Data", etc.)

        #  Test Steps
        list_of_test_steps = []
        for row in zephyr_rows:
            list_of_test_steps.append(row.cells[1].text)

        #  Test Conditions
        list_of_test_conditions = []
        for row in zephyr_rows:
            list_of_test_conditions.append(row.cells[2].text)

        #  Expected results
        list_of_exptected_results = []
        for row in zephyr_rows:
            list_of_exptected_results.append(row.cells[3].text)

        number_of_teststeps = zephyr_tests_table[0].rows[-1].cells[0].text
        file_save_path = download_path + "\\TestTemplates\\RWS-" + os.path.basename(file)

        rws_template = docx.Document(download_path + "\\SampleTestScripts1.docx")
        rws_table = rws_template.tables
        font = rws_template.styles['Normal'].font
        font.name = 'Calibri'
        paragraph = rws_template.styles['Normal'].paragraph_format
        paragraph.space_after = Pt(3)
        for x in range(1, int(number_of_teststeps)):
            rws_table[0].add_row()

        os.chdir(download_path)
        test_id = rws_table[0].rows[0].cells[2].paragraphs[0]
        test_id.add_run(jira_test_id)
        user_story = rws_table[0].rows[4].cells[2].paragraphs[0]
        user_story.add_run(jira_test_id)

        steps_only_table = rws_table[0].rows[6:]

        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[0].paragraphs[0].add_run(x + 1)
        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[1].paragraphs[0].add_run(list_of_test_steps[x])
        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[2].paragraphs[0].add_run(list_of_test_conditions[x])
        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[3].paragraphs[0].add_run(list_of_exptected_results[x])

        rws_template.save(file_save_path)


def move_files_to_epics():
    os.chdir(download_path + "\\TestTemplates")
    for filename, epic_name in epics_dictionary.items():
        if filename in os.path.exists(download_path + "\\TestTemplates\\" + filename):  # kinda unnecessary but whatever for now
            file_path = download_path + "\\Directories\\" + epic_name
            if filename.startswith("RWS") and not os.path.exists(download_path + epic_name):
                os.makedirs(file_path)
            shutil.move(download_path + "\\TestTemplates\\" + filename, file_path)



if __name__ == "__main__":

    # unittest.main()
    create_dir_hierarchy()
    move_doc_files()
    save_to_docx()
    # remove_doc_files()
    # read_docx_files()
    # move_files_to_epics()

