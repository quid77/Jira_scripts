from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import getpass
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import difflib
from docx.shared import Inches, Pt
import glob
import re
import win32com.client as win32
from win32com.client import constants
import docx
import fnmatch
import unittest


download_path = "C:\\PycharmProjects\\Jira_scripts\\Downloads";
user_login = ""
user_password = ""


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
                    try:  # DOuble space, ", break the script
                        test_element_title = test_element.find_element_by_xpath(".//span[@class='issue-link-summary']").text
                        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[text()[contains(.,\"%s\")]]" % test_element_title)))
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


def create_folders():
    for filename in os.listdir(download_path):
        if filename.startswith("SSMWE") and not os.path.exists(download_path + filename):
            file_path = download_path + "\\Directories\\" + filename
            os.makedirs(file_path[:-4])


def rename_files():
    os.chdir("C:\\PycharmProjects\\Jira_scripts\\Downloads")
    for filename in os.listdir(download_path):
        if filename.startswith("SSMWE"):
            os.rename(filename, filename.replace("SSMWE-", "Copy-"))


def move_files():
    if not os.path.exists(download_path + "\\DocFiles"):
        os.makedev(download_path + "\\DocFiles")
    # Create list of paths to .doc files
    for filename in os.listdir(download_path):
        if filename.endswith(".doc"):
            file_path = download_path + "\\" + filename
            shutil.move(file_path, download_path + "\\DocFiles")


# Create list of paths to .doc files
# paths_to_files = os.listdir(r"C:/PycharmProjects/Jira_scripts/Downloads/DocFiles")

# this function isn't standalone, use save_to_docx instead
def save_as_docx(paths_to_files):  # Github conversion solution
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(paths_to_files)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(paths_to_files)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)


def save_to_docx():  # chceck obs if it doesnt work
    file_names = os.listdir(r"C:\PycharmProjects\Jira_scripts\Downloads\DocFiles")
    paths_to_files = [download_path + "\\DocFiles\\" + name for name in file_names]
    for path in paths_to_files:
        save_as_docx(path)

def remove_doc_files():
    os.chdir(download_path + "\\DocFiles")
    filenames = os.listdir(download_path + "\\DocFiles")
    for file in filenames:
        if file.endswith(".doc"):
            os.remove(file)

    #  Remove existing RWS filled-template files
    os.chdir(download_path + "\\TestTemplates")
    for filename in os.listdir(download_path + "\\TestTemplates"):
        if filename.startswith("RWS"):
            os.remove(filename)
    time.sleep(2)


def read_docx_files():
    files = glob.glob(download_path + "\\DocFiles\\Copy*")
    os.chdir(download_path)
    for file in files:
        docx_handler = docx.Document(file)
        docx_tables = docx_handler.tables
        title = docx_tables[0].rows[0].cells[0].text
        jira_test_id = title.split()[0]
        zephyr_teststeps = docx_tables[2].rows[2].cells[0].text
        if "Zephyr" in zephyr_teststeps:
            zephyr_tests = docx_tables[2].rows[2].cells[1]
        else:
            zephyr_tests = docx_tables[2].rows[3].cells[1]
        zephyr_tests_table = zephyr_tests.tables
        #
        # number_of_teststeps = zephyr_tests_table[0].rows[-1].cells[0].text
        # print(number_of_teststeps)

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
            steps_only_table[x].cells[1].paragraphs[0].add_run(list_of_test_steps[x])
        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[2].paragraphs[0].add_run(list_of_test_conditions[x])
        for x in range(0, int(number_of_teststeps)):
            steps_only_table[x].cells[3].paragraphs[0].add_run(list_of_exptected_results[x])

        rws_template.save(file_save_path)


if __name__ == "__main__":

    unittest.main()
    # save_to_docx()
    # remove_doc_files()
    # read_docx_files()
    # read_docx_files()


