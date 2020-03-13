from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import getpass
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import difflib
import glob
import re
import win32com.client as win32
from win32com.client import constants
import docx
import fnmatch
import unittest


download_path = "C:\\PycharmProjects\\Jira_scripts\\Downloads";
user_login = "abc"
user_password = "abc"


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

        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "order-options")))
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "issue-link-key")))
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@class='issue-link-key']")))

    def test_3_download_tests(self):
        driver = self.driver
        wait = WebDriverWait(driver, 10)
        print("it starts now")
        wait.until(EC.element_to_be_clickable((By.XPATH, "(//*[@class='issue-link-key'])[1]")))
        tests_list = driver.find_elements_by_class_name("issue-link-key")
        print("Number of files: ", len(tests_list))   # looks obsolete, won't work for > 50 files as it checks only 1 page
        for test_element in tests_list:
            test_element.click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='viewissue-export']")))
            driver.find_element_by_xpath("//*[@id='viewissue-export']").click()
            wait.until(EC.element_to_be_clickable((By.XPATH,"//*[@id='jira.issueviews:issue-word']")))
            driver.find_element_by_xpath("//*[@id='jira.issueviews:issue-word']").click()
        time.sleep(3)
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
def save_as_docx(paths_to_files):  # Github convertion solution
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


def save_to_docx():
    paths_to_file = os.listdir(r"C:\PycharmProjects\Jira_scripts\Downloads\DocFiles")
    paths_to_files = [download_path + "\\DocFiles\\" + path for path in paths_to_file]
    for path in paths_to_files:
        save_as_docx(path)

def remove_doc_files():
    os.chdir(download_path + "\\DocFiles")
    filenames = os.listdir(download_path + "\\DocFiles")
    for file in filenames:
        if file.endswith(".doc"):
            os.remove(file)
    time.sleep(3)


def read_docx_files():
    files = glob.glob(download_path + "\\DocFiles\\Copy*")
    os.chdir(download_path)
    for file in files:
        docx_handler = docx.Document(file)
        docx_tables = docx_handler.tables
        title = docx_tables[0].rows[0].cells[0].text
        print(title)
        title_split = title.split()
        print(title_split)
        jira_test_id = title_split[0]
        print(jira_test_id)
        zephyr_teststeps = docx_tables[2].rows[2].cells[0].text
        print(zephyr_teststeps)
        if zephyr_teststeps == "Zephyr Teststep:":
            zephyr_tests = docx_tables[2].rows[2].cells[1]
        else:
            zephyr_tests = docx_tables[2].rows[3].cells[1]
        zephyr_tests_table = zephyr_tests.tables
        number_of_teststeps = zephyr_tests_table[0].rows[-1].cells[0].text  # teststep 1???

        print(number_of_teststeps)

        tab_zef = zephyr_tests_table[0].rows
        tab_zef = tab_zef[1:]  # obcinanie tytulu test stepow

        lista_stepow = []
        for row in tab_zef:
            lista_stepow.append(row.cells[1].text)

        lista_stepow[0]
        #Test Condition
        lista_TestCondition=[]

        for row in tab_zef:
            lista_TestCondition.append(row.cells[2].text)

     lista_ExpResult=[]
#     for row in tab_zef:
#         lista_ExpResult.append(row.cells[3].text)
#     if step1 == "2":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts.docx")
#         table1 = doc1.tables
#     elif step1 == "3":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts3.docx")
#         table1 = doc1.tables
#     elif step1 == "4":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts4.docx")
#         table1 = doc1.tables
#     elif step1 == "5":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts5.docx")
#         table1 = doc1.tables
#     elif step1 == "6":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts6.docx")
#         table1 = doc1.tables
#     elif step1 == "7":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts7.docx")
#         table1 = doc1.tables
#     elif step1 == "8":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts8.docx")
#         table1 = doc1.tables
#     elif step1 == "9":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts9.docx")
#         table1 = doc1.tables
#     elif step1 == "10":
#         doc1 = docx.Document("C:\\Users\\jpiet\\Downloads\\SampleTestScripts10.docx")
#         table1 = doc1.tables
#
#     idTest1 = table1[0].rows[0].cells[2]
#     add_idScenario = idTest1.add_paragraph(idTest)
#     user_story = table1[0].rows[4].cells[2]
#     add_UserStory = user_story.add_paragraph(idTest)
#
#     tabela_template = table1[0].rows
#
#     tabela_template = tabela_template[6:]
#
#     lista_TestStepow = []
#     for row in tabela_template:
#         lista_TestStepow.append(row.cells[1])
#
#
#         #Test Stepy
#     index = 0
#     for element in lista_TestStepow:
#         element.add_paragraph(lista_stepow[index])
#         index +=1
#
#         #Test Condition
#     lista_TestCondition1 =[]
#     for row in tabela_template:
#         lista_TestCondition1.append(row.cells[2])
#     index = 0
#     for element in lista_TestCondition1:
#         element.add_paragraph(lista_TestCondition[index])
#         index +=1
#
#         #Dodawanie test Condition
#
#     lista_ExpResult1 =[]
#     for row in tabela_template:
#         lista_ExpResult1.append(row.cells[3])
#
#     index = 0
#     for element in lista_ExpResult1:
#         element.add_paragraph(lista_ExpResult[index])
#         index +=1
#
#         #Dodawanie Excepted Result
#     doc1.save(file)


if __name__ == "__main__":

    # unittest.main()
    # paths = glob.glob('C:/PycharmProjects/Jira_scripts/Downloads/DocFiles/**/*.doc', recursive=True)
    # paths = os.listdir(r"C:/PycharmProjects/Jira_scripts/Downloads/DocFiles")
    # paths_to_files = os.listdir(r"C:/PycharmProjects/Jira_scripts/Downloads/DocFiles")
    # save_to_docx()
    # remove_doc_files()
    # read_docx_files()
    read_docx_files()


