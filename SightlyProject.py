import datetime
import time
import xlrd
import os
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from itertools import zip_longest
from pathlib import Path

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
driver.get("https://staging-newtargetview.sightly.com/")
driver.implicitly_wait(10)

driver.find_element_by_xpath("//input[@placeholder='Your email address']").clear()
driver.find_element_by_xpath("//input[@placeholder='Your email address']").send_keys("nick@sightly.com")
driver.find_element_by_xpath("//input[@type='password']").clear()
driver.find_element_by_xpath("//input[@type='password']").send_keys("a")
driver.find_element_by_xpath("//button[@class='login-button']").click()
w = WebDriverWait(driver, 7)
w.until(EC.presence_of_element_located((By.ID, "header-reports")))
print("login success")

driver.find_element_by_id("header-reports").click()
checkbox_item = "//div[@id='reports-list-main-content']/div[2]/div[3]/div[2]/div[1]/input[1]"
driver.find_element_by_xpath(checkbox_item).click()
time.sleep(2)
if driver.find_element_by_xpath(checkbox_item).is_selected():
    assert True
driver.find_element_by_xpath("//button[@class='create-report']").click()

driver.implicitly_wait(10)
driver.find_element_by_xpath("//*[@id='performanceDetailReportOption']/div/div[1]/input").click()


def select_values(element, value):
    selection = Select(element)
    selection.select_by_visible_text(value)
    time.sleep(0.8)


ele_grouping = driver.find_element_by_xpath("//*[@id='inputGroupSummaryContainer']/div/div[1]/div[2]/select")
ele_cost = driver.find_element_by_xpath("//*[@id='inputGroupSummaryContainer']/div/div[2]/div[2]/select")
ele_granularity = driver.find_element_by_xpath("//*[@id='inputGroupSummaryContainer']/div/div[3]/div[2]/select")
ele_additionalColumns = driver.find_element_by_xpath("//*[@id='inputGroupSummaryContainer']/div/div[4]/div[2]/select")
ele_timePeriod = driver.find_element_by_xpath("//*[@id='inputGroupSummaryContainer']/div/div[5]/div[2]/select")

select_values(ele_grouping, "Campaign")
select_values(ele_cost, "All")
select_values(ele_granularity, "Summary")
select_values(ele_additionalColumns, "None")
select_values(ele_timePeriod, "All Time")

driver.find_element_by_xpath("//strong[contains(text(),'Run Reports')]").click()
time.sleep(5)

downloads_path = str(Path.home() / "Downloads")


def getfilename():
    for file in os.listdir(downloads_path):
        if file.endswith(".xlsx") and file.startswith("PerformanceDetail-Campaign-"):
            file = (os.path.join(downloads_path, file))
            print(file)
            return file


downloadedFile = xlrd.open_workbook(getfilename())
'''Changing the Current Working Directory to where the python script is located'''
os.chdir(os.path.dirname(os.path.realpath(__file__)))
my_file = Path.cwd() / "Data/Data.xlsx"
hardcodedFile = xlrd.open_workbook(my_file)

downloadedSheet = downloadedFile.sheet_by_index(0)
hardcodedSheet = hardcodedFile.sheet_by_index(0)

'''Read Excel and print to output'''
for row in range(downloadedSheet.nrows):  # loop every row
    for col in range(downloadedSheet.ncols):  # in that row iterate every column
        if col in [0, 1] and row not in [0, 1, 2, 3, 4]:
            output = downloadedSheet.cell_value(row, col)
            convert_date = xlrd.xldate_as_tuple(output, downloadedFile.datemode)
            print_date = datetime.datetime(*convert_date).strftime("%m/%d/%y")
            print(print_date, end='')
        else:
            print(downloadedSheet.cell_value(row, col), end='')
        print('\t', end='')
    print()

'''Validate data downloaded to the hardcoded data'''
for rownum in range(2, max(downloadedSheet.nrows, hardcodedSheet.nrows)):
    if rownum < downloadedSheet.nrows:
        row_downloadedSheet = downloadedSheet.row_values(rownum)
        row_hardcodedSheet = hardcodedSheet.row_values(rownum)

        for colnum, (c1, c2) in enumerate(zip_longest(row_downloadedSheet, row_hardcodedSheet)):
            if c1 != c2:
                print("Row {} Col {} - {} != {}".format(rownum + 1, colnum + 1, c1, c2))
    else:
        print("Row {} missing".format(rownum + 1))


driver.find_element_by_xpath("//div[@id='modal-container']/div[1]/div[3]").click()
driver.implicitly_wait(3)
driver.find_element_by_id("header-orders").click()
time.sleep(2)
driver.find_element_by_xpath("//*[@id='nav-link-container']/div[5]").click()
print("******************************************")
print("Logout successfully")


driver.close()
driver.quit()
