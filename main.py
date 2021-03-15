import util
import os

from dotenv import load_dotenv
from airtest_selenium.proxy import WebChrome
from airtest.core.api import *

load_dotenv()

driver = WebChrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe')
driver.implicitly_wait(20)
target_url = os.getenv("WISDOM_URL")
driver.get(target_url)
driver.maximize_window()


def login_wisdom(account, passwd):
    driver.find_element_by_xpath("//input[@type='text']").send_keys(account)
    driver.find_element_by_xpath("//input[@type='password']").send_keys(passwd)
    driver.find_element_by_xpath("//input[@type='submit']").click()


def run():
    driver.find_element_by_class_name("select2-selection__arrow").click()
    lists = driver.find_element_by_class_name("select2-results__options")
    items = lists.find_elements_by_tag_name("li")

    for item in items:
        row_data = util.read_xlsx(sheet=util.parse_list_name(item.text))
        item.click()
        driver.find_element_by_xpath("//input[@type='submit']").click()
        driver.find_element_by_xpath("//a[@href='/wise/wiseadm/test.jsp']").click()

        driver.switch_to_new_tab()

        start = 3  # start with wisdom response
        response_collect = []

        for i, v in enumerate(row_data):
            if type(row_data[i]) == float:
                continue
            driver.find_element_by_xpath('//*[@id="message_body"]').send_keys(v)
            driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/form/div[1]/div[2]/button[1]').click()
            driver.get_screenshot_as_file("C:\\Users\\bingjiunchen\\Downloads\\snapshot\\response_%s.png" % str(i))
            msg = driver.find_elements_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]' % str(
                    start))
            start = start + 1
            for m in msg:
                # TODO: write in current excel file
                response_collect.append(m.text)

            if start % 2 == 0:
                start = start + 1
        driver.close()  # close current page
        break


if __name__ == '__main__':
    login_wisdom(os.getenv("WISDOM_ACCOUNT"), os.getenv("WISDOM_PASSWD"))
    run()
