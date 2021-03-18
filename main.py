import util

from dotenv import load_dotenv
from airtest_selenium.proxy import WebChrome
from airtest.core.api import *
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException

load_dotenv()

split_keyword = "盤點表展開"


def init_excel(worksheet):
    row_data = util.read_xlsx(worksheet, "標準答案")
    new_row_list = [{i: v.replace("\n", "")} for i, v in enumerate(row_data) if type(v) != float]
    for n in new_row_list:
        for k, v in n.items():
            row_data[k] = v
    util.write_xlsx(row_data, worksheet, col_name="標準答案", col_pos="K1")


class App:
    def __init__(self, driver, account, passwd):
        self.driver = driver
        self.account = account
        self.passwd = passwd

        driver.implicitly_wait(20)
        driver.get(os.getenv("WISDOM_URL"))
        driver.maximize_window()

    def login_wisdom(self):
        self.driver.find_element_by_xpath("//input[@type='text']").send_keys(self.account)
        self.driver.find_element_by_xpath("//input[@type='password']").send_keys(self.passwd)
        self.driver.find_element_by_xpath("//input[@type='submit']").click()

    def entry_wisdom_page(self, worksheet_name):
        self.driver.find_element_by_class_name("select2-selection__arrow").click()
        select_list = self.driver.find_element_by_class_name("select2-results__options")
        select_options = select_list.find_elements_by_tag_name("li")

        split_ws_name = worksheet_name.split(split_keyword)
        for option in select_options:
            if option.text == "鈊象電子-{}(null)".format(split_ws_name[0]):
                option.click()
        try:
            self.driver.find_element_by_xpath("//input[@type='submit']").click()
        except ElementNotInteractableException:
            print("cannot find submit button")

        self.driver.find_element_by_xpath("//a[@href='/wise/wiseadm/test.jsp']").click()
        self.driver.switch_to_new_tab()

    def __handle_fetch_text(self, count):
        msg = ""
        try:
            texts = self.driver.find_elements_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]' % str(
                    count))
            small_msg = self.driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/small' % str(
                    count)).text
            tell_msg = self.driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/button' % str(
                    count)).text
            sn_msg = self.driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/a/small' % str(
                    count)).text

            for t in texts:
                msg = t.text
                msg = msg.replace(small_msg, "")
                msg = msg.replace(sn_msg, "")
                msg = msg.replace(tell_msg, "")
                msg = msg.replace("\n", "")

        except NoSuchElementException:
            print("tests or msg not found: %s" % count)
        finally:
            return msg

    def run(self, worksheet_name, row_data):
        screenshot_sn = []
        msg_collection = []
        correctness_collection = []

        start = 3

        for i, r in enumerate(row_data):
            self.driver.find_element_by_xpath('//*[@id="message_body"]').send_keys(r)
            self.driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/form/div[1]/div[2]/button[1]').click()

            sn = "{0}_{1}.png".format(worksheet_name, str(i))
            screenshot_sn.append(sn)
            time.sleep(0.5)
            self.driver.get_screenshot_as_file(os.getenv("SCREENSHOT") + sn)

            message = self.__handle_fetch_text(count=start)
            msg_collection.append(message)

            if util.compare_response_with_answer(require=message, sheet_name=worksheet_name, row_name="標準答案",
                                                 row_index=i):
                # click like
                self.driver.find_element_by_xpath(
                    '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/div/span[1]' % str(
                        start)).click()
                correctness_collection.append("是")
            else:
                # click dislike
                correctness_collection.append("否")

            start += 1
            if start % 2 == 0:
                start += 1

        util.write_xlsx(row_list=msg_collection, sheet_name=worksheet_name, col_name="回答", col_pos="N1")
        util.write_xlsx(row_list=correctness_collection, sheet_name=worksheet_name, col_name="比較結果", col_pos="O1")
        util.write_xlsx(row_list=screenshot_sn, sheet_name=worksheet_name, col_name="截圖流水號", col_pos="P1")

    # driver.close()  # close current page
    # driver.switch_to_previous_tab()
    # driver.back()


if __name__ == '__main__':

    app = App(driver=
              WebChrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe'),
              account=os.getenv("WISDOM_ACCOUNT"),
              passwd=os.getenv("WISDOM_PASSWD"))

    app.login_wisdom()

    for i in ['金好運盤點表展開']:
        init_excel(worksheet=i)

        app.entry_wisdom_page(worksheet_name=i)

        data = [q for q in util.read_xlsx(sheet="金好運盤點表展開", row_name="品檢問題填寫處") if type(q) != float]

        app.run(worksheet_name=i, row_data=data)
