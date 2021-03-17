import util

from dotenv import load_dotenv
from airtest_selenium.proxy import WebChrome
from airtest.core.api import *
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException

load_dotenv()

driver = WebChrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe')
driver.implicitly_wait(20)
target_url = os.getenv("WISDOM_URL")
driver.get(target_url)
driver.maximize_window()

sheets_list = ['金好運盤點表展開']


def init():
    for s in sheets_list:
        row_data = util.read_xlsx(s, "標準答案")
        new_row_list = [{i: v.replace("\n", "")} for i, v in enumerate(row_data) if type(v) != float]
        for n in new_row_list:
            for k, v in n.items():
                row_data[k] = v
        util.write_xlsx(row_data, s, col_name="標準答案", col_pos="K1")


def login_wisdom(account, passwd):
    driver.find_element_by_xpath("//input[@type='text']").send_keys(account)
    driver.find_element_by_xpath("//input[@type='password']").send_keys(passwd)
    driver.find_element_by_xpath("//input[@type='submit']").click()


def run():
    for idx, s in enumerate(sheets_list):
        driver.find_element_by_class_name("select2-selection__arrow").click()
        select_lists = driver.find_element_by_class_name("select2-results__options")
        items = select_lists.find_elements_by_tag_name("li")
        items[idx + 2].click()

        try:
            driver.find_element_by_xpath("//input[@type='submit']").click()
        except ElementNotInteractableException:
            print("cannot find submit button")

        row_question = util.read_xlsx(sheet=s, row_name="品檢問題填寫處")
        row_sn = util.read_xlsx(sheet=s, row_name="知識編號")

        driver.find_element_by_xpath("//a[@href='/wise/wiseadm/test.jsp']").click()

        driver.switch_to_new_tab()

        count = 0  # count knowledge number ten for a round
        start = 3  # start with wisdom response
        response_collect = []
        compare_result = []

        small_msg = ""
        tell_msg = ""
        sn_msg = ""
        texts = []

        for i, v in enumerate(row_question):
            count = count + 1

            if type(row_question[i]) == float:
                continue
            driver.find_element_by_xpath('//*[@id="message_body"]').send_keys(v)
            driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/form/div[1]/div[2]/button[1]').click()

            driver.get_screenshot_as_file(os.getenv("SCREENSHOT") + "{0}_{1}.png".format(str(row_sn[i]), str(count)))

            try:
                texts = driver.find_elements_by_xpath(
                    '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]' % str(
                        start))

                small_msg = driver.find_element_by_xpath(
                    '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/small' % str(
                        start)).text
                tell_msg = driver.find_element_by_xpath(
                    '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/button' % str(
                        start)).text
                sn_msg = driver.find_element_by_xpath(
                    '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/a/small' % str(
                        start)).text

            except NoSuchElementException:
                print("tests or msg not found: %s" % v)

            for t in texts:
                msg = t.text
                msg = msg.replace(small_msg, "")
                msg = msg.replace(sn_msg, "")
                msg = msg.replace(tell_msg, "")
                msg = msg.replace("\n", "")
                response_collect.append(msg)
                if util.compare_response_with_answer(msg, sheet_name=s, row_name="標準答案", row_index=i):
                    try:
                        # 點讚
                        driver.find_element_by_xpath(
                            '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/div/span[1]' % str(
                                start)).click()

                    except NoSuchElementException:
                        print("cannot be like: %s" % v)
                    compare_result.append("是")
                else:
                    compare_result.append("否")

            start = start + 1

            if start % 2 == 0:
                start = start + 1
            if count == 10:
                count = 0  # reset count

    util.write_xlsx(row_list=response_collect, sheet_name=s, col_name="回答", col_pos="N1")

    util.write_xlsx(row_list=compare_result, sheet_name=s, col_name="比較結果", col_pos="O1")

    # driver.close()  # close current page
    driver.switch_to_previous_tab()
    driver.back()


if __name__ == '__main__':
    init()
    login_wisdom(os.getenv("WISDOM_ACCOUNT"), os.getenv("WISDOM_PASSWD"))
    run()

    # row = util.read_xlsx(sheet="明星3缺1盤點表展開", row_name="回答")
    # for i in range(1, 21):
    #     print(util.compare_response_with_answer(require=row[i], sheet_name="明星3缺1盤點表展開", row_name="標準答案", row_index=i))

