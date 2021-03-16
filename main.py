import util

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
        title = item.text
        row_question = util.read_xlsx(sheet=util.parse_list_name(item.text), row_name="品檢問題填寫處")
        row_sn = util.read_xlsx(sheet=util.parse_list_name(item.text), row_name="知識編號")
        item.click()
        driver.find_element_by_xpath("//input[@type='submit']").click()
        driver.find_element_by_xpath("//a[@href='/wise/wiseadm/test.jsp']").click()

        driver.switch_to_new_tab()

        count = 0  # count knowledge number ten for a round
        start = 3  # start with wisdom response
        response_collect = []

        for i, v in enumerate(row_question):
            count = count + 1

            if type(row_question[i]) == float:
                continue
            driver.find_element_by_xpath('//*[@id="message_body"]').send_keys(v)
            driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/form/div[1]/div[2]/button[1]').click()
            time.sleep(0.5)

            driver.get_screenshot_as_file(os.getenv("SCREENSHOT") + "{0}_{1}.png".format(str(row_sn[i]), str(count)))

            driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/div[2]/span[1]' % str(
                    start)).click()

            texts = driver.find_elements_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]' % str(
                    start))

            small_msg = driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/small' % str(
                    start)).text
            sn_msg = driver.find_element_by_xpath(
                '//*[@id="content-wrapper"]/div/div[2]/div[1]/div/div/div/div/div/div/div/div[1]/ul/li[%s]/div[3]/a/small' % str(
                    start)).text

            start = start + 1
            for t in texts:
                msg = t.text
                msg = msg.replace(small_msg, "")
                msg = msg.replace(sn_msg, "")
                response_collect.append(msg)

            if start % 2 == 0:
                start = start + 1
            if count == 10:
                count = 0  # reset count

        util.write_xlsx(row_list=response_collect, sheet_name=util.parse_list_name(title), col_name="回答",
                        col_pos="N1")
        # driver.close()  # close current page
        driver.switch_to_previous_tab()
        driver.back()
        # break


if __name__ == '__main__':
    login_wisdom(os.getenv("WISDOM_ACCOUNT"), os.getenv("WISDOM_PASSWD"))
    run()
