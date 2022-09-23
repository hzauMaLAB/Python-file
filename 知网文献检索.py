import time
import xlrd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from numpy import *

class cnki():
    def __init__(self, save_Path):
        self.save_Path = 'E:\\cnki\\{}'.format(save_Path)    # 创建保存文件夹
        # 定义一个无界面的浏览器
        self.options = webdriver.ChromeOptions()
        self.prefs = {'profile.default_content_settings.popups': 0,
                 'profile.managed_default_content_settings.images': 2,
                 'download.default_directory': self.save_Path}
        self.options.add_experimental_option('prefs', self.prefs)
        self.options.add_argument('--headless')
        self.options.add_argument('--disable-gpu')
        s = Service("chromedriver.exe")              #安装此执行程序在当前文件夹
        self.browser = webdriver.Chrome(service=s)
        # 300s无响应就down掉
        self.wait = WebDriverWait(self.browser, 300)
        # 定义窗口最大化
        self.browser.maximize_window()

    def getHtml(self, zy_mc, zy_str):
        try:
            self.browser.get('https://kns.cnki.net/kns8/AdvSearch?dbprefix=SCDB&&crossDbcodes=CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCJFN%2CCCJD')
            # PDF那个按钮
            self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//ul[@class="search-classify-menu"]/li[4]'))).click()
            input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//textarea[@class="textarea-major ac_input"]'))
            )
            # 清除里面的数字
            input.clear()
            input.send_keys(zy_str)
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//input[@class="btn-search"]'))
            ).click()
            time.sleep(3)
            total = self.browser.find_element(By.XPATH, '//*[@id="countPageDiv"]/span/em').text
            print(zy_mc+"一共有"+total+"条数据")
            page = int(total)//20+1   #知网里面一页为20个文献
            a = 1
            for p in range(page):
                for i in range(1, 20):
                    link = self.browser.find_element(By.XPATH,'//*[@id="gridTable"]/table/tbody/tr[%d]/td[2]/a' % i)
                    print(link)
                    flag1 = self.isElementExist('//*[@id="gridTable"]/table/tbody/tr[%d]/td[2]' % i)
                    if flag1:
                        self.browser.execute_script("arguments[0].scrollIntoView();", link)
                        time.sleep(3)
                        actions = ActionChains(self.browser)
                        actions.move_to_element(link)
                        actions.click(link)
                        actions.perform()
                        time.sleep(10)
                        windows = self.browser.window_handles
                        self.browser.switch_to.window(windows[-1])
                        time.sleep(3)
                        try:
                            flag2 = self.isElementExist('//*[@id="pdfDown"]')
                            if flag2:
                                pdf = self.browser.find_element(By.XPATH,'//*[@id="pdfDown"]')
                                self.browser.execute_script("arguments[0].scrollIntoView();", pdf)
                                time.sleep(3)
                                self.wait.until(EC.presence_of_element_located(
                                    (By.XPATH, '//*[@id="pdfDown"]'))).click()
                            else:
                                pass
                        except Exception as e:
                            print(e)
                        time.sleep(10)
                        self.browser.close()
                        time.sleep(5)
                        self.browser.switch_to.window(windows[0])
                        print("-----正在爬取--" + zy_mc + '--药品的第' + str(int(p)+1) + '页' + str(a) + "条数据------")
                        a = a + 1
                    else:
                        break
                flag3 = self.isElementExist('//*[@id="PageNext"]')
                # 点击下一页
                if flag3:
                    time.sleep(10)
                    next_page = self.browser.find_element(By.XPATH,'//*[@id="PageNext"]')
                    self.browser.execute_script("arguments[0].scrollIntoView();", next_page)
                    self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="PageNext"]'))).click()
                    time.sleep(5)
                else:
                    break
        except Exception as e:
            print(e)
    # 该方法用来确认元素是否存在，如果存在返回flag=true，否则返回false
    def isElementExist(self, element):
        flag = True
        try:
            self.browser.find_element(By.XPATH,element)
            return flag
        except:
            flag = False
            return flag

def main():
    try:
        data = xlrd.open_workbook('药名.xlsx')
        table = data.sheets()[0]       # 获取第一页
        zymc_lists = table.col_values(0)
        zcy_lists = table.col_values(1)
        for i in range(0, 1700000):
            str1 = zymc_lists[i]
            str2 = zcy_lists[i]
            print(str2)
            print(30 * '==' + str(i) + 30 * '==')
            c = cnki(str1)
            time.sleep(1)
            c.getHtml(str1, str2)
            time.sleep(10)

    except Exception as e:
        print(e)
    finally:
        pass


if __name__ == '__main__':
    main()
