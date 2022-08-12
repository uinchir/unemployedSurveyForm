# -*- coding:utf-8 -*-
__author__ = "残联——川"

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import selenium.common.exceptions as ex
from retrying import retry


class Config:
    file_name = "./data/tongyang_data.xlsx"
    driver_path = r'C:\Users\Administrator\.wdm\drivers\chromedriver\win32\103.0.5060\chromedriver.exe'  # 驱动的路径
    IP = "127.0.0.1:9527"  # 本地浏览器IP
    frame_3_id = "wjyzkIframe"  # 第三层frame的ID
    frame_4_name = "Opencaiji"  # 弹窗frame的name
    frame_5_id = "wjyzkdcIframe"  # 弹窗第二层frame的ID
    wait_time = 5  # 显式等待最大等待时间
    el_1 = "//*[@id='form']/table/tbody/tr[6]/td/a/span/span[1]"  # 第一个查询按钮路径
    el_2 = "/html/body/div[2]/div[2]/div/div[2]/div[2]/div[2]/table/tbody/tr/td[7]/div"  # 调查状况text路径
    el_3 = "//*[@id='query']/span/span[1]"  # “查看”按钮路径
    el_5 = "/html/body/div[2]/div[2]/table/tbody/tr[3]/td[2]/span[1]/span/a"  # ”生活自理能力“按钮路径
    el_6 = "/html/body/div[2]/div[2]/table/tbody/tr[3]/td[4]/span[1]/span/a"  # “工作意愿”按钮路径
    el_7 = "/html/body/div[2]/div[2]/table/tbody/tr[5]/td[2]/span[1]/span/a"  # “特殊情况”按钮路径
    el_last_1 = "/html/body/div[2]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[1]/td/div/a"
    el_last_2 = "/html/body/div[1]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[1]/td/div/a"  # 最后的弹窗关闭按钮路径
    el_abnormal_last = "//*[@id='main']/div[1]/div/table/t" \
                       "body/tr[2]/td[2]/div/table/tbody/tr[1]/td/div/a"  # 异常窗口的关闭按钮
    el_b1 = "/html/body/div[3]/div[2]/div/div/div[1]/div[2]"  # “就业培训”按钮路径
    el_b2 = "/html/body/div[3]/div[2]/div/div/div[2]/div/div[4]/div[1]/div[3]/a[2]"  # “综合查询”按钮路径
    el_b3 = "//*[@id='smz-easyui-accordion']/div[4]/div[2]/ul/li[1]"  # “残疾人”按钮路径
    el_b4 = "//*[@id='tzTab']/div[1]/div[3]/ul/li[6]/a/span[1]"  # “未就业状况管理”按钮路径


class Main(Config):
    def __init__(self):
        # 加载已打开网页
        self.my_options = Options()
        # self.my_options.add_argument("--headless")  # # => 为Chrome配置无头模式
        self.my_options.add_experimental_option("debuggerAddress", self.IP)
        self.browser = webdriver.Chrome(self.driver_path, options=self.my_options)
        # 显式等待初始设定
        self.wait = WebDriverWait(self.browser, self.wait_time)
        self.browser.set_page_load_timeout(self.wait_time)

    def find(self, by_style, content):
        """引用了显式等待的查找元素,并检测是否显示"""
        self.wait.until(lambda p: p.find_element(by_style, content).is_displayed())
        return self.browser.find_element(by_style, content)

    @retry(stop_max_attempt_number=10, wait_random_min=500, wait_random_max=1000)
    def load_web(self):
        """如果报错就重复运行，最多10次，间隔5000ms~6000ms"""
        print("进入页面")
        self.browser.refresh()  # 刷新页面
        self.find(By.XPATH, self.el_b1).click()
        self.find(By.XPATH, self.el_b2).click()
        self.find(By.XPATH, self.el_b3).click()
        self.browser.switch_to.frame(1)  # 跳转进第一层(索引--第二个)
        self.browser.switch_to.frame(0)  # 跳转进第二层（索引--第一个）
        self.find(By.XPATH, self.el_b4).click()  # 点击“未就业状况调查”按钮

    def switch_to_third_frame(self):
        self.browser.switch_to.default_content()
        self.browser.switch_to.frame(1)  # 跳转进第一层(索引--第二个)
        self.browser.switch_to.frame(0)  # 跳转进第二层(索引--第一个)
        self.browser.switch_to.frame(self.frame_3_id)  # 跳转进第三层iframe

    def input_data_and_search(self, pram):
        """输入身份证号并查询    pram:残疾人证号码"""
        input_box = self.find(By.ID, "idcard")  # 查找元素
        input_box.clear()  # 清除输入框
        input_box.send_keys(pram[0:18])  # 输入身份证号（残疾人证号码前18位）
        self.browser.execute_script("arguments[0].click();", self.find(By.XPATH, self.el_1))  # 查找查询按钮并点击

    def investigated_text(self, n, name):  # 未引用，此函数有问题，需要改进
        attempts = 1
        while attempts < 2:
            try:
                a = self.find(By.XPATH, self.el_2)
                print(f"{n} {name} {a.text}")
                return a.text
            except ex.StaleElementReferenceException:
                attempts += 1

    def tick_and_check(self, pram):
        """打勾并点击查看，pram:残疾证号码"""
        self.find(By.XPATH, f"//*[@value='{pram[0:18]}']").click()
        self.browser.execute_script("arguments[0].click();", self.find(By.XPATH, self.el_3))

    def jump_to_lay_frame(self):
        """切换到新跳出窗口，并点击进入’未就业状况调查‘。"""
        self.browser.switch_to.default_content()
        self.browser.switch_to.frame(self.frame_4_name)
        self.find(By.XPATH, "//span[text()='未就业状况调查']").click()
        self.browser.switch_to.frame(self.frame_5_id)

    def fill_form(self, pram1, pram2, pram3):
        """填表"""
        if pram3 != "请选择":
            self.find(By.XPATH, self.el_7).click()  # 查找“特殊情况”元素
            self.find(By.XPATH, f"//div[text()='{pram3}']").click()
        else:
            self.find(By.XPATH, self.el_5).click()  # 查找“生活自理能力”元素
            self.find(By.XPATH, f"//div[text()='{pram1}']").click()  # 按条件填入元素
            self.find(By.XPATH, self.el_6).click()  # 查找“就业意愿”元素
            self.find(By.XPATH, f"//div[text()='{pram2}']").click()  # 按条件填入元素
        self.find(By.XPATH, "//span[text()='保存']").click()  # 点保存

    def close_lay_frame(self):
        self.browser.switch_to.default_content()
        self.find(By.XPATH, "//button[text()='关闭']").click()
        self.browser.execute_script("arguments[0].click();", self.find(By.XPATH, self.el_last_1))

    def close_abnormal_frame(self):
        self.browser.switch_to.default_content()
        self.find(By.XPATH, self.el_abnormal_last).click()
