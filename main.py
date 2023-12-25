# -*- coding: utf-8 -*-
# Auther : SHL
# Date : 2023/12/25 9:37
# File : main.py
import time

from lib.common.BasePage import *

class feishu_auto:

    def __init__(self,web_url):
        # 浏览器DW路径
        self.application_path = os.path.abspath(os.getcwd())
        browser_path = os.path.join(self.application_path + "\\docker", "chromedriver.exe")
        print(browser_path)
        self.web = BasePage(browser_path)
        print(self.web)
        # 启动浏览器
        self.web.open_web(web_url)

    def test1(self):
        ''' 登录飞书 '''
        # 判断是否有弹窗
        pop_up = self.web.find_element(By.XPATH,"//div[@data-elem-id='PecYiJ9cU8']")
        if pop_up:
            pop_up.click()
        self.web.find_element(By.XPATH,"//a[contains(text(),'登录')]").click()

        if self.web.find(By.XPATH,"//div[contains(text(),'欢迎使用飞书')]"):
           pass
        else:
            self.web.find_element(By.XPATH,"//div[@class='switch-login-mode-box']").click()

        account = self.web.find_element(By.XPATH, "//input[@data-test='login-phone-input']")
        account.click()
        account.send_keys("15820407637")
        time.sleep(2)
        self.web.find_element(By.XPATH, "//span[contains(text(),'我已阅读并同意 ')]").click()
        self.web.find_element(By.XPATH, "//button[@data-test='login-phone-next-btn']").click()
        # 账号密码登录
        if self.web.find(By.XPATH, "//div[contains(text(),'输入密码')]"):
            pass
        else:
            self.web.find_element(By.XPATH, "//button[contains(text(),'密码登录')]").click()
        # 输入密码
        password = self.web.find_element(By.XPATH, "//input[@data-test='login-pwd-input']")
        password.click()
        password.send_keys("hyc13653864815")
        self.web.find_element(By.XPATH, "//button[@data-test='login-pwd-next-btn']").click()
        self.web.find_element(By.XPATH, "//span[contains(text(),'阿成')]").click()

        # 获取名称
        get_name = self.web.find_element(By.XPATH, "//button[@data-elem-id='SzcitVFaYB']")
        if get_name == "阿成":
            print("登录成功")
        else:
            print("登录失败")

    def test2(self):
        # 点击9点，进入消息
        self.web.find_element(By.XPATH,"//div[@class='headerExtra_productList']").click()
        self.web.find_element(By.XPATH, "//button[contains(text(),'消息')]").click()
    def test3(self):
        # 进入通讯录
        self.web.find_element(By.XPATH, "//section[@data-tip='tip-contacts']").click()

    def test4(self):
        # 进入聊天界面
        self.web.find_element(By.XPATH, "//div[@class='list_items']").click()
        self.web.find_elements(By.XPATH,"//button[@class='larkc-usercard__cta-button']")[0].click()

    def test5(self):
        # 进入通讯录
        self.web.find_element(By.XPATH, "//div[@class='lark-editor-wrap']").click()
        self.web.find_element(By.XPATH, "//div[@class='lark-editor-wrap']").send_keys("帅哥，今天天气很好，咱们一起去郊外游玩吧")

if __name__ == '__main__':
    demo = feishu_auto("https://www.feishu.cn/")
    demo.test1()








