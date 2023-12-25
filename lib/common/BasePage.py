# -*- coding: utf-8 -*-
# Auther : SHL
# Date : 2023/8/24 10:29
# File : BasePage.py
import time,os
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
class BasePage(object):
    def __init__(self,driver_path=None,url=None):
        if driver_path:
            self.driver_path = driver_path
        else:
            self.driver_path = "../../docker/chromedriver.exe"
        self.url = url
        # 先启动webdricer插件  "../../docker/chromedriver.exe"
        self.driver = webdriver.Chrome(self.driver_path)
        self.wait = WebDriverWait(self.driver,10)
    def open_web(self,url):

        # 打开站点
        # 浏览器进入百度网站
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.find_elements()

    def find_elements(self,*loc):
        find_elements = self.driver.find_elements(*loc)
        return find_elements

    def find(self,*loc):
        # find = self.wait.until(EC.visibility_of(self.driver.find_element(*loc)))
        # return find
        try:
            print(self.wait.until(EC.visibility_of(self.driver.find_element(*loc))))
            # WebDriverWait(self.driver,3,0.1).until(EC.visibility_of_element_located(loc))
            find = self.wait.until(EC.visibility_of(self.driver.find_element(*loc)))
            self.driver.execute_script("arguments[0].scrollIntoView();", find)
            # print(self.wait.until(EC.visibility_of_element_located(self.driver.find_element(*loc))))
            return True
        except:
            return False

    def finds(self,*loc):
        try:
            print(self.wait.until(EC.visibility_of(self.driver.find_elements(*loc))))
            # WebDriverWait(self.driver,3,0.1).until(EC.visibility_of_element_located(loc))
            find = self.wait.until(EC.visibility_of(self.driver.find_elements(*loc)))
            self.driver.execute_script("arguments[0].scrollIntoView();", find)
            # print(self.wait.until(EC.visibility_of_element_located(self.driver.find_element(*loc))))
            return True
        except:
            return False

    def find_element(self,*loc):
        # 网页完全加载完成后，开始定位元素,返回一个元素
        try:
            time.sleep(2)
            find_element = self.wait.until(EC.visibility_of(self.driver.find_element(*loc)))
            self.driver.execute_script("arguments[0].scrollIntoView();", find_element)
            # find_element = WebDriverWait(self.driver,20,1).until(EC.visibility_of_element_located(loc))

            return find_element
        except:
            lt = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
            print(lt,'-找不到元素：',*loc)

            print("再来找一次：",*loc)
            try:
                find_element = self.wait.until(EC.visibility_of(self.driver.find_element(*loc)))
                self.driver.execute_script("arguments[0].scrollIntoView();", find_element)
                return find_element
            except:
                lt = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
                print(lt, '-依然未不到', *loc)
            time.sleep(1)

    def find_all_element(self,*loc):
        # 网页完全加载完成后，开始定位元素,返回一个list列表
        try:
            win = self.driver.current_window_handle
            print(win)
            all_handles = self.driver.window_handles
            self.driver.switch_to.window(all_handles[-1])
            # 刷新当前页面
            self.driver.refresh()
            find_element = self.wait.until(EC.visibility_of_all_elements_located(self.driver.find_element(*loc)))
            self.driver.execute_script("arguments[0].scrollIntoView();", find_element)
            # find_element = WebDriverWait(self.driver,20,1).until(EC.visibility_of_element_located(loc))

            return find_element
        except:
            lt = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
            print(lt,'-找不到',*loc)
            print("再来找一次：",*loc)
            try:
                find_element = self.wait.until(EC.visibility_of_all_elements_located(self.driver.find_element(*loc)))
                self.driver.execute_script("arguments[0].scrollIntoView();", find_element)
                return find_element
            except:
                lt = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
                print(lt, '-依然未不到', *loc)
            time.sleep(1)



    ## 网上  # *loc 代表任意数量的位置参数
    def findElement(self, *loc):
         """
         查找单一元素
         :param loc:
         :return:
         """
         try:
             find_element = WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(loc))
             # log.logger.info('The page of %s had already find the element %s'%(self,loc))
             # return self.driver.find_element(*loc)
         except Exception as e:
             lt = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
             print(lt, '-依然未不到', *loc)
         else:
             return find_element

    def inputValue(self, inputBox, value):
        """
        后期修改其他页面直接调用这个函数
        :param inputBox:
        :param value:
        :return:
        """
        inputB = self.findElement(*inputBox)

        try:
            inputB.clear()
            inputB.send_keys(value)
        except Exception as e:

            raise e
        else:
            print('inputValue:[%s] is receiveing value [%s]' % (inputBox, value) )
            # 获取元素数据
    def getValue(self, *loc):
        """
        获取text，或value值
        :param loc:
        :return:
        """
        element = self.findElement(*loc)
        try:
             value = element.text  #return value
        except Exception:
             #element = self.find_element_re(*loc) # 2018.09.21 for log
             value = element.get_attribute('value')
             print('reading the element [%s] value [%s]' % (loc, value))
             return value
        except:
             raise Exception
        else:
             return value

    # 执行js脚本
    def jScript(self,src):
        """

        :param src:
        :return:
        """
        try:
            self.driver.execute_script(src)
        except Exception as e:
            raise e
        else:
            print('execute js script [%s] successed ' %src)


    # 截图
    def saveScreenShot(self, filename):
        """

        :param filename:
        :return:
        """
        list_value = []

        list = filename.split('.')
        for value in list:
            list_value.append(value)
        if list_value[1] == 'png' or list_value[1] == 'jpg' or list_value[1] == 'PNG' or list_value[1] == 'JPG':
            if 'fail' in list_value[0].split('_'):
                 try:
                     self.driver.save_screenshot(os.path.join(self.driver_path, filename))
                 except Exception:
                     print('save screenshot failed !')
                 else:
                     print('the file [%s]  save screenshot successed under [%s]' % (filename, self.driver_path))
            elif 'pass' in list_value[0]:
                try:
                    self.driver.save_screenshot(os.path.join(self.driver_path, filename))
                except Exception:
                    print('save screenshot failed !')
                else:
                    print(
                     'the file [%s]  save screenshot successed under [%s]' % (filename,self.driver_path))
            else:
                print('save screenshot failed due to [%s] format incorrect' %filename)
        else:
            print('the file name of [%s] format incorrect cause save screenshot failed, please check!' % filename)








if __name__ == '__main__':
    web = BasePage("../../docker/chromedriver.exe")
    web.open_web("https://cowork.lenovo.com/departments/quality/Lists/Geo%20Packaging/Item/editifs.aspx?List=21d7a9c7%2D272e%2D4af4%2Dab7f%2D1398a1369718&ID=8238&Source=https%3A%2F%2Fcowork%2Elenovo%2Ecom%2Fdepartments%2Fquality%2FSitePages%2FPackaging%20Data%2Easpx&Web=a09bbe21%2D1eb2%2D4a63%2D9eb5%2Dd99b56aec218")
    aa = web.find_elements(By.XPATH, "//a[contains(text(),'Click to view')]")
    print(aa)
    print(len(aa))






