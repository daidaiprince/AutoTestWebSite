"""=========================
    功能:自動化檢測網站內容
    作者:吳彥楓
    日期:2023/11/02
   ========================="""


""" 匯入模組 """
import unittest
import smtplib
import time
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from XTestRunner import HTMLTestRunner
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class TestNarl(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Chrome()
        cls.base_url = "https://www.narlabs.org.tw/"
       
    def narlabs_search(self,click_text,search_title):
        self.driver.maximize_window()  # 視窗最大化
        self.driver.get(self.base_url) # 開啟國研院官網 
     
        try:
            element = self.driver.find_element(By.ID,"privity") # 尋找隱私權同意表單[我同意]按鈕元素
            element.click() # 按下[我同意]按鈕元素
            sleep(1)
            x = self.driver.find_element(By.PARTIAL_LINK_TEXT,click_text)  # 頁面尋找[XX]文字元素
            actions = ActionChains(self.driver)
            actions.move_to_element(x).click().perform()  # 滑鼠移至[XX]文字元素並按下
            self.assertEqual(self.driver.title, search_title,"網頁標題比對不正確" )
            sleep(1)
            self.images.append(self.driver.get_screenshot_as_base64())  # 最大化截圖

        except NoSuchElementException:
            x = self.driver.find_element(By.PARTIAL_LINK_TEXT,click_text)  # 頁面尋找[XX]文字元素
            actions = ActionChains(self.driver)
            actions.move_to_element(x).click().perform()  # 滑鼠移至[XX]文字元素並按下
            self.assertEqual(self.driver.title, search_title,"網頁標題比對不正確")
            sleep(1)
            self.images.append(self.driver.get_screenshot_as_base64())  # 最大化截圖

    def test_search_key_aboutus(self):
        """點擊[關於國研院]文字，比對網頁標題內容"""
        self.narlabs_search("關於國研院","任務願景 | 國家實驗研究院")
            
    def test_search_key_ncee(self):
        """點擊[研究發展]文字，比對網頁標題內容"""
        self.narlabs_search("研究發展","地震工程 | 國家實驗研究院")

    def test_search_key_tc(self):
        """點擊[技術合作]文字，比對網頁標題內容"""
        self.narlabs_search("技術合作","技術服務手冊2023 | 國家實驗研究院")

    def test_search_key_si(self):
        """點擊[科技影響力]文字，比對網頁標題內容"""
        self.narlabs_search("科技影響力","社會參與 | 國家實驗研究院3")

    def test_search_key_mc(self):
        """點擊[媒體中心]文字，比對網頁標題內容"""
        self.narlabs_search("媒體中心","科普講堂 | 國家實驗研究院")

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()


if __name__ == '__main__':
    suit = unittest.TestSuite()
    suit.addTests([
        TestNarl("test_search_key_aboutus"),
        TestNarl("test_search_key_ncee"),
        TestNarl("test_search_key_tc"),
        TestNarl("test_search_key_si"),
        TestNarl("test_search_key_mc")
    ])

    """ 產生 HTML格式測試報告 """
    now_time = time.strftime("%Y年%m月%d日%H時%M分%S秒")
    global temp
    temp = now_time + ' - 測試報告.html'
    fp = open(temp,'wb')
    runner = HTMLTestRunner(
        stream=fp,
        tester="測試者",
        title='網站測試報告--國家實驗研究院',
        description=['類型：selenium'],
        language='zh-CN',
        rerun=1
        )
    runner.run(suit)
    fp.close()



""" 使用GMAIL信箱寄信 """
mail = MIMEMultipart()
mail['Subject'] = "[國家實驗研究院官網] 測試報告" #郵件標題
mail["from"] = "xxx@gmail.com"  #寄件者
mail["to"] = "xxx@xxx.xxx.xxx" #收件者
#信件夾檔
htmlload = MIMEApplication(open(temp,'rb').read()) 
htmlload.add_header('Content-Disposition', 'attachment', filename=temp) 
mail.attach(htmlload)
#寄送電子郵件
with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
    try:
        smtp.ehlo()  # 驗證SMTP伺服器
        smtp.starttls()  # 建立加密傳輸
        smtp.login("xxx@gmail.com", "應用程式密碼")  # 登入寄件者GMAIL
        smtp.send_message(mail)  # 寄送郵件
        print("電子郵件傳送完成!")
    except Exception as e:
        print("失敗訊息: ", e)
