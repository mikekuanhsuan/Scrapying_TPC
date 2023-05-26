import logging
import re
import time
import datetime
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException, ElementNotInteractableException
from urllib3.exceptions import MaxRetryError
from selenium.webdriver.support.ui import Select
from database import *
from typing import List, Tuple
class GirlsLD:
    def __init__(self, factory_id, account, password, months, dates):
        warnings.filterwarnings('ignore')

        self.factory_id = factory_id
        self.account = account
        self.password = password
        self.months = months
        self.dates = dates

        self.chrome_options = Options()
        self.chrome_options.add_argument("--incognito")
        self.driver = webdriver.Chrome(options=self.chrome_options)

        self.time_stamps = None
        self.interval = 2

    def execute_sql_command(self, sql):
        with db() as OCCDB:
            OCCDB.run_cmd(sql, 230)

    def run(self):
        max_retries = 10  # 最大重试次数
        retry_count = 0  # 当前重试计数

        while retry_count < max_retries:
            print(self.factory_id)
            try:
                self.login()
                if not self.goto_UID_meter_no_list():
                    continue
                for month, date in zip(self.months, self.dates):

                    sql = f"""DELETE FROM G_TPC WHERE FactoryID = '{self.factory_id}' 
                    AND Dtime >= DATETIMEFROMPARTS(YEAR(getdate()), {month}, {date}, 0, 15, 0, 0)
                    AND DTime <= DATEADD(DAY, 1, DATEFROMPARTS(YEAR(getdate()), {month}, {date}))
                    """
                    with db() as OCCDB:
                        OCCDB.run_cmd(sql, 230)

                    self.time_stamps = self.create_time_stamps(month, date)
                    
                    if not self.goto_power_analyze(month, date):
                        continue
                    
                    time.sleep(30)  # 等待數據出現
                    data = self.get_data()
                    off_peak = data['off_peak']
                    half_rush_Saturday = data['half_rush_Saturday']
                    half_rush_sp = data['half_rush_sp']
                    rush_hour = data['rush_hour']

                    # print(self.time_stamps)
                    # print(f"off_peak: {off_peak}")
                    # print(f"half_rush_Saturday: {half_rush_Saturday}")
                    # print(f"half_rush_sp: {half_rush_sp}")
                    # print(f"rush_hour: {rush_hour}")

                    values = {}

                    # 遍歷timestamp清單並更新values字典
                    for i, ts in enumerate(self.time_stamps):
                        if off_peak[i] != 'null':
                            values[ts] = {'name': '離峰時段', 'value': off_peak[i]}

                        elif half_rush_Saturday[i] != 'null':
                            if half_rush_sp =='週六半尖峰':
                                half_rush_sp = '週六半尖峰時段'

                            values[ts] = {'name': f'{half_rush_sp}', 'value': half_rush_Saturday[i]}
                        elif rush_hour[i] != 'null':
                            values[ts] = {'name': '尖峰時段', 'value': rush_hour[i]}

                    for timestamp, info in values.items():
                        name = info['name']
                        value = info['value']
                        if name == "週六半尖峰":
                            name = '週六半尖峰時段'
                        sql = f"""
                                INSERT INTO G_TPC
                                VALUES ('{self.factory_id}', LEFT('{timestamp}', 23), '{name}', {value})
                            """
                        self.execute_sql_command(sql)
                        # self.driver.quit()

                    
                # 成功完成循环，跳出while循环
               
                break

            except MaxRetryError as e:
                retry_count += 1
                logging.exception('連線異常：', e)
                print(f"重试次数：{retry_count}")
                
                # 关闭self.driver
                self.driver.quit()
                # 重新启动self.driver
                self.driver = webdriver.Chrome()  # 或者您使用的其他浏览器驱动

            finally:
                time.sleep(self.interval)

        if retry_count >= max_retries:
            print("达到最大重试次数，无法继续尝试。")

    def login(self):
        self.driver.get("https://hvcs.taipower.com.tw/Account/NewLogon")
        time.sleep(5)
        # 输入账号和密码
        self.driver.find_element(By.ID, 'UserName').send_keys(self.account)
        self.driver.find_element(By.ID, 'Password').send_keys(self.password)
        # 登录
        self.driver.find_element(By.CLASS_NAME, 'btn_Sign_in').click()
        time.sleep(5)
        try:
            wait = WebDriverWait(self.driver, 10)
            alert = wait.until(EC.alert_is_present())
            alert.dismiss()  # 取消对话框
        except TimeoutException:
            print("对话框未出现或未在预期的时间内出现。")

    def goto_UID_meter_no_list(self):
        if not self.click_element('//*[@id="sb-site"]/div[6]/div[2]/div/ul/li[1]/a'):
            return False
        
        if self.factory_id=='LD-T1HIST':
            time.sleep(10)
            if not self.click_element("//*[@id='Clickbox']/div/input[1]"):
                return False
  
        time.sleep(5)
        self.driver.get('https://hvcs.taipower.com.tw/Customer/Module/PowerAnalyze')
        time.sleep(5)
        return True
        
    def goto_power_analyze(self, month, date):
        time.sleep(10)
        if not self.click_element('//*[@id=" carrusel"]/div/div[1]/div/div[5]/div/a'):
            return False
        time.sleep(10)

        month_dropdown = '//*[@id="tab5"]/div[1]/ul/li[1]/select[2]'
        if self.element_exists(month_dropdown):
            Select(WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, month_dropdown)))).select_by_value(month)

        time.sleep(10)

        date_dropdown = '//*[@id="tab5"]/div[1]/ul/li[1]/select[3]'
        if self.element_exists(date_dropdown):
            Select(WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, date_dropdown)))).select_by_value(date)

        time.sleep(10)

        if not self.click_element('//*[@id="tab5"]/div[1]/ul/li[2]/input'):
            return False

        return True
        

    def get_data(self):
        time.sleep(10)
        page_source = self.driver.page_source
        
        # 提取数据
        data = {}


        # 離峰時段

        p1 = re.compile(r'highchart_x11 =(.*?);')
        r = p1.findall(page_source)
        p2 = re.compile(r'\d+|null')
        z = p2.findall(r[0])
        
        for i in range(z.count("0000")):
            z.remove("7")
            z.remove("0000")
        data['off_peak'] = z

 





        # 半尖峰時段和週六半尖峰時段

        p1 = re.compile(r'highchart_x12 =(.*?);')
        r = p1.findall(page_source)
        p2 = re.compile(r'\d+|null')
        x = p2.findall(r[0])
        for i in range(x.count("0000")):
            x.remove("7")
            x.remove("0000")
        data['half_rush_Saturday'] = x



        # rush名稱

        p = re.compile(r'highchart_titleName2 =(.*?);')
        r = p.findall(page_source)
        rr = r[0].replace("'","").replace(" ","")
        data['half_rush_sp']  = rr


        # 尖峰時段

        p1 = re.compile(r'highchart_x13 =(.*?);')
        r = p1.findall(page_source)
        p2 = re.compile(r'\d+|null')
        c = p2.findall(r[0])
        for i in range(c.count("0000")):
            c.remove("7")
            c.remove("0000")
        data['rush_hour'] = c



        return data


    def click_element(self, xpath):
        try:
            print(xpath)
            element = WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            return True
        except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException) as e:
            logging.exception(f'Failed to click element: {xpath}')
            return False
    

    
    
    def element_exists(self, xpath):
        try:
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, xpath)))
            return True
        except (NoSuchElementException, TimeoutException):
            return False


    def create_time_stamps(self, month: str, day: str) -> list:
        """
        為給定的日期創建一個包含15分鐘間隔的時間戳列表。

        參數:
        month (str): 需要生成時間戳的日期的月份。
        day (str): 需要生成時間戳的日期的天。

        返回:
        list: 一個時間戳列表。
        """

        # 獲取當前年份
        current_year = datetime.datetime.now().year

        # 如果月份和日期是單個數字，則使用0進行填充
        month = month.zfill(2)
        day = day.zfill(2)

        # 組合成日期字符串
        date_str = f'{current_year}-{month}-{day}'

        # 將日期字符串轉換為 datetime.datetime 對象
        date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')

        # 設定開始時間和結束時間
        start_time = datetime.datetime.combine(date_obj, datetime.time.min) + datetime.timedelta(minutes=15)
        end_time = datetime.datetime.combine(date_obj + datetime.timedelta(days=1), datetime.time.min)

        # 創建時間戳記列表
        time_stamps = []
        current_time = start_time
        while current_time <= end_time:
            time_stamps.append(current_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3])
            current_time += datetime.timedelta(minutes=15)

        return time_stamps
    

   

if __name__ == '__main__':

    def get_factory_data() -> List[Tuple]:
        """Fetches factory data from database."""
        sql = "SELECT FactoryID, FactoryName, Tpc_act, Tpc_pwd FROM Factory WHERE Tpc_act IS NOT NULL ORDER BY aOrder"
        with db() as OCCDB:
            factory_datatable = OCCDB.get_datatable(sql, 230)
        return factory_datatable


    def process_factory_data(factory_datatable: List[Tuple]) -> Tuple:
        """Processes factory data to extract relevant information."""
        factory_all = [i[0] for i in factory_datatable] # 工廠代碼
        act_all = [i[2] for i in factory_datatable]     # 帳號
        pwd_all = [i[3] for i in factory_datatable]     # 密碼
        msg_all = [i[1] for i in factory_datatable]     # 工廠
        return factory_all, act_all, pwd_all, msg_all

    def print_factory_msg(msg_all: List[str]) -> None:
        """Prints factory messages."""
        print(msg_all)

    factory_data = get_factory_data()
    factories, accounts, passwords, msg_all = process_factory_data(factory_data)
 

    accounts = ['29369948']
    passwords = ['Afa29369948']
    factories = ['HL-T1HIST']
    months = [ '5']
    dates = [ '26']


    max_failures = 5  # 指定最大失败次数
    for factory_id, account, password in zip(factories, accounts, passwords):
        failures = 0  # 追踪失败次数

        while True:
            try:
                girlsld = GirlsLD(factory_id, account, password, months, dates)
                girlsld.run()  # 执行GirlsLD对象的run()方法

                # 如果一切正常，跳出循环
                break
            except Exception as e:
                print('GirlsLD对象执行出错:', str(e))
                failures += 1

                # 如果失败次数超过最大次数，跳出循环
                if failures >= max_failures:
                    print('达到最大失败次数，程序关闭。')
                    break