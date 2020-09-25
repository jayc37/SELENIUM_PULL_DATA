#! coding: utf-8
import os
import sys
import codecs
import pyodbc
import logging
import argparse
from time import sleep
from random import randint
from selenium import webdriver
from datetime import datetime,timedelta
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC


#---------------------------------------------------------------------------------------------------
#Thêm vào lớp vetify bắt try cath chặt chẽ#
def random_sleep(min_s, max_s):
    sleep(randint(min_s, max_s))



class pulldata:
    def __init__(self,servername,dbname,uid,pwd):
        self.username = ''
        self.password = ''
        self.data_day = None
        self.driver = webdriver.Chrome(ChromeDriverManager().install(),options=op)
        self.conn = pyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (servername, dbname, uid, pwd)) 
        self.curr = self.conn.cursor()

        
    def login(self):
        pul_URL = ""

        self.driver.get(pul_URL)
        self.driver.set_page_load_timeout(120)
        self.driver.maximize_window()
        
        try:
            random_sleep(3,5)
            username_ele = self.driver.find_element_by_name("SWEUserName")
            username_ele.send_keys(Keys.F5)
            random_sleep(2,3)
            username_ele.send_keys(self.username)
            random_sleep(1,2)

            password_ele = self.driver.find_element_by_name("SWEPassword")
            password_ele.send_keys(self.password)
            random_sleep(1,2)
            

            try:
                login_ele = self.driver.find_element_by_xpath(".//*[@id='s_swepi_22']")
                login_ele.click()
            except Exception as e:
                print(str(e))
                self.verify_func(login_ele)
                login_ele = self.driver.find_element_by_xpath(".//*[@id='s_swepi_22']")
                actions = ActionChains(self.driver)
                actions.move_to_element(login_ele).click().perform()
                pass
            random_sleep(5, 8)

            try:
                click_thongtinxe = self.driver.find_element(By.XPATH,'//a[text()="Thông tin xe"]')
                click_thongtinxe.click()
            except Exception as e:
                print(str(e))
                self.verify_func(click_thongtinxe)
                click_thongtinxe = self.driver.find_element(By.XPATH,'//a[text()="Thông tin xe"]')
                actions = ActionChains(self.driver)
                actions.move_to_element(click_thongtinxe).click().perform()
                pass
            random_sleep(3, 5)
            try:
                click_dsxe = self.driver.find_element(By.XPATH,'//a[text()="Danh sách xe"]')
                click_dsxe.click()
            except Exception as e:
                print(str(e))
                self.verify_func(click_dsxe)
                click_dsxe = self.driver.find_element(By.XPATH,'//a[text()="Danh sách xe"]')
                actions = ActionChains(self.driver)
                actions.move_to_element(click_dsxe).click().perform()
                pass

            random_sleep(3, 5)
            
            try:
                click_search = self.driver.find_element_by_xpath("//*[@id='s_2_1_11_0_Ctrl']")
                click_search.click()
            except Exception as e:
                print(str(e))
                pass

            # find_head_date = self.driver.find_element_by_id("1_s_2_l_Evaluation_Date")
            # find_head_date.click()
            try:
                find_head_date = self.driver.find_element_by_id("1_s_2_l_Evaluation_Date")
                find_head_date.click()
            except Exception as e:
                print(str(e))
                self.verify_func(find_head_date)
                find_head_date.click()
                pass
            random_sleep(1, 2)

            fill_head_date = self.driver.find_element_by_name("Evaluation_Date")
            try:

                find_head_date.click()
            except Exception as e:
                print(str(e))
                self.verify_func(find_head_date)
                find_head_date.click()
                pass
            random_sleep(1, 2)
            fill_head_date.send_keys(self.data_day)
            fill_head_date.send_keys(Keys.ENTER)
            random_sleep(5, 8)
            self.get_data2()
            
        except TimeoutException as e:
            print(str(e))
            pass

    def verify(self,element):
        try:
            element
            return True
        except Exception as e:
            print(str(e))
            return False

    def get_data(self,file):
        try:
            soup = BeautifulSoup (open(file, encoding="UTF-16LE"), features="lxml")
            tables = soup.findAll("table", {"":""})

            for tn, table in enumerate(tables):
                    dab = []
                    for i in range(1,4):
                        rows = table.findAll("tr")
                        row_lengths = [len(r.findAll(['th', 'td'])) for r in rows]
                        ncols = max(row_lengths)
                        nrows = len(rows)
                        data = []

                        for i in range(nrows):
                            rowD = []
                            for j in range(ncols):
                                rowD.append('')
                            data.append(rowD)

                        for i in range(len(rows)):
                            row = rows[i]
                            rowD = []
                            cells = row.findAll(["td", "th"])

                            for j in range(len(cells)):
                                cell = cells[j]
                                cspan = int(cell.get('colspan', 1))
                                rspan = int(cell.get('rowspan', 1))
                                l  = 0
                                
                                for k in range(rspan):

                                    while data[i + k][j + l]:
                                        l += 1
                                
                                    for m in range(cspan):
                                        cell_n = j + l + m
                                        row_n = i + k
                                        cell_n = min(cell_n, len(data[row_n])-1)
                                        data[row_n][cell_n] += cell.text

                            data.append(rowD)
            self.format_data(data)
        except Exception as e:
            print(str(e))
            pass

    def get_data2(self):
        try:
            file_name = 'output.HTML'
            path = self.get_download_path()
            file = '{}\\{}'.format(path,file_name)
            if os.path.exists(file):
                os.remove(file)
            else:
                print("------------wait a minutue...")
            try:
                wait = WebDriverWait(self.driver, 20)
                random_sleep(4,6)
                button_menu = wait.until(EC.element_to_be_clickable((By.ID,'s_at_m_2'))).click()
                # button_menu = self.driver.find_element_by_id('s_at_m_2')
                # actions = ActionChains(self.driver)
                # actions.move_to_element(button_menu).click().perform()
            except Exception as e:
                try:
                    self.verify_func(button_menu)
                finally:
                    button_menu = self.driver.find_element_by_id('s_at_m_2')
                    actions = ActionChains(self.driver)
                    actions.move_to_element(button_menu).click().perform()
                pass

            random_sleep(3,5)
            try:
                find_export = self.driver.find_element(By.XPATH,'//a[text()="Export..."]')
                find_export.click()
            except Exception as e:
                print(str(e))
                self.verify_func(find_export)
                find_export = self.driver.find_element(By.XPATH,'//a[text()="Export..."]')
                actions = ActionChains(self.driver)
                actions.move_to_element(find_export).click().perform()
                pass
            random_sleep(3,5)
            dropdown = ActionChains(self.driver)
            dropdown.send_keys(Keys.TAB * 4)
            dropdown.send_keys(Keys.ARROW_DOWN * 2)
            dropdown.perform()

            dropdown.send_keys(Keys.ENTER)
            dropdown.perform()
            try:              
                bnexts = self.driver.find_element(By.XPATH,'//a[text()="Next"]')
                bnexts.click()
            except Exception as e:
                print(str(e))
                self.verify_func(bnexts)
                bnexts = self.driver.find_element(By.XPATH,'//a[text()="Next"]')
                actions = ActionChains(self.driver)
                actions.move_to_element(bnexts).click().perform()
                pass
            random_sleep(20,25)
            print("Get ready pull to database !!!")
            self.get_data(file)
            print("Pull Success to Database!!")

            if os.path.exists(file):
                os.remove(file)
            else:
                print("clean download folder")
        except Exception as e:
            print(str(e))
            pass
        finally:
            self.log_out()

    def verify_func(self,element):
        try:
            actions = ActionChains(self.driver)
            print("---------------Reconnect 1-----------------")
            actions.send_keys(Keys.F5).perform()
            random_sleep(3,4)
            element
            random_sleep(1,2)
            print("reconnected successfully")
        except Exception as e:
            print("---------------Connection failed-----------------")
            actions.send_keys(Keys.F5).perform()
            self.verify_func2(element,actions)
            pass

    def verify_func2(self,element,actions):
        try:
            print("---------------Reconnect 2-----------------")
            random_sleep(3,4)
            element
            print("reconnected successfully")
        except Exception as e:
            print("---------------Connection failed-----------------")
            actions.send_keys(Keys.F5).perform()
            random_sleep(6,7)
            self.verify_func3(element,actions)

    def verify_func3(self,element,actions):
        try:
            print("---------------Reconnect 3-----------------")
            element
            print("reconnected successfully")
        except:
            print("---------------Connection failed-----------------")
            random_sleep(1,2)
            print("---------------QUIT-----------------")
            self.driver.quit()
    def log_out(self):
        print("done with" + str(self.username))
        random_sleep(1,2)
        print("Continues with next account" )
            

    def format_data(self,data):
        for i in range(1,len(data)):
            if data[0] == '':
                data = data[1:]
            v = data[i]
            arg_v = self.format_string(v)
            try:
                query = """
                INSERT INTO TABLE VALUES {}""".format(tuple([str(arg_v[x]) for x in range(len(arg_v))]))
                self.curr.execute(str(query))
                self.conn.commit()
                print('insert success at location: ' + str(i))
                self.return_date()
            except Exception as e:
                print(str(e))
                pass
            

    def format_string(self,arg):
        x = []
        if arg[0] == '':
            for i in range(1,len(arg)):
                strg = arg[i].replace('\xa0','')
                x.append(strg)
            x.append(str(self.username))
        else:
            for i in range(len(arg)):
                strg = arg[i].replace('\xa0','')
                x.append(str(strg))
            x.append(str(self.username))
        return x

    def getacount(self):
        self.insert_user()
        self.curr.execute("SELECT username,passwords,date_update from USER_HMS")
        result = self.curr.fetchall()
        actions = ActionChains(self.driver)
        try:
            for i in range(len(result)):

                self.username = result[i][0]
                self.password = result[i][1]
                date = str(result[i][2])
                date = datetime.strptime(date, "%Y-%m-%d").strftime("%d/%m/%Y")
                self.data_day = str(date)
                print(str(i) + ":: Data date: " + self.data_day)
                print(str(i) + ":: User Name: " +self.username)

                self.login()
        except Exception as e:
            print('account' + str(e))
            print(str(e))
            pass

    def return_date(self):
        try:

            """Form"""
            data_da = str(self.data_day)
            data_da = datetime.strptime(data_da, "%d/%m/%Y").strftime("%Y-%m-%d")
            date = datetime.strptime(str(data_da), "%Y-%m-%d")
            modified_date = date + timedelta(days=1)
            data_date = datetime.strftime(modified_date, "%Y/%m/%d")

            self.curr.execute("UPDATE USER_HMS SET DATE_UPDATE = '{}' WHERE USERNAME = '{}'".format(str(data_date),str(self.username)))
            self.conn.commit()

        except Exception as e:
            print('<return_date> + {}'.format(str(e)))


    def insert_user(self):
        """ Insert User"""
        self.curr.execute("SELECT USERNAME FROM USER_HMS")
        result = self.curr.fetchall()
        if len(result) == 0:  
            print(len(result))  
            with open('account.txt', 'r') as a:
                account =  a.readlines()
            a.close()
            for i in range(len(account)):
                acni = account[i]
                acni = acni.replace('\n','')
                acni = acni.replace(' ',',')
                ac = acni.split(',')
                print(ac)
                for a in range(len(ac)):
                    print(str(ac[0]),str(ac[1]),str(ac[2]),str(ac[3]))
                    self.curr.execute("INSERT INTO USER_HMS VALUES (%s,%s,%s,%s)" %(str(ac[0]),str(ac[1]),str(ac[2]),str(ac[3])))
                    self.conn.commit()
                    print("insert account success")
    
    def get_download_path(self):
        """Returns the default downloads path for linux or windows"""
        if os.name == 'nt':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
            return location
        else:
            return os.path.join(os.path.expanduser('~'), 'downloads')         

            
if __name__ == '__main__':
    with open('SQL_server_connection.txt','r') as f:
        x = f.readlines()
        servername = str(x[0]).replace('\n','')
        print(servername)
        dbname = str(x[1]).replace('\n','')
        uid = str(x[2]).replace('\n','')
        pwd = str(x[3]).replace('\n','')
    f.close()
    

    op = webdriver.ChromeOptions()
    op.add_argument("--disable-gpu")
    op.add_argument('start-maximized')
    op.add_argument("--disable-extensions")
    op.add_argument("enable-automation")
    op.add_argument("--no-sandbox")
    op.add_argument("--disable-infobars")
    op.add_argument("--disable-dev-shm-usage")
    op.add_argument("--disable-browser-side-navigation")


    pul = pulldata(servername,dbname,uid,pwd)
    try:
        pul.getacount()
        print("Finish")
    except:
        pul.driver.close()
        print("Quit process")    
    pul.driver.close()