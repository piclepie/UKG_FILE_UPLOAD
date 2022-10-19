import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
import pandas as pd

# you need to download a chromedriver https://chromedriver.chromium.org/downloads which matches your chrome browser version
chromedriverpath = 'C:/Users/YOURUSERNAME/Desktop/chromedriver_win32/chromedriver.exe'
defalutdownloadpath = 'path for your downloadfiles'


def selenium_driver():
    svs = Service(
        chromedriverpath)
    options_dvr = webdriver.ChromeOptions()
    # headless can make chrome driver to run at backend.
    # options_dvr.add_argument('headless')

    options_dvr.add_experimental_option("prefs", {
        "download.default_directory": defalutdownloadpath,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(service=svs, options=options_dvr)

    return driver

# #fileuploader
def UKG_fileuploader(Employeeidlist,filepath):
    # initiating selenium driver
    driver = selenium_driver()
    yourukghostname = input("Please enter your UKG hostname: ")
    ukgusername = input("Please enter your UKG UserName: ")
    ukgpassword = input("Please enter your UKG password: ")
    # ukgrole = input("Please enter your role: ")
    # initiating ukg login
    url = f'https://{yourukghostname}.ultipro.com/login.aspx'
    driver.get(url)
    driver.find_element(
        By.ID, 'ctl00_Content_Login1_UserName').send_keys(ukgusername)
    driver.find_element(
        By.ID, 'ctl00_Content_Login1_Password').send_keys(ukgpassword)
    driver.find_element(By.ID, 'ctl00_Content_Login1_LoginButton').click()

    def is_element_exist(driver,id):
        try:
            driver.find_element(By.ID, id)
            return True
        except:
            return False

    #select email auth. IF need auth. If element exist. Then Email Auth. Else Login
    if is_element_exist(driver,'ctl00_Content_radioButtonEmail'):
        driver.find_element(By.ID,'ctl00_Content_radioButtonEmail').click()
        #submit
        driver.find_element(By.ID,'ctl00_Content_ButtonMultiFactDeliveryMethod').click()
        #get the access code from spreadsheet. wait 10 seconds
        code = input("Enter verification code: ")
        print(code)
        #send code
        driver.find_element(By.ID,'ctl00_Content_txtMultiFactorAccessCode').send_keys(code)
        driver.find_element(By.ID,"ctl00_Content_chkRememberDevice").click()
        #login 
        driver.find_element(By.ID,'ctl00_Content_ButtonMultiFactorAccessCodeEntry').click()

        print("Login Sucessfully")
        
   
    for id in Employeeidlist:
                    
        driver.get(f'https://{yourukghostname}.ultipro.com/pages/VIEW/eeMain.aspx?pendopdk=eeadm&pendoWfiHow=new-menu&pendoSID=ybduja0vzuksy4enj23nb2wn&pendoTID=a5759026-2856-4393-95a1-2050bf372e4d&destinationId=424&USParams=PK%3dEEADM!MenuID%3d424!PageRerId%3d424!ParentRerId%3d346!subDivRerID%3d674')
        print("get element")
        role = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_Content_roleSelector1"]')))
        Select(role).select_by_index(0)
        print(" Default Role selected")
        Select(driver.find_element(By.XPATH,'//*[@id="GridView1_firstSelect_0"]')).select_by_value('Employee number') 
        time.sleep(2)
        input = None
        try:
            input = WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((By.XPATH,'//*[@id="GridView1_TextEntryFilterControlInputBox_0"]')))
        except:
            input = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.XPATH,'//*[@id="GridView1_TextEntryFilterControlInputBox_3"]')))
        input.clear()
        input.send_keys(id)
        driver.find_element(By.XPATH,'//*[@id="GridView1_filterButton"]').click()                          

        profile = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Content_GridView1"]/tbody/tr/td[1]/a')))
        profile.click()
        #pop up new window
        window_before = driver.window_handles[0]
        print(driver.window_handles[0])
        window_popout = driver.window_handles[1]
        driver.switch_to.window(window_popout)
        doc_tab = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH,'//*[@id="1305"]')))
        doc_tab.click()
        driver.switch_to.frame("ContentFrame")
        add = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.ID,'ctl00_btnAdd')))
        add.click()                                 
        #ctl00_Content_frmDocument_fileUpload
        fileselect = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.ID,'ctl00_Content_frmDocument_fileUpload')))              
        #send file path     
        filepath = filepath + id +".pdf"                           
        fileselect.send_keys(filepath)
        # change file description here
        Document_txbDescription = "YOUR file description"
        driver.find_element(By.ID,'ctl00_Content_frmDocument_txbDescription').send_keys(Document_txbDescription)
        category = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_Content_frmDocument_csCategory"]')))
        Select(category).select_by_index(0) 
        # select first document category
        #document viewable by employee? check is yes.
        driver.find_element(By.ID,"ctl00_Content_frmDocument_chkViewableByEe").click()
        #fill file name chose category, save.            
        driver.find_element(By.ID,'ctl00_btnSave').click()                
        time.sleep(2)
        driver.close()
        driver.switch_to.window(window_before)
        
        print("done")


def readfile():
    # default fitst sheet
    yourfilepathcontainseeid = input("Enter file path: ")
    df = pd.read_excel(yourfilepathcontainseeid,0)
    return df['EmployeeID'].values



def process():
    
    Employeeidlist = readfile()
    # file path should contains files name with employeeID
    filepath = f'C:/Users/YOURUSERNAME/Desktop/'
    UKG_fileuploader(Employeeidlist,filepath)



