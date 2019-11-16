import sys
path="F:\\CJMDXTtest\\"
sys.path.append(path)

from selenium import webdriver
from config.VarConfig import ieDriverFilePath
from config.VarConfig import chromeDriverFilePath
from config.VarConfig import firefoxDriverFilePath
from config.VarConfig import cookiePath
from util.ObjectMap import getElement,getElements
from selenium.webdriver.support.ui  import Select
from util.ClipboardUtil import Clipboard
from util.KeyBoardUtil import KeyboardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
from config.VarConfig import *
from util.Log import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import json

# driver=webdriver.Chrome()
# driver.find_element_by_id('kw').send_keys(Keys.CONTROL,'x')

#定义全局driver变量
driver=None

#全局的等待实例变量
waitUtil=None


# 打开浏览器
def open_browser(browserName,*args):

    global driver,waitUtil
    try:
        if browserName.lower()=='ie':
            driver=webdriver.Ie()
        elif browserName.lower()=="firefox":
            driver = webdriver.Firefox(executable_path=firefoxDriverFilePath)
        else:
            # # 创建Chrome浏览器的一个Options实例对象
            # chrome_options = Options()
            # # 添加屏蔽--ignore-certificate-errors提示信息的设置参数项
            # chrome_options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
            # driver = webdriver.Chrome(executable_path=chromeDriverFilePath, chrome_options=chrome_options)
            # driver=webdriver.Chrome(executable_path=chromeDriverFilePath)


            # 个人资料路径
            user_data_dir = r'--user-data-dir=C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
            # 加载配置数据
            option = webdriver.ChromeOptions()
            option.add_argument(user_data_dir)
            # 启动浏览器配置
            driver = webdriver.Chrome(chrome_options=option,executable_path=chromeDriverFilePath)
            # driver = webdriver.Chrome(executable_path=chromeDriverFilePath)

        #driver对象创建成功后，创建等待类实例对象
        waitUtil=WaitUtil(driver)
    except Exception as e:
        raise


#访问某个网址
def visit_url(url,*args):
    global driver
    try:
        driver.get(url)
    except Exception as e:
        raise e


#关闭浏览器
def close_browser(*args):
    global driver
    try:
        driver.quit()
    except Exception as e:
        raise e


#强制等待
def sleep(sleepSeconds,*args):
    try:
        time.sleep(int(sleepSeconds))
    except Exception as e:
        raise e


#清除输入框默认内容
def clear(locationType,locatorExpression,*args):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).clear()
    except Exception as e:
        raise e


#在页面输入框中输入数据
def input_string(locationType,locatorExpression,inputContent):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).send_keys(inputContent)
    except Exception as e:
        raise e



#获取一组数据
def input_string_s(locationType,locatorExpression,num):
    global driver,waitUtil
    try:
        Elements=getElements(driver,locationType,locatorExpression)
        # print(Elements)
        for i in Elements:
            input_string(i[0], i[1], num)

    except Exception as e:
        raise e




#在页面输入框中输入数据
def input_data(locationType,locatorExpression,inputContent):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).send_keys(inputContent)
    except Exception as e:
        raise e




#单击页面元素
def click(locationType,locatorExpression,*args):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).click()
    except Exception as e:
        raise e




#断言页面源码是否存在某关键字或关键字符串
def assert_string_in_pagesource(assertString,*args):
    global driver,waitUtil
    try:
        assert assertString in driver.page_source,"%s not found in page source。。。"%assertString
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e
    else:
        logger.info("\033[1;32;m关键字%s存在\033[0m"%assertString)



#断言页面源码是否存在某关键字或关键字符串
def assert_pagesource(assertString,*args):
    global driver,waitUtil
    try:
        TODAY =today()
        TODAY=str(TODAY)
        assertString =str(assertString)
        pageString = TODAY + assertString + global_variable
        assert pageString in driver.page_source,"%s not found in page source。。。"%assertString
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e
    else:
        logger.info("\033[1;32;m关键字%s存在\033[0m"%pageString)



#断言页面标题是否存在给定的关键字符串
def assert_title(titleStr,*args):
    global driver
    try:
        assert titleStr in driver.title,"%s not found in title。。。"%titleStr
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e




#获取页面标题
def getTitle(*args):
    global driver
    try:
        return driver.title
    except Exception as e:
        raise e



#获取页面源码
def getPageSource(*args):
    global driver
    try:
        return driver.page_source
    except Exception as e:
        raise e



#切换进入frame
def switch_to_frame(locationType,framelocatorExpression,*args):
    global driver
    try:
        driver.switch_to.frame(getElement(driver,locationType,framelocatorExpression))
    except Exception as e:
        print('frame error。。。')
        raise e


#切换进入frame,0
def switchto_frame(tagName,*args):
    global driver
    try:
        driver.switch_to.frame(driver.find_element_by_tag_name(tagName))
    except Exception as e:
        print('frame error。。。')
        raise e



#切出frame
def switch_to_default_content(*args):
    global driver
    try:
        driver.switch_to.default_content()
    except Exception as e:
        raise e


#点击多选框
def clickCheckBox(locationType,framelocatorExpression,*args):
    global driver,waitUtil
    try:
        ListElement=getElements(driver,locationType,framelocatorExpression)
        print(ListElement)
        for element in ListElement:
            element.click()

    except Exception as e:
        raise e


#下拉列表选择
def selecter_list(locationType,framelocatorExpression,textValue,*args):
    global driver,waitUtil
    try:
        select_element=Select(getElement(driver,locationType,framelocatorExpression))
        select_element.select_by_visible_text(textValue)
        value= select_element.all_selected_options[0].text
        assert value==textValue,"断言失败,当前选择的不是%s"%textValue
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:

        raise e
    else:
        logger.info("\033[1;32;m关键字%s选择成功\033[0m" % textValue)




#模拟Ctrl+V操作
def paste_string(pasteString,*args):
    try:
        imagePath=parentDirPath+"\\testData\\image\\"+pasteString
        Clipboard.setText(imagePath)
        time.sleep(2)
        KeyboardKeys.twoKeys('ctrl','v')
    except Exception as e:
        raise e


#模拟Ctrl+x操作
def shear(locationType,framelocatorExpression,*args):
    global driver,waitUtil
    try:
        total=getElement(driver,locationType,framelocatorExpression)
        total.send_keys(Keys.CONTROL,'x')
    except Exception as e:
        raise e


#获取元素属性的值
def elementAttribute(locationType,framelocatorExpression,Attribute,*args):
    global driver,waitUtil
    try:
        element=getElement(driver,locationType,framelocatorExpression)
        AttributeValue=element.get_attribute(Attribute)
        return AttributeValue

    except Exception as e:
        raise e


#删除页面属性
def del_attribute(locationType,framelocatorExpression,attributeName):
    global driver,waitUtil
    try:
        elementobj=getElement(driver,locationType,framelocatorExpression)
        driver.execute_script("arguments[0].removeAttribute(arguments[1])",elementobj, attributeName)
    except Exception as e:
        raise e



#输入日期
def input_date(locationType,locatorExpression,inputContent,*args):
    global driver,waitUtil
    try:
        input_string(locationType,locatorExpression,inputContent)
    except Exception as e:
        raise e





#模拟tab键
def press_tab_key(*args):
    try:
        KeyboardKeys.oneKey('tab')
    except Exception as e:
        raise e


#模拟向下按键
def press_down_key(locationType,locatorExpression,*args):
    global driver,waitUtil
    try:
        getElement(driver,locationType,locatorExpression).send_keys(Keys.DOWN)
    except Exception as e:
        raise e





#模拟Enter键
def press_enter_key(*args):
    try:
        KeyboardKeys.oneKey('enter')
    except Exception as e:
        raise e



#窗口最大化
def maximize_browser():
    global driver
    try:
        driver.maximize_window()
    except Exception as e:
        raise e


#添加cookies
def add_cookies():
    global driver
    try:
        cookie=driver.get_cookie('access_token')
        f=open(cookiePath,'w',encoding='utf-8')
        f.write(json.dumps(cookie))
        f.close()
    except Exception as e:
        raise e


#获取cookie
def get_cookiesww():
    global driver
    try:
        file= open(cookiePath,'r',encoding='utf-8')
        data = json.loads(file.read())
        driver.add_cookie(data)
        print('pass')
        file.close()
    except Exception as e:
        raise e


#刷新页面
def refresh():
    global driver
    try:
        driver.refresh()
    except Exception as e:
        raise



#截取屏幕图片
def capture_screen(*args):
    global driver
    currTime=getCurrentTime()
    picNameAndPath=str(createCurrentDataDir())+"\\"+str(currTime)+".png"
    print(picNameAndPath)
    try:
        driver.get_screenshot_as_file(picNameAndPath.replace('\\',r'\\'))
        # driver.get_screenshot_as_file(screenPicturesDir+getCurrentDate()+getCurrentTime()+'.png')

    except Exception as e:
        raise e
    else:
        return picNameAndPath



'''
显示等待页面元素出现在DOM中，但并不一定可见，
存在则返回该页面元素对象
'''
def waitPresenceOfElementLocated(locationType,locatorExpression,*args):
    global waitUtil
    try:
        waitUtil.presenceOfElementLocated(locationType,locatorExpression)
    except Exception as e:
        raise e




# 检查frame是否存在,存在则切换进frame控件中
def waitFrameToBeAvailableAndSwitchToIt(locationType,locatorExpression,*args):
    global waitUtil,driver
    try:
        waitUtil.frameToBeAvailableAndSwitchToIt(locationType,locatorExpression)
    except Exception as e:
        raise e




#显示等待页面元素出现在DOM中，并且可见
def waitVisibilityOfElementLocated(locationType,locatorExpression,*args):
    global waitUtil
    try:
        waitUtil.visibilityOfElementLocated(locationType,locatorExpression)
    except Exception as e:
        raise e


#判断当前按钮是否选中
def button_is_selected(locationType,locatorExpression,*args):
    global driver,waitUtil
    try:
        if not getElement(driver, locationType, locatorExpression).is_selected():
            click(locationType, locatorExpression)
    except Exception as e:
        raise e


#将页面的滚动条滑动到页面的最下方
def scroll_to_buttom():
    global driver,waitUtil
    try:

        driver.execute_script('window.scrollTo(1000,document.body.scrollHeight);')
        # driver.execute_script('window.scrollTo(1000,document.body.scrollHeight);')
    except Exception as e:
        raise e



#将元素滚动到屏幕中间
def scrollIntoView(*args):
    global driver,waitUtil
    try:
        js = "var q=document.documentElement.scrollTop=10000"
        driver.execute_script(js)
    except Exception as e:
        raise e



#将元素滚动到指定元素
def specified(locationType,locatorExpression,*args):
    global driver,waitUtil
    try:
        target = getElement(driver,locationType,locatorExpression)
        driver.execute_script("arguments[0].scrollIntoView();", target)
    except Exception as e:
        raise e



