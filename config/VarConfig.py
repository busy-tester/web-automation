import sys
path="F:\\CJMDXTtest\\"
sys.path.append(path)
import os


#获取当前文件所在目录的父目录的绝对路径
parentDirPath=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#异常截图存放目录绝对路径
screenPicturesDir=parentDirPath+"\\exceptionpictures\\"

#测试数据文件存放绝对路径
dataFilePath=parentDirPath+"\\testData\\CJMCASE\\超级码测试用例.xlsx"

#cookie路径
cookiePath=parentDirPath+"\\config\\add_cookie.txt"

#浏览器驱动
chromeDriverFilePath=parentDirPath+"\\config\\driver\\chromedriver.exe"
firefoxDriverFilePath=parentDirPath+"\\config\\driver\\geckodriver.exe"
ieDriverFilePath=''

#全局变量
global_variable='1'




#测试数据文件中，测试用例表中部分列对应的数字序号
testCase_testCaseName=2
testCase_frameWorkName=4
testCase_testStepSheetName=5
testCase_dataSourceSheetName=6
testCase_isExecute=7
testCase_runTime=8
testCase_testResult=9


#用例步骤表中，部分列对应的数字序号
testStep_testStepDescribe=2
testStep_keyWords=3
testStep_locationType=4
testStep_locatorExpression=5
testStep_operateValue=6
testStep_runTime=7
testStep_testResult=8
testStep_errorInfo=9
testStep_errorPic=10


#数据表源中，是否执行列对应的数字编号
dataSource_isExecute=7
dataSource_email=3
dataSource_runTime=8
dataSource_result=9

