import sys
path="F:\\CJMDXTtest\\"
sys.path.append(path)

from . import *
from testScripts import CreateContacts
from WriteTestResult import writeTestResult
import traceback
from testScripts.Excel_Obj import *
from util.Log import *


def cjmTest():
    try:
        #根据Excel文件中的sheet名获取sheet对象
        caseSheet=excelObj.getSheetByName('测试用例')

        #获取测试用例sheet中是否执行该列对象
        isExecuteColumn=excelObj.getColumn(caseSheet,testCase_isExecute)

        #记录执行成功的测试用例个数
        successfulCase=0

        #记录需要执行的用例个数
        requiredCase=0

        for idx,i in enumerate(isExecuteColumn[1:]):

            caseName=excelObj.getCellOfValue(caseSheet,rowNo=idx+2,colsNo=testCase_testCaseName)
            logging.info('\033[1;32;m-----开始执行%s用例-------\033[0m' % caseName)

            #循环遍历”测试用例“表中的测试用例，执行被设置为执行的用例
            if i.value.lower()=='y':
                requiredCase+=1

                #获取测试用例表中，第idx+1行中用测试执行时所使用的框架类型
                userFrameWorkName=excelObj.getCellOfValue(caseSheet,rowNo=idx+2,colsNo=testCase_frameWorkName)

                #获取测试用例表中，第idx+1行中执行用例的步骤sheet名
                stepSheetName=excelObj.getCellOfValue(caseSheet,rowNo=idx+2,colsNo=testCase_testStepSheetName)

                if userFrameWorkName=='数据':
                    logging.info("\033[1;32;m*******调用数据驱动*****\033[0m")


                    dataSheetName=excelObj.getCellOfValue(caseSheet,rowNo=idx+2,colsNo=testCase_dataSourceSheetName)

                    #获取第idx+1行测试用例的步骤sheet对象
                    stepSheetObj=excelObj.getSheetByName(stepSheetName)

                    #获取第idx+1行测试用例使用的数据sheet对象
                    dataSheetObj=excelObj.getSheetByName(dataSheetName)


                    #通过数据驱动框架执行添加联系人
                    result=CreateContacts.dataDrivenFun(dataSheetObj,stepSheetObj)

                    if result:
                        logging.info("\033[1;32;m用例%s执行成功\033[0m"%caseName)
                        successfulCase+=1
                        writeTestResult(caseSheet,rowNo=idx+2,colsNo="testCase",testResult='pass')

                    else:
                        logging.info("\033[0;43;41m用例%s执行失败\033[0m" % caseName)
                        successfulCase += 1
                        writeTestResult(caseSheet, rowNo=idx + 2, colsNo="testCase", testResult='faild')

                elif userFrameWorkName=='关键字':
                    logging.info("\033[1;32;m****调用关键字驱动***\033[0m")

                    caseStepObj=excelObj.getSheetByName(stepSheetName)
                    stepNums=excelObj.getRowsNumber(caseStepObj)
                    successfulSteps=0
                    logging.info("\033[1;32;m测试用例共%s步\033[0m"%stepNums)

                    for index in range(2,stepNums+1):
                        #获取步骤sheet中第index行对象
                        stepRow=excelObj.getRow(caseStepObj,index)

                        #获取关键字作为调用的函数名
                        keyWord=stepRow[testStep_keyWords-1].value


                        #获取操作元素定位方式作为调用的函数的参数
                        locationType=stepRow[testStep_locationType-1].value

                        #获取操作元素的定位表达式作为调用函数的参数
                        locatorExpression=stepRow[testStep_locatorExpression-1].value

                        #获取操作值作为调用函数的参数
                        operateValue=stepRow[testStep_operateValue-1].value


                        if keyWord=='input_data':
                            Today=today()
                            operateValue=str(operateValue)
                            operateValue=Today+operateValue+global_variable


                        if isinstance(operateValue,int):

                            operateValue=str(operateValue)


                        tmpStr = "'%s','%s'" % (locationType.lower(), locatorExpression.replace("'",'"')) if locationType and locatorExpression else ""


                        if tmpStr:
                            tmpStr += ",'" + operateValue + "'" if operateValue else ""

                        else:
                            tmpStr += "'" + operateValue + "'" if operateValue else ""


                        runStr = keyWord + "(" + tmpStr + ")"



                        try:

                            eval(runStr)



                        except Exception as e:
                            logging.info('\033[4;31;m执行步骤%s发生异常\033[0m'%stepRow[testStep_testStepDescribe-1].value)

                            #截取异常图片
                            capturePic=capture_screen()

                            #获取详细的异常堆栈信息
                            errorInfo=traceback.format_exc()
                            writeTestResult(caseStepObj,rowNo=index,colsNo="caseStep",testResult="faild",errorInfo=str(errorInfo),picPath=capturePic)
                            # break
                        else:
                            successfulSteps+=1
                            logging.info('\033[1;32;m步骤%s成功\033[0m'%stepRow[testStep_testStepDescribe-1].value)
                            writeTestResult(caseStepObj,rowNo=index,colsNo="caseStep",testResult="pass")

                    if successfulSteps==stepNums-1:
                        successfulCase+=1
                        logging.info('\033[1;32;m用例%s执行通过\033[0m'%caseName)
                        writeTestResult(caseSheet,rowNo=idx+2,colsNo='testCase',testResult='pass')

                    else:
                        logging.info('\033[4;31;m用例%s执行失败\033[0m' % caseName)
                        writeTestResult(caseSheet, rowNo=idx + 2, colsNo='testCase', testResult='faild')

            else:
                '''
                清空不需要执行的执行时间和执行结果，
                异常信息，异常图片单元格
                '''
                logging.info('\033[0;34;m用例%s设置的为不执行，请检查确认。。\033[0m'%caseName)
                writeTestResult(caseSheet,rowNo=idx+2,colsNo='testCase',testResult="")

                logging.info("\033[1;32;m共%d条用例，%d条需要被执行，成功执行%d条\033[0m"%(len(isExecuteColumn)-1,requiredCase,successfulCase))

    except Exception as e:
        logging.debug(traceback.print_exc())



# cjmTest()







