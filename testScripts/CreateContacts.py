import sys
path="F:\\CJMDXTtest\\"
sys.path.append(path)

# from . import *
from WriteTestResult import writeTestResult
from testScripts.Excel_Obj import *
from util.Log import *


def dataDrivenFun(dataSourceSheetObj,stepSheetObj):
    try:
        #获取数据表中是否执行该列对象  联系人表中是否执行
        dataIsExecuteColumn=excelObj.getColumn(dataSourceSheetObj,dataSource_isExecute)


        emailColumn=excelObj.getColumn(dataSourceSheetObj,dataSource_email)


        #获取测试步骤表中存在数据区域的行数
        stepRowNums=excelObj.getRowsNumber(stepSheetObj)


        #记录成功执行的数据总数
        successDatas=0

        #记录被设置为执行的数据总数
        requiredDatas=0


        '''
        遍历数据表，准备进行数据驱动测试
        因为第一行是标题行，所以从第二行开始遍历
        '''
        for idx,data in enumerate(dataIsExecuteColumn[1:]):
            if data.value=='y':
                logging.info('开始添加联系人%s'%emailColumn[idx+1].value)
                requiredDatas+=1

                #记录执行成功步骤数变量
                successStep=0

                for index in range(2,stepRowNums+1):
                    #获取数据驱动测试步骤表中，第index行对象
                    rowObj=excelObj.getRow(stepSheetObj,index)

                    #获取关键字作为调用的函数名
                    keyWord=rowObj[testStep_keyWords-1].value

                    #获取操作元素定位方式作为调用的函数的参数
                    locationType=rowObj[testStep_locationType-1].value

                    #获取操作元素的定位表达式作为调用函数的参数
                    locatorExpression=rowObj[testStep_locatorExpression-1].value

                    #获取操作值作为函数调用的参数
                    operateValue=rowObj[testStep_operateValue-1].value

                    if isinstance(operateValue,int):
                        operateValue=str(operateValue)

                    if operateValue and operateValue.isalpha():# isalpha()方法检测字符串是否只由字母组成

                        coordiante=operateValue+str(idx+2)


                        operateValue=excelObj.getCellOfValue(dataSourceSheetObj,coordinate=coordiante)


                    tmpStr="'%s','%s'"%(locationType.lower(),locatorExpression.replace("'",'"')) if locationType and locatorExpression else ""



                    operateValue = str(operateValue)
                    if tmpStr:
                        tmpStr+=",'"+operateValue+"'" if operateValue else ""

                    else:

                        tmpStr+="'"+operateValue+"'" if operateValue else ""

                    runStr=keyWord+"("+tmpStr+")"


                    try:

                        if operateValue!='否':

                            eval(runStr)
                    except Exception as e:
                        logging.info('执行步骤%s时发生异常'%rowObj[testStep_testStepDescribe-1].value,traceback.print_exc())


                    else:
                        successStep+=1
                        logging.info('执行步骤%s成功'%rowObj[testStep_testStepDescribe-1].value)

                if stepRowNums==successStep+1:
                    successDatas+=1

                    writeTestResult(sheetObj=dataSourceSheetObj,rowNo=idx+2,colsNo="dataSheet",testResult='pass')

                else:
                    #写入失败信息
                    writeTestResult(sheetObj=dataSourceSheetObj,rowNo=idx+2,colsNo="dataSheet",testResult='faild')

            else:
                #将不要执行的数据行的执行时间和执行结果单元格清空
                writeTestResult(sheetObj=dataSourceSheetObj,rowNo=idx+2,colsNo="dataSheet",testResult='')

        if requiredDatas==successDatas:
            '''
            只有当成功执行的数据条数等于被设置为需要执行的数
            据条数，才表示调用数据驱动的测试用例执行通过
            '''
            return 1
        return 0
    except Exception as e:
        raise e









