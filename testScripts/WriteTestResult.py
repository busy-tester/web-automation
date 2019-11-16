import sys
import os
path="F:\\CJMDXTtest\\"
sys.path.append(path)

#往Excel中写入测试结果信息的公共方法

# from testScripts import *
from testScripts.Excel_Obj import *
# from . import *


#用例或用例步骤执行结束后，向Excel中写执行结果信息
def writeTestResult(sheetObj,rowNo,colsNo,testResult,errorInfo=None,picPath=None):
    #测试通过结果信息为绿色，失败为红色
    colorDict={"pass":"green","faild":"red","":None}

    '''
    因为"测试用例"工作表和"用例步骤sheet表"中都有测试执行时间和
    测试结果列，定义此字典对象是为了区分具体应该写哪个表
    '''
    colsDict={
        "testCase":[testCase_runTime,testCase_testResult],
        "caseStep":[testStep_runTime,testStep_testResult],
        "dataSheet":[dataSource_runTime,dataSource_result]#联系人表
    }
    try:
        #在测试步骤sheet中，写入测试结果,testResult为pass或faild
        excelObj.writeCell(sheetObj,content=testResult,rowNo=rowNo,colsNo=colsDict[colsNo][1],style=colorDict[testResult])
        if testResult=="":
            #清空时间单元格内容
            excelObj.writeCell(sheetObj,content='',rowNo=rowNo,colsNo=colsDict[colsNo][0])

        else:
            #在测试步骤sheet中，写入测试时间
            excelObj.writeCellCurrentTime(sheetObj,rowNo=rowNo,colsNo=colsDict[colsNo][0])

        if errorInfo and picPath:
            #在测试步骤sheet中，写入异常信息
            excelObj.writeCell(sheetObj,content=errorInfo,rowNo=rowNo,colsNo=testStep_errorInfo)

            #在测试步骤sheet中，写入异常截图路径
            excelObj.writeCell(sheetObj,content=picPath,rowNo=rowNo,colsNo=testStep_errorPic)

        else:
            if colsNo=="caseStep":
                #在测试步骤sheet中，清空异常信息单元格
                excelObj.writeCell(sheetObj,content='',rowNo=rowNo,colsNo=testStep_errorInfo)

                # 在测试步骤sheet中，清空异常图片路径单元格
                excelObj.writeCell(sheetObj, content='', rowNo=rowNo, colsNo=testStep_errorPic)

    except Exception as e:
        print('写excel时发生错误')
        print(traceback.print_exc())

