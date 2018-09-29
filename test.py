# -*- coding: utf-8 -*-
import xlrd
import os,time
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import base64
#处理从excel中读取的float类型数据的类
#目前集成两种处理：（1）float到int型的转换（2）float到str型的转换，后续有需要可以增加方法以集成其他类型的转换
class judgeFloat:
    def floatToInt(self, variable):
        variable="%d"%variable
        return variable
    def floatToStr(self, variable):
        variable=xlrd.xldate_as_tuple(variable, 0)
        variable=list(variable)
        if variable[1] < 10:
            variable[1] = '0'+str(variable[1])
            variable=str(variable[0])+str(variable[1])+str(variable[2])
        return variable


#写邮件的函数
def mailWrite(filename,address):
    header='<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>'
    th='<body text="#000000">committed缺陷详情：<table border="1" cellspacing="0" cellpadding="3" bordercolor="#000000" width="1800" align="left" ><tr bgcolor="#F79646" align="left" ><th>标识</th><th>摘要</th><th>状态</th><th>优先级</th><th>严重性</th><th>标记</th><th>所有者</th><th>创建时间</th><th>修改时间</th></tr>'
    #打开文件
    filepath = address+filename
    book=xlrd.open_workbook(filepath)
    sheet=book.sheet_by_ind ex(0)
    #获取行列的数目，并以此为范围遍历获取单元数据
    nrows = sheet.nrows-1
    ncols = sheet.ncols
    body=''
    cellData=1
    for i in range(1,nrows+1):
        td=''
        for j in range(ncols):
    #读取单元格数据，赋给cellData变量供写入HTML表格中
            cellData=sheet.cell_value(i,j)
    #调用浮点型转换方法解决读取的日期内容为浮点型数据的问题
            if isinstance(cellData,float):
                if j==0 and i>0:
                    cellDataNew=judgeFloat()
                    cellData=cellDataNew.floatToInt(cellData)
                else:
                    cellDataNew=judgeFloat()
                    cellData=cellDataNew.floatToStr(cellData)
            else:
                pass
            tip='<td>'+cellData+'</td>'
        #并入tr标签
        td=td+tip
    tr='<tr>'+td+'</tr>'
    #为解决字符串拼接问题增设语句，tr从excel中读取出来是unicode编码，转换成UTF-8编码即可拼接
    tr=tr.encode('utf-8')
    #并入body标签
    body=body+tr
    tail='</table></body></html>'
    #将内容拼接成完整的HTML文档
    mail=header+th+body+tail
    return mail


    #发送邮件
def mailSend(mail):
    #设置发件人
    sender = '***'
    #设置接收人
    receiver = '***@***.com'
    #设置邮件主题
    subject = '测试邮件，请忽略！'
    #设置发件服务器，即smtp服务器
    smtpserver = 'smtp.***.net'
    #设置登陆名称
    username = '***@***.net'
    #设置登陆密码
    password = '******'
    #实例化写邮件到正文区，邮件正文区需要以HTML文档形式写入
    msg = MIMEText(mail,'html','utf-8')
    #输入主题
    msg['Subject'] = subject
    #调用邮件发送方法，需配合导入邮件相关模块
    smtp = smtplib.SMTP()
    #设置连接发件服务器
    smtp.connect('smtp.***.net')
    #输入用户名，密码，登陆服务器
    smtp.login(username, password)
    #发送邮件
    smtp.sendmail(sender, receiver, msg.as_string())
    #退出登陆并关闭与发件服务器的连接
    smtp.quit()
    #入口函数，配置文件地址和文件名

def main():
    filename='Sheet1.xlsx'
    address='d:/defectManage/'
    #openFile(filename,address)
    mail=mailWrite(filename,address)
    mailSend(mail)


#调用执行main函数
if __name__=="__main__":
    main()
