# -*- coding: utf-8 -*-
import easygui as g
#邮箱登录名
username = 'zhiwei.yin@yingheying.com'
#邮箱密码
password='!QAZ!QAZ1qaz'
#发送显示的用户名称，邮箱需要与登录邮箱一致
mailFrom='wisdom <zhiwei.yin@yingheying.com>'
sender='zhiwei.yin@yingheying.com'

#邮件主题
gEmailSubject = "9月份考勤数据"
#邮件正文内容
gEmailBodyHeader = "您好，以下是您9月份的考勤数据！\n\n"

#文件所在目录与文件名
#文件所在文件夹
gDirectory = "D:/python/tools/excelFile/"
#文件名
gFileName = "tttt.xls"
#员工工号与邮箱的对应数据
gEmployDataFile = "employdata.xlsx"
#处理完后显示的文件名
gProcessedFileName = "processed.xls"

#本月的工作日1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31
gDutyTup = [32,33]

def Login():
    psd = g.passwordbox(msg='请直接输入密码登录：', title='Logining')
    if psd.upper()== 'WISDOM':
        return True
    else:
        return False

def InitDialog():
    global username, password, mailFrom, gEmailSubject, gEmailBodyHeader, gDirectory, gFileName, gEmployDataFile, gProcessedFileName, gDutyTup
    strStr = ('邮件主题:', '邮件正文:', '文件所在目录: 例如：D:/files/', '工号邮箱对照表文件名:', '考勤文件名:', '保存文件:', '使用的发送邮箱:', '发送邮箱的密码:', '发送邮件的昵称:', '本月非工作日日期:')
    defualtValue = ('上月考勤数据', '您好，以下是您上月考勤数据!', 'D:/python/tools/excelFile/', 'employdata.xlsx', 'tttt.xls', 'processed.xls', 'zhiwei.yin@yingheying.com', '!QAZ!QAZ1qaz','wisdom', '3,4,5,6,7')
    if Login():
        (gEmailSubject, gEmailBodyHeader, gDirectory, gEmployDataFile, gFileName, gProcessedFileName, username, \
            password, mailFrom, dutyStr) = g.multenterbox('考勤系统\n以下参数均有默认值，且均为必填项\n*文件目录与文件名不支持中文', '泉眼科技', strStr, defualtValue)
        if gEmailSubject.strip() == '' or gEmailBodyHeader.strip() == '' or gDirectory.strip() == ''\
                or gFileName.strip() == '' or gEmployDataFile.strip() == '' or gProcessedFileName.strip() == '' \
                or username.strip() == '' or password.strip() == '' or mailFrom.strip() == '' or dutyStr.strip() == '':
            print "All entry can not be empty"
            return False
        else:
            gEmailSubject = gEmailSubject.encode(encoding='UTF-8')
            gEmailBodyHeader = gEmailBodyHeader.encode(encoding='UTF-8')
            gDirectory = gDirectory.encode(encoding='UTF-8')
            gEmployDataFile = gEmployDataFile.encode(encoding='UTF-8')
            gFileName = gFileName.encode(encoding='UTF-8')
            gProcessedFileName = gProcessedFileName.encode(encoding='UTF-8')
            username = username.encode(encoding='UTF-8')
            password = password.encode(encoding='UTF-8')
            mailFrom = mailFrom.encode(encoding='UTF-8')
            dutyStr = dutyStr.encode(encoding='UTF-8')
            mailFrom += ' <' + username + '>'
            for day in dutyStr.split(","):
                gDutyTup.append(int(day))

            print gDutyTup
            return True
    return False

def main():
    print "main func begins:";
    InitDialog()
    print '邮件主题:', gEmailSubject, '邮件正文:', gEmailBodyHeader, '文件所在目录:', gDirectory, '工号邮箱对照表文件名:', gEmployDataFile,\
        '考勤文件名:',gFileName, '保存文件:',gProcessedFileName, '使用的发送邮箱:', username, '发送邮箱的密码:', password, '发送邮件的昵称:', mailFrom, '本月非工作日日期:', gDutyTup

#调用执行main函数
if __name__=="__main__":
    main()
