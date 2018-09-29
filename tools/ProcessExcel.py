#! /usr/bin/env python
# -*- coding:utf-8 -*-
##############################################################################################
# * purpose:  处理excel考勤文件，拆分表格发送给相应的人
# * author :  wisdom/zhiwei.yin@yingheying.com
# * date   :  2018-09-28
# * readme:
# * 需要将手动填写的变量提前赋值
##############################################################################################
import xlrd
import xlwt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
################################################################################################
#以下内容需要手动填写
#邮箱登录名
username = 'zhiwei.yin@yingheying.com'
#邮箱密码
password='!QAZ!QAZ1qaz'
#发送显示的用户名称，邮箱需要与登录邮箱一致
mailFrom='wisdom <zhiwei.yin@yingheying.com>'/
sender='zhiwei.yin@yingheying.com'

#邮件主题
gEmailSubject = "请添加标题"
#邮件正文内容
gEmailBodyHeader = "你好，请添加正文！\n\n"

#文件所在目录与文件名
#文件所在文件夹
gDirectory = "D:/python/tools/excelFile/"
#文件名
gFileName = "testdata.xls"
#处理完后显示的文件名
gProcessedFileName = "processed.xls"

#本月的工作日1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31
gDutyTup = (1,4,5,6,7,8,11,12,13,14,15,19,20,21,22,25,26,27,28,29,30)
################################################################################################

#邮箱服务器与登录账号
smtpserver = 'smtp.quncaotech.com'
gMailHand = smtplib.SMTP()
gErrorFile = ''
gProcessFile = ''

#yunhui.mou@yingheying.com, "2":"zhiwei.yin@yingheying.com", "3":"zhiwei.yin@yingheying.com"
#"2":"yunhui.mou@yingheying.com", "3":"yunhui.mou@yingheying.com"
gDictEmial = {"1":"zhiwei.yin@yingheying.com", "2":"zhiwei.yin@yingheying.com","3":"zhiwei.yin@yingheying.com"}


def loginInMail(user, psw):
    print "loginInMail....."
    global gMailHand
    gMailHand = smtplib.SMTP()
    gMailHand.connect(smtpserver)
    gMailHand.login(user, psw)
    print "loginInMail[", user, "] successful."

def DisConnectMail():
    global gMailHand
    gMailHand.quit()
    print "DisConnectMail successful."


#设置excel显示格式, pattern.pattern_fore_colour的值如下：
# May be: 8 through 63. 0 = Black, 1 = White, 2 = Red,
# 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan,
# 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown),
# 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
def SetStyle(default=""):
    style = xlwt.XFStyle()  # Create the Style
    pattern = xlwt.Pattern()  # Create the Pattern 背景色
    # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    if default == "":
        pattern.pattern_fore_colour = 1
    elif default == "red":
        pattern.pattern_fore_colour = 2
    elif default == "yellow":
        pattern.pattern_fore_colour = 5
    elif default == "green":
        pattern.pattern_fore_colour = 3
    style.pattern = pattern  # Add Pattern to Style

    #增加边框
    borders = xlwt.Borders()  # Create Borders
    # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED,
    # THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED,
    # SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    style.borders = borders  # Add Borders to Style

    #设置对齐方式
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 垂直对齐
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 水平对齐
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行
    style.alignment = alignment
    return style



def WriteToSheet(writeSheet, sheet, rowBegin, rowEnd, col, style):
    for row in range(rowBegin, rowEnd):
        writeSheet.write(row, col, label=sheet.cell_value(row, col), style=style)

# 定义style，防止报错，More than 4094 XFs (styles)
# 对于大量数据写入excel文件,如果使用了表格样式而且在循环中定义了样式，
# 就是产生了easyxf对象,那么最多只能新建4094个对象
defaultStyle = SetStyle("")
redStyle = SetStyle("red")
yellowStyle = SetStyle("yellow")
greenStyle = SetStyle("green")
def ProcessFile(writeSheet, sheet, startLine, endLine):
    #title, name and date
    for col in range(1, sheet.ncols):
        WriteToSheet(writeSheet, sheet, startLine, startLine+2, col, yellowStyle)

    for col in range(1, sheet.ncols):
        cellStartData = sheet.cell_value(startLine+2, col)
        cellStartData = cellStartData.strip()
        cellEndData = sheet.cell_value(startLine+2, col)
        cellEndData = cellEndData.strip()
        for row in range(startLine+3, endLine):
            cellEndDatatmp = sheet.cell_value(row, col)
            if cellEndDatatmp.strip() != '':
                cellEndData = cellEndDatatmp.strip()

        startTime = ''
        endTime = ''
        if len(cellEndData) == 5:
            startTime = cellStartData[:5].encode('utf8')
            endTime = cellEndData[:5].encode('utf8')
        else:
            startTime = cellStartData[:5].encode('utf8')
            endTime = cellEndData[-5:].encode('utf8')

        aaa = startTime[:2]
        if aaa == '': #没有记录
            if col in gDutyTup:
                WriteToSheet(writeSheet, sheet, startLine+2, endLine, col, redStyle)
            else:
                WriteToSheet(writeSheet, sheet, startLine + 2, endLine, col, defaultStyle)
            continue
        startHour = int(aaa)*60 + int(startTime[-2:])
        endMinHour = int(endTime[:2])*60 + int(endTime[-2:])

        if startHour > 10*60 or endMinHour-startHour < 9*60: #时间不对
            if col in gDutyTup:
                WriteToSheet(writeSheet, sheet, startLine+2, endLine, col, redStyle)
            else:
                WriteToSheet(writeSheet, sheet, startLine + 2, endLine, col, greenStyle)
        else:#时间正常
            if col in gDutyTup:
                WriteToSheet(writeSheet, sheet, startLine+2, endLine, col, defaultStyle)
            else:
                WriteToSheet(writeSheet, sheet, startLine + 2, endLine, col, greenStyle)



def FormatBody(sheet, startLine, endLine):
    header = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>'
    header += '<title>' + gEmailSubject + '</title>'
    body = '''
    <body>
    <div id="container">
      <p><strong text="#000000">''' + gEmailBodyHeader + '''</strong></p>
      <div id="content">
         <table width="1200" border="2" bordercolor="#000000" cellspacing="2" "table-layout"="fixed" align="left">
      <tr bgcolor="#F79646" align="left">'''

    for col in range(1, sheet.ncols):
        cellData = sheet.cell_value(startLine, col)
        if cellData:
            body += '<td><strong>' + cellData.encode('utf8') + '</strong></td>'
        else:
            body += '<td><strong></strong></td>'
    body += '</tr>'

    body += '<tr bgcolor="#F79646" align="left">'
    for col in range(1, sheet.ncols):
        body += '<td><strong>' + str(col) + '</strong></td>'
    body += '</tr>'

    for row in range(startLine+2, endLine) :
        body += '<tr bgcolor="#CFCFCF" align="left">'
        for col in range(1, sheet.ncols):
            cellData = sheet.cell_value(row, col)
            if cellData:
                body += '<td>' + cellData.encode('utf8') + '</td>'
            else:
                body += '<td></td>'

        body += '</tr>'
    body += '''
        </table>
      </div>
    </div>
    </div>
    </body>'''

    tail = '''</html>'''
    html = header + body + tail
    return html


def SendEmail(employNo, sheet, startLine, endLine):
    global gMailHand, gErrorFile
    if endLine - 2 < startLine:
        print "\033[1;35mLine[", startLine, "--", endLine, "] Process employNo[", employNo, "], email[", gDictEmial[employNo], "] failed. \033[0m"
        print >> gErrorFile, "Line[", startLine, "--", endLine, "] Process employNo[", employNo, "], email[", gDictEmial[employNo], "] failed."
        return

    if gDictEmial.has_key(employNo) == False:
        print "\033[1;35mLine[", startLine, "--", endLine, "] Process employNo[", employNo, "] has no email. \033[0m"
        print >> gErrorFile, "Line[", startLine, "--", endLine, "] Process employNo[", employNo, "] has no email."
        return

    msg = MIMEMultipart('related')
    msg['Subject'] = gEmailSubject
    msg['From'] = mailFrom
    msg['To'] = gDictEmial[employNo]

    htmlMsg = FormatBody(sheet, startLine, endLine)
    # 构造表格格式体
    context = MIMEText(htmlMsg, _subtype='html', _charset='utf-8')  # 解决乱码
    msg.attach(context)
    gMailHand.sendmail(sender, gDictEmial[employNo], msg.as_string())
    print "Line[",startLine+1, "--", endLine, "] Send employNo:", employNo, ", email:", gDictEmial[employNo], " successful."

def ProcessExcelFile():
    #open file
    filePath = gDirectory + gFileName
    workbook = xlrd.open_workbook(filePath)
    sheet = workbook.sheet_by_index(0)
    print sheet.name, sheet.nrows, sheet.ncols

    #for write
    writeBook = xlwt.Workbook(encoding='utf-8')
    writeSheet = writeBook.add_sheet("wisdom")

    startLine = 4
    endLine = 4
    employNo = ""
    for row in range(0, sheet.nrows):
        for col in range(0, sheet.ncols):
            if sheet.cell_value(row, col) == u"工号：":
                endLine = row
                if (endLine > startLine):
                    SendEmail(employNo, sheet, startLine, endLine)
                    ProcessFile(writeSheet, sheet, startLine, endLine)
                employNo = str(sheet.cell_value(row, col + 2)) #工号在后面隔一列
                startLine = endLine
                break
    SendEmail(employNo, sheet, startLine, sheet.nrows)
    ProcessFile(writeSheet, sheet, startLine, sheet.nrows)
    #save to file
    filePath = gDirectory + gProcessedFileName
    writeBook.save(filePath)


def main():
    global gErrorFile
    print("Begin Process excel:")
    errorFilePath = gDirectory + "errormsg.txt"
    gErrorFile = open(errorFilePath, 'w')
    loginInMail(username, password)
    ProcessExcelFile()
    DisConnectMail()
    gErrorFile.close()
    print("End Process excel!")


if __name__ == '__main__':
    main()