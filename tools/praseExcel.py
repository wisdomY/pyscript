#! /usr/bin/env python
# -*- coding:utf-8 -*-

import xlrd
import xlwt
from datetime import date,datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header

# 设置smtplib所需的参数
# 下面的发件人，收件人是用于邮件传输的。
# smtpserver = 'smtp.163.com'
# username = 'yzw5164129@163.com'
# password='wisdomyzw412516'
#
#
# mailFrom='wisdom <yzw5164129@163.com>'
# sender='yzw5164129@163.com'
# receiver='zhiwei.yin@yingheying.com'
#收件人为多个收件人
#receiver=['zhiwei.yin@yingheying.com']

smtpserver = 'smtp.quncaotech.com'
username = 'zhiwei.yin@yingheying.com'
password='!QAZ!QAZ1qaz'
mailFrom='wisdom <zhiwei.yin@yingheying.com>'
sender='zhiwei.yin@yingheying.com'

#邮件主题
gEmailSubject = "Subject"
#邮件正文
gEmailBody = "    email body:\n    你好，这是测试邮件！\n\n\n    Best Regards!"

#文件所在目录
gDirectory = "D:/python/tools/excelFile/"

#yunhui.mou@yingheying.com
gDictEmial = {"2":"yzw5164129@163.com", "3":"yzw5164129@163.com"}

def GenerateMailFormat(receiver, fileName):
    # 构造邮件对象MIMEMultipart对象
    # 下面的主题，发件人，收件人，日期是显示在邮件页面上的。
    msg = MIMEMultipart('mixed')
    msg['Subject'] = gEmailSubject
    msg['From'] = mailFrom
    msg['To'] = receiver
    # 收件人为多个收件人,通过join将列表转换为以;为间隔的字符串
    #msg['To'] = ";".join(receiver)

    # 构造文字内容
    text_plain = MIMEText(gEmailBody, 'plain', 'utf-8')
    msg.attach(text_plain)

    # 构造附件
    filePath = gDirectory + fileName
    sendfile = open(filePath, 'rb').read()
    text_att = MIMEText(sendfile, 'base64', 'utf-8')
    text_att["Content-Type"] = 'application/octet-stream'
    # 以下附件可以重命名成aaa.txt
    text_att["Content-Disposition"] = 'attachment; filename="kaoqin.xls"'
    # 另一种实现方式
    #text_att.add_header('Content-Disposition', 'attachment', filename='ttt.xls')
    # 以下中文测试不ok
    # text_att["Content-Disposition"] = u'attachment; filename="中文附件.xls"'.decode('utf-8')
    msg.attach(text_att)
    return msg


def SendEmail(receiver, fileName):
    msg = GenerateMailFormat(receiver, fileName)
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver)
    smtp.login(username, password)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()

#设置excel显示格式
def SetStyle(line=0):
    style = xlwt.XFStyle()  # Create the Style
    if line < 2:
        # 字体
        font = xlwt.Font()  # Create the Font
        font.name = 'Times New Roman'
        font.bold = True
        font.italic = True
        style.font = font  # Apply the Font to the Style

    if line < 2:
        #背景色
        pattern = xlwt.Pattern()  # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        style.pattern = pattern  # Add Pattern to Style

    if (line > 0) :
        #边框
        borders = xlwt.Borders()  # Create Borders
        borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        borders.left_colour = 0x40
        borders.right_colour = 0x40
        borders.top_colour = 0x40
        borders.bottom_colour = 0x40
        style.borders = borders  # Add Borders to Style
    return style


def WriteExcel(fileName, sheet, startLine, endLine):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet(fileName)
    if endLine - 2 < startLine :
        print "Error Data ", fileName, ", Line range: ", startLine,  "~", endLine
        return

    for row in range(startLine, startLine+2) :
        style = SetStyle(line=row-startLine)
        for col in range(0, sheet.ncols):
            worksheet.write(row-startLine, col, label=sheet.cell_value(row, col), style=style)

    for row in range(startLine+2, endLine) :
        style = SetStyle(line=row-startLine)
        for col in range(0, sheet.ncols):
            worksheet.write(row-startLine, col, label=sheet.cell_value(row, col), style=style)

    fileWithEnds = fileName + ".xls"
    filePath = gDirectory + fileWithEnds
    workbook.save(filePath)
    SendEmail(gDictEmial[fileName], fileWithEnds)
    print fileWithEnds, gDictEmial[fileName]


def ReadExcel():
    #open file
    filePath = gDirectory + "testdata.xlsx"
    workbook = xlrd.open_workbook(filePath)
    print workbook.sheet_names()
    #sheet = workbook.sheet_by_name('excelTest')
    sheet = workbook.sheet_by_index(0)
    print sheet.name, sheet.nrows, sheet.ncols

    startLine = 0
    endLine = 0
    fileName = ""
    for row in range(0, sheet.nrows):
        for col in range(0, sheet.ncols):
            if sheet.cell_value(row, col) == u"工号：":
                endLine = row
                if (endLine > startLine):
                    WriteExcel(fileName, sheet, startLine, endLine)
                fileName = str(sheet.cell_value(row, col + 2))
                startLine = endLine
                break
    WriteExcel(fileName, sheet, startLine, sheet.nrows)



if __name__ == '__main__':
    print("Begin Process excel:")
    ReadExcel()
    print("End Process excel!")