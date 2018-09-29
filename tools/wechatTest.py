#! /usr/bin/env python
# -*- coding:utf-8 -*-

import itchat
import time

gUserName = ''

def SendMsg(msg, name) :
    print "I will send", msg, "to", name
    # 使用备注名来查找实际用户名
    users = itchat.search_friends(name=name)
    # 获取`UserName`,用于发送消息
    userName = users[0]['UserName']
    itchat.send(msg, toUserName=userName)
    print('Send msg successful.')

def SendMsgToCharGroup(msg, groupName):
    #rooms = itchat.get_chatrooms(update=True)  # 拿到所有微信群
    rooms = itchat.search_chatrooms(groupName)  # 搜索指定微信群（模糊搜索）
    userName = rooms[0]['UserName']
    itchat.send(msg, toUserName=userName)


def CountNum():
    friends = itchat.get_friends(update=True)[0:]
    male = female = other = 0
    total = len(friends[1:])
    for i in friends[1:]:
        sex = i["Sex"]
        if sex == 1:
            male += 1
        elif sex == 2:
            female += 1
        else:
            other += 1
            print "other", i["UserName"].encode('utf8')

    print "male:", male, "female:", female, "total:", total
    print u"男性好友：%.2f%%" % (float(male) / total * 100)
    print u"女性好友：%.2f%%" % (float(female) / total * 100)
    print u"其他：%.2f%%" % (float(other) / total * 100)

#自动回复
@itchat.msg_register('Text')
def AutoReply(msg):
    global gUserName
    print "reply:", msg['FromUserName'], "name:", gUserName
    if msg['FromUserName'] == gUserName:
    # 发送一条提示给文件助手
        itchat.send_msg(u"[%s]收到好友@%s 的信息：%s\n" %
            (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(msg['CreateTime'])),
             msg['User']['NickName'], msg['Text']), 'filehelper')
    # 回复给好友
        return u'[自动回复]\n已经收到您的的信息：%s\n' % (msg['Text'])


if __name__=='__main__':
    global gUserName
    # 首次扫描登录后后续自动登录
    itchat.auto_login(hotReload=True)
    username = u'name'
    SendMsg("test", username)
    # SendMsgToCharGroup("This is a test, please ignore", u'缘来一家人')
    # CountNum()
    # myUserName = itchat.get_friends(update=True)[0]["UserName"]
    users = itchat.search_friends(name=username)
    gUserName = users[0]['UserName']
    print "main:", gUserName
    itchat.run()
    print("test")