#! /usr/bin/env python
# -*- coding:utf-8 -*-

import time
from websocket import create_connection

#消息类型 Event_Light_Off，Event_Light_On， Event_Light_Query， Event_Light_Push
msg = '''{"type" : "Event_Light_Off","data":{"type" : 1,"reason" : "for python test.","site" : [{"siteId":3560}, {"siteId":3561}, {"siteId":3801}, {"siteId":3802}, {"siteId":3803}, {"siteId":4151}, {"siteId":4854}, {"siteId":4855}]}}'''

def SendWebSocket(addr, sendMsg):
    ws = create_connection(addr)
    print("Sent:", msg)
    ws.send(msg)

    result = ws.recv()
    print("Received: ",result)
    time.sleep(3000)
    ws.close()

if __name__ == '__main__':
    print "Sent Msg to ws://127.0.0.1:9002/"
    SendWebSocket("ws://127.0.0.1:9002/", msg)
    #SendWebSocket("ws://10.0.2.15:9002/", msg)