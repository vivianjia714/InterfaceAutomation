#coding=utf-8

#登录接口
#post请求 fiddler抓包中name为jsonParams value为json请求体，返回为json对象，后转码为中文

import requests

import json

import sys

from email.header import UTF8

reload(sys) 
 
sys.setdefaultencoding('utf8')   

payload = "jsonParams=%s" % json.dumps({"username":"13942005807","G":{"at":"0","sv":"0200","av":"320","dt":"2","deviceId":"97f576754c8b1a33ef98259d2eb2943540a5ceff"},"password":"111111"})


headers = {"Content-Type": "application/x-www-form-urlencoded"} 

r = requests.post("http://innerapi1.jiemodou.net/User/login",data = payload, headers=headers)

print r.status_code

r1 = r.json()

print json.dumps(r1, encoding="UTF-8", ensure_ascii=False)

file = open("E:\\test environmen\\jiemo\\src\\jiemo_interface\\file\\login.json",'w+')

file.write(json.dumps(r1,encoding="UTF-8", ensure_ascii=False))

file.flush()

file.close()