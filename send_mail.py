#-*- coding:utf-8 -*-
# Create by MaLimin
# Create on 2020/3/19

import zmail
import logging

#设置邮件内容
with open("C:\\code\\Hbuild\\pic.html",'r',encoding='utf-8') as h:
    content_html = h.read()
mail_content={
    "subject":"【银行系统】RPA处理结果",
    # "headers":"处理结果为",
    'content_html': content_html,
    "attachments":["C:\code\python_mail\第1次执行结果统计.txt",
                   "C:\code\python_mail\第1次执行结果统计 - 副本.txt",
                   ],
}
try:
    #第一个参数：邮箱账号，第二个参数：授权码
    server=zmail.server("1052163211@qq.com","aumisagqwyfabcaf",
                        # smtp_port=25,smtp_ssl=False,
                        )
    #发送邮件，第一个参数：收件人地址，第二个参数：要发送的内容
    server.send_mail(["18911411264@163.com"],mail_content)
except Exception as e:
    print("发送失败,原因是:",e)
else:
    print("发送成功！")