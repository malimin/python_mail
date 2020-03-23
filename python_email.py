#-*- coding:utf-8 -*-
# Create by MaLimin
# Create on 2020/3/22

import zmail
import configparser
import logging
import os


def send_email(username, password, rec, cc, subject, email_content_text, attachments):
    """发送邮件"""
    #设置邮件内容
    mail_content={
        "subject": subject,
        # "headers":"处理结果为",
        'content_text': email_content_text,
        "attachments":attachments,
    }
    try:
        #第一个参数：邮箱账号，第二个参数：授权码
        server=zmail.server(
            # "1052163211@qq.com;","aumisagqwyfabcaf",
                            username,
                            password,
                            smtp_port = 25,
                            smtp_host = "smtp.qq.com",
                            smtp_ssl = False,
                            )
        #发送邮件，第一个参数：收件人地址u，第二个参数：要发送的内容
        server.send_mail(recipients = rec,mail = mail_content, cc=cc)
    except Exception as e:
        print("发送失败,原因是:",e)
    else:
        print("发送成功！")

def get_email_ini_dir():
    """获取当前文件夹下的配置文件路径"""
    cur_dir = os.getcwd()
    # email_ini_dir = cur_dir + "\python_email.ini"
    email_ini_dir = os.path.join(cur_dir, "python_email.ini")
    email_ini_dir = email_ini_dir.replace("\\", "\\\\")
    # print(email_ini_dir)
    return email_ini_dir

def write_config(config, email_ini_dir):
    config.read(email_ini_dir, encoding="utf-16 ")
    list = config.sections()
    print(list)
    if "收件人" not in list:
        config.add_section("收件人")
        config.set("收件人", "收件人邮箱", "1052163211@qq.com")
        print("写入数据：", config.sections())

        with open(email_ini_dir,"a",encoding="utf-16") as f:
            config.write(f)


def read_normal_config(config, email_ini_dir, option, key = None):
    config.read(email_ini_dir, encoding="utf-16")
    # print(config.get("收件人", "收件人邮箱2"))
    # print(config.items(section="收件人"))
    # print(config.options("收件人"))

    if key == None:
        # 收件人
        # global result
        if len(config.options(option))>1:
            print("多个%s" %option)
            result = []
            for option in config.items(section=option):
                # print(option[1])
                result.append(option[1])
            # print("result:", result)
        elif len(config.options(option))==1:
            print("只有一个%s" %option)
            result = config.items(section=option)[0][1]
            # print(result)
    else:
        result = config.get(option, key)
    return result
if __name__ == '__main__':
    # 获取配置文件路径
    email_ini_dir =  get_email_ini_dir()


    ############ 读取配置文件 ############

    # 实例化配置文件
    config = configparser.ConfigParser()

    # 向配置文件中写入内容
    # write_config(config, email_ini_dir)

    # 读取配置文件中【发件人邮箱】【发件人密码】
    username = read_normal_config(config, email_ini_dir, "发件人", "发件人邮箱")
    print(username)
    password = read_normal_config(config, email_ini_dir, "发件人", "发件人密码")
    print(password)

    # 读取配置文件中【收件人】【抄送人】【主题】【正文】【附件】的内容
    # read_config(config, email_ini_dir)
    rec = read_normal_config(config, email_ini_dir, "收件人")
    print(rec)
    cc = read_normal_config(config, email_ini_dir, "抄送人")
    print(cc)
    subject = read_normal_config(config, email_ini_dir, "主题")
    print(subject)
    email_content_text = read_normal_config(config, email_ini_dir, "正文")
    print(email_content_text)
    attachments = read_normal_config(config, email_ini_dir, "附件")
    print(attachments)

    # 发送邮件
    send_email(username, password, rec, cc, subject, email_content_text, attachments)

