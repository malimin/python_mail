#-*- coding:utf-8 -*-
# Create by MaLimin
# Create on 2020/3/22

import zmail
import configparser
import logging
import os
import import_logging


def str_to_bool(str):
    """参数所需的值为布尔类型，此函数将字符串类型的传入参数转换为布尔类型，并返回"""
    if str.lower() == 'true':
        return True
    elif str.lower() == "false":
        return False
    else:
        logging.warning("smtp_ssl不正确")
        return False


# def send_email(username, password, smtp_port, smtp_host, smtp_ssl, rec, cc, subject, email_content_text, attachments):
def send_email(username, password, rec, cc, subject, email_content_text, attachments, smtp_host: str, smtp_port:int = 25, smtp_ssl:bool = False, ):
    """发送邮件函数

    所需参数为：发件人邮箱，发件人密码，收件人邮箱，抄送人邮箱，主题，正文，附件，smtp_host，smtp_port，smtp_ssl"""
    """发送邮件"""
    #设置邮件内容
    mail_content={
        "subject": subject,
        'content_text': email_content_text,
        "attachments":attachments,
    }
    try:
        # 第 一个参数：邮箱账号，第二个参数：授权码
        server=zmail.server(username,
                            password,
                            smtp_port = smtp_port,
                            smtp_host = smtp_host,
                            smtp_ssl = str_to_bool(smtp_ssl),
                            )
        # 发送邮件，第一个参数：收件人地址，第二个参数：要发送的内容
        server.send_mail(recipients = rec,mail = mail_content, cc=cc)
    except Exception as e:
        logging.warning("发送失败,原因是:",e)
        print("发送失败,原因是:",e)
    else:
        logging.info("发送成功！")
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
    """编辑ini配置文件"""
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
    """获取指定选项下所有值或获取指定选项下的指定键的值，将结果返回"""
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
        elif len(config.options(option)) == 0:
            print("没有%s这个选项" %option)
            result = None
    else:
        try:
            result = config.get(option, key)
        except Exception as e:
            logging.warning("获取发件箱中的--{}--{}--内容失败".format(option, key))
            print("获取发件箱中的--{}--{}--内容失败".format(option, key))

    return result
if __name__ == '__main__':
    logging.info("=============开始邮件处理部分=============")
    logging.info("开始读取邮件配置信息")
    # 获取配置文件路径
    email_ini_dir =  get_email_ini_dir()


    ############ 读取配置文件 ############

    # 实例化配置文件
    config = configparser.ConfigParser()

    # 向配置文件中写入内容
    # write_config(config, email_ini_dir)

    # 读取配置文件中【发件人】选项中的【发件人邮箱】【发件人密码】【smtp_port】【smtp_host】【smtp_ssl】
    username = read_normal_config(config, email_ini_dir, "发件人", "发件人邮箱")
    print(username)
    password = read_normal_config(config, email_ini_dir, "发件人", "发件人密码")
    print(password)
    smtp_port = read_normal_config(config, email_ini_dir, "发件人", "smtp_port")
    print(smtp_port)
    smtp_host = read_normal_config(config, email_ini_dir, "发件人", "smtp_host")
    print(smtp_host)
    smtp_ssl = read_normal_config(config, email_ini_dir, "发件人", "smtp_ssl")
    print(smtp_ssl)

    logging.info("取配置文件中【发件人】选项中的\n【发件人邮箱】{}\n【发件人密码】{}\n【smtp_port】{}\n【smtp_host】{}\n【smtp_ssl】{}".format(username, password, smtp_host, smtp_port, smtp_ssl))

    # 读取配置文件中【收件人】【抄送人】【主题】【正文】【附件】选项的的内容
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

    logging.info("读取配置文件中邮件的内容\n【收件人】{}\n【抄送人】{}\n【主题】{}\n【正文】{}\n【附件】{}".format(rec, cc, subject, email_content_text, attachments))

    # 发送邮件
    # send_email(username, password, smtp_port, smtp_host, smtp_ssl, rec, cc, subject, email_content_text, attachments)
    logging.info("开始发送邮件")
    send_email(username, password, rec, cc, subject, email_content_text, attachments, smtp_host,smtp_port, smtp_ssl)

    logging.info("=============邮件处理部分结束=============\n\n")

