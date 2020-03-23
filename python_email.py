#-*- coding:utf-8 -*-
# Create by MaLimin
# Create on 2020/3/22

import zmail
import configparser
import logging
import os


# 存放 【收件人】和【抄送人】邮箱地址 文件
RPA_add_path = "C:\code\python_mail\新建文件夹\email.txt"
# 存放 【数据处理结果】【附件路径】 的文件
RPA_data_path = "D:\RPA\codes\rpa_code\1_setting\email_data.txt"
# 存放 【运行记录】 的路径
RPA_record_path = "D:\RPA\codes\rpa_code\1_setting\email_data_record.txt"

global rec  # 收件人
global cc  # 抄送人
global email_type # 数据处理结果
global email_atta_path # 处理结果附件路径
global email_atta_record_path # RPA运行记录附件路径
global email_content_text # 邮件正文内容
# rec = ""
# cc = ""
# email_type = ""
# email_atta_path = ""
# email_atta_record_path = ""
# email_content_text = ""

def read_record_path():
    """读取RPA运行记录文件的路径"""
    global email_atta_record_path
    with open(RPA_record_path, "r", encoding="utf-8") as f:
        email_atta_record_path = f.read()
    if not os.path.exists(email_atta_record_path):
        print("RPA运行记录文件【不存在】")
        raise(("RPA运行记录文件【不存在】"))


def read_email_data():
    """读取数据处理的结果，及处理结果附件路径"""
    with open(RPA_data_path, "r", encoding="utf-8") as f:
        email_data = f.read()
    email_data = email_data.split("@")

    global email_type  # 数据处理结果
    global email_atta_path  # 处理结果附件路径
    global email_content_text  # 邮件正文内容

    email_type = email_data[0]
    email_atta_path = email_data[1]

    # 根据处理结果，设定发送邮件的正文内容
    if int(email_type) == 0:
        email_content_text = "本次转换明细为【0】条，附件处理结果文档内容为空！"
    elif int(email_type) > 0:
        email_content_text = "本次转换明细为【%s】条，详细内容见附件。" %email_type
    elif int(email_type) == -1:
        email_content_text = "下载文件已损坏，无法完成数据转换！"
        email_atta_path = "D:\RPA\download\%s\银行账户明细表.xls" %email_content_text
    else:
        print("处理结果格式不正确，无法获取话术")
        raise("处理结果格式不正确，无法获取话术")

    # 判断处理结果附件是否存在
    if not os.path.exists(email_atta_path):
        print("处理结果不存在")
        raise("处理结果不存在")

def read_message():
    """读取【收件人】和【抄送人】邮箱"""
    with open(RPA_add_path, "r", encoding="utf-8-sig") as f:
        address_all = f.read()
    address_all = address_all.split(">")
    global rec  # 收件人
    global cc  # 抄送人
    if len(address_all) == 2:
        rec = address_all[0]
        cc = address_all[1]
    elif len(address_all) == 1:
        rec = address_all[0]
        cc = ""
    else:
        print("收件人或抄送人设置不正确")
        raise("收件人或抄送人设置不正确")

    # 将收件人和抄送人写为列表格式
    rec = str_to_list(rec)
    cc = str_to_list(cc)
    if cc == []:
        cc = None

    print(rec)
    print(cc)

def str_to_list(old_str):
    """将地址的字符串转化为列表返回，并去除列表中的空值"""
    new_list = old_str.split(";")

    # 遍历删除列表中的空元素
    for data in new_list[::-1]:
        if data == "":
            new_list.remove(data)

    # 便利给列表中的元素加上‘;’号
    for i in range(0,len(new_list)):
        new_list[i] = str(new_list[i]) + ";"

    return new_list


def sent_email():
    """发送邮件"""
    #设置邮件内容
    mail_content={
        "subject":"【银行系统】RPA处理结果",
        # "headers":"处理结果为",
        'content_text': email_content_text,
        "attachments":[email_atta_path,
                       email_atta_record_path,
                       ],
    }
    try:
        #第一个参数：邮箱账号，第二个参数：授权码
        server=zmail.server("1052163211@qq.com",
                            "aumisagqwyfabcaf",
                            # smtp_port=25,
                            # smtp_ssl=False,
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

def read_config(config, email_ini_dir):
    config.read(email_ini_dir, encoding="utf-16")
    # print(config.get("收件人", "收件人邮箱2"))
    # print(config.items(section="收件人"))
    # print(config.options("收件人"))

    # 收件人
    global rec
    if len(config.options("收件人"))>1:
        print("多个收件人")
        rec = []
        for value in config.items(section="收件人"):
            # print(value[1])
            rec.append(value[1])
        print("rec:", rec)
    elif len(config.options("收件人"))==1:
        print("只有一个收件人")
        rec = config.items(section="收件人")[0][1]
        print(rec)

    pass


if __name__ == '__main__':
    # 获取配置文件路径
    email_ini_dir =  get_email_ini_dir()


    ############ 读取配置文件 ############

    # 实例化配置文件
    config = configparser.ConfigParser()

    # 向配置文件中写入内容
    # write_config(config, email_ini_dir)

    # 读取配置文件中的内容
    read_config(config, email_ini_dir)



    # read_message()
    # read_email_data()
    # read_record_path()
    # sent_email()
