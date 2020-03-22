#-*- coding:utf-8 -*-
# Create by MaLimin
# Create on 2020/3/20

import base64
def pic_to_base64(pic_path = None):
    """
    将图片转换为base64编码

    param pic_path: 参数1.图片绝对路径

    return: 返回base64编码
    """
    path = pic_path
    with open(r"C:\code\Hbuild\img\1.png", 'rb') as f:
        base64_data = base64.b64encode(f.read())
        s = base64_data.decode()
        print('data:image/jpeg;base64,%s' % s)
        return s
pic_to_base64()