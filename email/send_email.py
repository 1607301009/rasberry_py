# -*- encoding: utf-8 -*-
'''
@Author  :   杨结万
@License :   (C) Copyright 2020-2021,Yang JW
@Contact :   876124112@qq.com
@Software:   PyCharm python3.8
@File    :   python_qq_email.py
@Time    :   2020/3/27 0:30
@Desc    :   QQ邮箱发送程序
'''
import json
import smtplib
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class QqEmail():
    """
    QQ邮箱发送程序,初始化时传入，发送人邮箱和授权码。
    使用 send 方法 发送邮箱带附件的邮箱！
    """

    def __init__(self, msg_from, authorization):
        """
        QQ邮箱需开通 IMAP/SMTP服务
        :param msg_from: 发送方QQ邮箱，如：123456789@qq.com
        :param authorization: 发送方邮箱授权码，如：jfdscohkgtwqbfge。
        """
        self.msg_from = msg_from  # 发送人邮箱
        self.authorization = authorization  # 发送方邮箱的授权码
        # 登录邮箱
        self.s = smtplib.SMTP_SSL("smtp.qq.com", 465)
        self.s.login(msg_from, authorization)

    def send(self, msg_to, **kwargs):
        """
        发送邮箱的方法，使用该方法前，需要首先初始化！
        :param msg_to: 收件方的邮箱。如：987654321@163.com,注意需要与发件邮箱不一致！
        :param kwargs: 发送内容的具体信息，包括
            add_file : 添加的附件，类型为list。
            subject  : 主题， 类型为str。
            text     ： 正文， 类型为str。
        :return: None
        """
        # 创建一个实例
        self.message = MIMEMultipart()
        self.message['From'] = Header(self.msg_from, 'utf-8')
        self.message['To'] = Header(msg_to, 'utf-8')
        if kwargs.get('add_file'):
            # 添加全部附件
            for i in kwargs.get('add_file'):
                self.add_file(i)
        # 添加主题
        if kwargs.get('subject'):
            self.message['Subject'] = Header(kwargs.get('subject'), 'utf-8')
        # 添加正文
        if kwargs.get('text'):
            self.message.attach(MIMEText(kwargs.get('text'), 'plain', 'utf-8'))
        # 开始发送
        self.s.sendmail(self.msg_from, msg_to, self.message.as_string())

    def quit(self):
        # 退出邮箱
        self.s.quit()

    # 给邮件添加附件
    def add_file(self, file_send):
        pdfApart = MIMEApplication(open(file_send, 'rb').read())
        pdfApart.add_header('Content-Disposition', 'attachment', filename=file_send)
        self.message.attach(pdfApart)


def main():
    # 读取数据
    with open('qq_email.json', 'r') as f:
        data = json.load(f)
        msg_from = data['QQEmail']
        authorization = data['authorization']
    email = QqEmail(msg_from, authorization)
    # 收件人邮箱
    msg_to = '873450354@qq.com'
    kwargs = {
        # 'add_file': ['qq_email.json'],
        'subject': "[登录提示]python发QQ邮箱测试成功！",
        'text': 'hi:\n\t\t发送成功！\npython\n',
    }

    email.send(msg_to, **kwargs)
    email.quit()
    print('send ok')
    pass


if __name__ == '__main__':
    main()
