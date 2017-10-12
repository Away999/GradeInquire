# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr,formataddr
import smtplib


#学生信息
class Student(object):
    def __init__(self,StuNo,StuId):
        self.StuNo = '159074423' #input("请输入学号:")
        self.StuId = '341126199703290219' #input("请输入身份证号:")

# 邮箱信息
class MailInfo(object):
    def __init__(self, from_addr, password, to_addr, smtp_server):
        self.from_addr = '619132253@qq.com'
        self.password = 'zcvhqqmrnauvbcbg'
        self.to_addr = 'weixiaonienie@qq.com'
        self.smtp_server = 'smtp.qq.com'

#格式化邮件地址
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))

#设置邮件抬头
def SetMailHead(recipient,grade,mailinfo):
    msg = MIMEText(grade,'plain','utf-8')
    msg['From'] = _format_addr('管理员<%s>' %mailinfo.from_addr)
    msg['To'] = _format_addr(recipient+'<%s>'%mailinfo.to_addr)
    msg['Subject'] = Header(recipient+'的成绩单','utf-8').encode()
    return msg

#发送邮件
def SendEmail(mailinfo,msg):
    server = smtplib.SMTP_SSL(mailinfo.smtp_server,465)
    server.set_debuglevel(1)
    server.login(mailinfo.from_addr,mailinfo.password)
    server.sendmail(mailinfo.from_addr,[mailinfo.to_addr],msg.as_string())
    server.quit()

#保存到本地
def SavetoLocal(title,contents):
    file = open(title+'.txt','w',encoding='utf-8')
    file.write(contents)
    file.close()

#读取成绩
def GetScoreHtml(StuNo,StuId):
    browser = webdriver.Chrome()
    browser.get('http://211.70.149.134:8080/stud_score/brow_stud_score.aspx')
    elem_StuNo = browser.find_element_by_id("TextBox1")
    elem_StuNo.send_keys(StuNo)
    elem_IdNum = browser.find_element_by_id("TextBox2")
    elem_IdNum.send_keys(StuId)
    elem_XN = browser.find_element_by_id("drop_xn")
    elem_XN.send_keys("2016-2017")
    elem_XQ = browser.find_element_by_id("drop_xq")
    elem_XQ.send_keys("2")
    cjcx = browser.find_element_by_id("Button_cjcx")
    cjcx.click()
    return browser

#打印成绩单
def PrintScore(browser):
    ScoreHtml = browser.page_source
    ScoreBFSP = BeautifulSoup(ScoreHtml,'lxml')
    ScoreList = ScoreBFSP.find_all(id='GridView_cj')
    ScoreText = BeautifulSoup(str(ScoreList),'lxml')
    grade = {}
    StuInfo = BeautifulSoup(str(ScoreBFSP.find_all(id = 'Label1')),'lxml')
    stustring = StuInfo.text
    gradestring = stustring + "\n"
    for scoreitem in ScoreText.find_all(align = 'left'):
        scoreinfo = BeautifulSoup(str(scoreitem),'lxml')
        i = 0
        for tdd in scoreinfo.find_all('td'):
            if i == 3:
                if tdd.text != u'\xa0':
                    key = tdd.text
            elif i == 8:
                if tdd.text != u'\xa0':
                    grade[key] = tdd.text
            i = i + 1
    for k,v in grade.items():
        #print("科目:" +k)
        gradestring = gradestring + "科目：" + k + "\n"
        #print("成绩:"+v)
        gradestring = gradestring + "成绩：" + v + "\n"

    return gradestring

if __name__ == '__main__':
    StuInfo = Student('','')
    browser = GetScoreHtml(StuInfo.StuNo,StuInfo.StuId)
    grade = PrintScore(browser)
    browser.close()
    mailinfo = MailInfo('','','','')
    msg = SetMailHead(grade,mailinfo)
    SendEmail(mailinfo,msg)