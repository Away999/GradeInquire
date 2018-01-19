# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr,formataddr
import smtplib
import pickle
import os
import operator

#学生信息
class Student(object):
    def __init__(self,StuNo,StuId):
        self.StuNo = '159074423' #input("请输入学号:")
        self.StuId = '341126199703290219' #input("请输入身份证号:")

# 邮箱信息
class MailInfo(object):
    def __init__(self, from_addr, password, to_addr, smtp_server):
        self.from_addr = '619132253@qq.com'
        self.password = 'hkdbwoljhxzfbcfb'
        self.to_addr = 'weixiaonienie@qq.com'
        self.smtp_server = 'smtp.qq.com'

#格式化邮件地址
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))

#设置邮件抬头
def SetMailHead(recipient,newgrade,grade,mailinfo):
    msg = MIMEText(grade,'plain','utf-8')
    msg['From'] = _format_addr('管理员<%s>' %mailinfo.from_addr)
    msg['To'] = _format_addr(recipient+'<%s>'%mailinfo.to_addr)
    msg['Subject'] = Header(newgrade,'utf-8').encode()
    return msg

#发送邮件
def SendEmail(mailinfo,msg):
    server = smtplib.SMTP_SSL(mailinfo.smtp_server,465)
    server.set_debuglevel(1)
    server.login(mailinfo.from_addr,mailinfo.password)
    server.sendmail(mailinfo.from_addr,[mailinfo.to_addr],msg.as_string())
    server.quit()


#读取成绩
def GetScoreHtml(StuNo,StuId):
    browser = webdriver.Chrome()
    browser.get('http://211.70.149.134:8080/stud_score/brow_stud_score.aspx')
    elem_StuNo = browser.find_element_by_id("TextBox1")
    elem_StuNo.send_keys(StuNo)
    elem_IdNum = browser.find_element_by_id("TextBox2")
    elem_IdNum.send_keys(StuId)
    elem_XN = browser.find_element_by_id("drop_xn")
    elem_XN.send_keys("2017-2018")
    elem_XQ = browser.find_element_by_id("drop_xq")
    elem_XQ.send_keys("1")
    cjcx = browser.find_element_by_id("Button_cjcx")
    cjcx.click()
    return browser

#获取成绩单dict
def PrintScore(browser):
    ScoreHtml = browser.page_source
    ScoreBFSP = BeautifulSoup(ScoreHtml,'lxml')
    ScoreList = ScoreBFSP.find_all(id='GridView_cj')
    ScoreText = BeautifulSoup(str(ScoreList),'lxml')
    grade = {}
    StuInfo = BeautifulSoup(str(ScoreBFSP.find_all(id = 'Label1')),'lxml')
    stustring = StuInfo.text
    grade['stuinfo'] = stustring
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
    return grade

#构造成绩单
def GetGradeString(grade):
    gradestring = grade['stuinfo'] + '\n'
    for k,v in grade.items():
        if k == 'stuinfo':
            continue
        else:
            # print("科目:" +k)
            gradestring = gradestring + "科目：" + k + "\n"
            # print("成绩:"+v)
            gradestring = gradestring + "成绩：" + v + "\n"
    return gradestring

#获取收件人
def GetRecipient(grade):
    str = grade['stuinfo'][13:]
    name = ""
    i = 0
    flag = 0
    while(str[i] != '性'):
        if str[i] == '：':
            flag = 1
            i += 1
            continue
        if flag == 0:
            i += 1
            continue
        elif flag == 1:
            name += str[i]
            i += 1
    return name.strip()

#保存到本地
def SavetoLocal(file_name,contents):
    data = pickle.dumps(obj = contents)
    with open(file_name+'.pickle', mode= 'wb') as fp:
        fp.write(data)

#比较本地成绩单
def Compare(file_name,contents):
    flag = 0
    newsubject = {}
    if os.path.exists(file_name+'.pickle'):
        with open(file_name+'.pickle', mode='rb') as fp:
            data = pickle.loads(fp.read())
            if operator.eq(data,contents):
                return False
            else:
                for key in contents:
                    if key in data:
                        continue
                    else:
                        newsubject[key] = contents[key]
                        SavetoLocal(file_name,contents)
                        flag = 1
    else:
        SavetoLocal(file_name, contents)
        with open(file_name + '.pickle', mode='rb') as fp:
            data = pickle.loads(fp.read())
            for key in contents:
                newsubject[key] = contents[key]
                SavetoLocal(file_name, contents)
                flag = 1
    if flag == 1:
        return newsubject
    else:
        return False

if __name__ == '__main__':
    StuInfo = Student('','')
    browser = GetScoreHtml(StuInfo.StuNo,StuInfo.StuId)
    grade = PrintScore(browser)
    gradestring = GetGradeString(grade)
    recipient = GetRecipient(grade)
    browser.close()
    comflag = Compare(recipient,grade)
    if comflag:
        newsubject = ""
        for key in comflag:
            if key == 'stuinfo':
                continue
            else:
                newsubject += '科目:'+key+' 成绩:'+comflag[key]+'|'
        mailinfo = MailInfo('', '', '', '')
        msg = SetMailHead(recipient,newsubject, gradestring, mailinfo)
        SendEmail(mailinfo, msg)