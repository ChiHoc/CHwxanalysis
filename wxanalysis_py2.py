#!/usr/bin/python
#-*-coding:utf-8-*-

import urllib
import os
import sqlite3
import requests
import json
import xlwt
import math
from Tkinter import *
from datetime import *
from tkMessageBox import showwarning
from pycookiecheat import chrome_cookies
import sys
import platform
import time

reload(sys)
sys.setdefaultencoding('utf-8')

if getattr(sys, 'frozen', False):
    # we are running in a bundle
    OS = platform.system()
    if OS == 'Darwin':
        bundle_dir = os.path.abspath(os.path.dirname(sys.executable) + "/../../../")
    else:
        bundle_dir = os.path.dirname(sys.executable)
else:
    # we are running in a normal Python environment
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

regularStyle = xlwt.easyxf(
    'alignment: horiz centre; font: name Microsoft YaHei')

window = Tk()

# 检查请求次数，返回响应日期


def getRequestCount(begin_date, end_date):

    beginDate = datetime.strptime(begin_date, '%Y-%m-%d')

    endDate = datetime.strptime(end_date, '%Y-%m-%d')

    count = math.ceil((endDate - beginDate).days / 30.0)

    dates = []

    for index in range(0, int(count) - 1):

        beginDate = (beginDate + timedelta(days=30))

        dates.append(beginDate.strftime('%Y-%m-%d'))

    print(u'== 根据时间间隔需要请求' + str(int(count)) + u'次 ==')

    return dates

# 获取含头部的sheet


def getSheet(begin_date, end_date):

    wb = xlwt.Workbook()

    sheet = wb.add_sheet(begin_date + ' - ' + end_date)

    sheet.write(0, 0, u'发布时间', regularStyle)

    sheet.write(0, 1, u'标题', regularStyle)

    sheet.write(0, 2, u'送达人数', regularStyle)

    sheet.write(0, 3, u'图文阅读人数', regularStyle)

    sheet.write(0, 4, u'图文阅读次数', regularStyle)

    sheet.write(0, 5, u'分享人数', regularStyle)

    sheet.write(0, 6, u'分享次数', regularStyle)

    sheet.write(0, 7, u'收藏人数', regularStyle)

    sheet.write(0, 8, u'收藏次数', regularStyle)

    sheet.write(0, 9, u'公众号会话阅读（一次传播）', regularStyle)

    sheet.write(0, 10, u'公众号分享朋友圈（一次传播）', regularStyle)

    sheet.write(0, 11, u'朋友圈分享（二次传播）', regularStyle)

    sheet.write(0, 12, u'朋友圈阅读（二次传播）', regularStyle)

    sheet.write(0, 13, u'公众号会话', regularStyle)

    sheet.write(0, 14, u'好友转发', regularStyle)

    sheet.write(0, 15, u'朋友圈', regularStyle)

    sheet.write(0, 16, u'历史消息', regularStyle)

    sheet.write(0, 17, u'其它', regularStyle)

    return wb, sheet


def saveSheet(wb, name):

    wb.save(bundle_dir + '/' + name + '.xls')


def writeSheet(sheet, start, items):

    print(u'== 开始写入 ==')

    index = start

    for item in items:

        sheet.write(index, 0, item['publish_date'], regularStyle)

        sheet.write(index, 1, item['title'], regularStyle)

        sheet.write(index, 2, item['target_user'], regularStyle)

        sheet.write(index, 3, item['int_page_read_user'], regularStyle)

        sheet.write(index, 4, item['int_page_read_count'], regularStyle)

        sheet.write(index, 5, item['share_user'], regularStyle)

        sheet.write(index, 6, item['share_count'], regularStyle)

        sheet.write(index, 7, item['add_to_fav_user'], regularStyle)

        sheet.write(index, 8, item['add_to_fav_count'], regularStyle)

        sheet.write(index, 9, item['int_page_from_session_read_user'], regularStyle)

        sheet.write(index, 10, item['feed_share_from_session_user'], regularStyle)

        sheet.write(index, 11, item['feed_share_from_feed_user'], regularStyle)

        sheet.write(index, 12, item['int_page_from_feed_read_user'], regularStyle)

        sheet.write(index, 13, '%.2f%%'%(item['int_page_from_session_read_count']/float(item['int_page_read_count'])*100), regularStyle)
        
        sheet.write(index, 14, '%.2f%%'%(item['int_page_from_friends_read_count']/float(item['int_page_read_count'])*100), regularStyle)
        
        sheet.write(index, 15, '%.2f%%'%(item['int_page_from_feed_read_count']/float(item['int_page_read_count'])*100), regularStyle)
        
        sheet.write(index, 16, '%.2f%%'%(item['int_page_from_hist_msg_read_count']/float(item['int_page_read_count'])*100), regularStyle)
        
        sheet.write(index, 17, '%.2f%%'%(item['int_page_from_other_read_count']/float(item['int_page_read_count'])*100), regularStyle)

        index = index + 1

    print(u'== 写入成功 ==')

    return index

# 获取微信数据


def getWechatData(begin_date, end_date, token):

    print(u'== 开始请求 ==')

    url = 'http://mp.weixin.qq.com/misc/appmsganalysis'

    params = {'action': 'all',
              'begin_date': begin_date,
              'end_date': end_date,
              'order_by': '1',
              'order_direction': '1',
              'token': token,
              'lang': 'zh_CN',
              'f': 'json',
              'ajax': '1'}

    headers = {
        'cache-control': 'no-cache',
    }

    cookies = chrome_cookies('http://mp.weixin.qq.com')

    response = requests.request(
        'GET', url, headers=headers, params=params, cookies=cookies)

    respJson = json.loads(response.text)

    if respJson['base_resp']['ret'] == 0:

        items = json.loads(respJson['total_article_data'])['list']

        print(u'== 请求成功，请求量为' + str(len(items)) + u'个 ==')

        return items

    elif respJson['base_resp']['ret'] == -1:

        print(u'== 请使用Chrome到mp.weixin.qq.com进行登录 ==')

        showwarning(title=u'温馨提示', message=u'请使用Chrome到mp.weixin.qq.com进行登录')

        return -1

    elif respJson['base_resp']['ret'] == 200040:

        print(u'== Token错误 ==')

        print(respJson)

        showwarning(title=u'Token错误', message=u'请重新填写Token')

        return -1

    elif respJson['base_resp']['ret'] == 200003:

        print(u'== Session过期，请重新登录 ==')

        print(respJson)

        showwarning(title=u'Session过期', message=u'请重新登录')

        return -1

    else:

        print(u'== 未知错误，请联系陈艾森 ==')

        print response.text

        showwarning(title=u'未知错误', message=u'请联系陈艾森')

        return -1

# 开始请求数据


def runRequestData(begin_date, end_date, token):

    wb, sheet = getSheet(begin_date, end_date)

    index = 1

    requestDates = getRequestCount(begin_date, end_date)

    startTime = begin_date

    for pos in range(0, len(requestDates)):

        date = requestDates[pos]

        print(u'== 第 ' + str(pos + 1) + u' 次 : ' +
              startTime + u' - ' + date + u' ==')

        items = getWechatData(startTime, date, token)

        if items == -1:

            return

        index = writeSheet(sheet, index, items)

        startTime = date

        time.sleep(1)

    saveSheet(wb, begin_date + '-' + end_date)

    print(u'== 执行成功 ==')

    showwarning(title=u'执行成功', message=u'文件保存为: ' + 
        bundle_dir + '/' + begin_date + u'-' + end_date + u'.xls')

# 初始化输入框


def initInputWindow():

    window.title(u'微信统计')

    beginInput = StringVar()
    endInput = StringVar()
    tokenInput = StringVar()

    today=date.today() 
    sevenDays=timedelta(days=7) 
    sevenDaysAgo=today-sevenDays 
    beginInput.set(sevenDaysAgo)
    endInput.set(today)

    Label(window, text=u'开始时间：').grid(row=0)
    Label(window, text=u'结束时间：').grid(row=1)
    Label(window, text=u'Token：').grid(row=2)

    Entry(window, width=20, textvariable=beginInput).grid(row=0, column=1)
    Entry(window, width=20, textvariable=endInput).grid(row=1, column=1)
    Entry(window, width=20, textvariable=tokenInput).grid(row=2, column=1)

    Button(window, text=u'执行', command=lambda: runRequestData(
        beginInput.get(), endInput.get(), tokenInput.get())).grid(row=3, columnspan=2)

    window.mainloop()


initInputWindow()
