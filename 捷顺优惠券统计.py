#! /usr/bin/env python3
# -*- coding:utf-8 -*-

import requests
import json, re, copy
from string import digits, ascii_letters
import time, datetime
import tesserocr
from tesserocr import PyTessBaseAPI
from PIL import Image
import xlwt, openpyxl
from io import BytesIO
# 引入配置文件
from conf import *

# 模拟登录
class Login():
	def __init__(self):
		self.headers = {
			'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.117 Safari/537.36',
		}
		# 获取验证码图片的网址
		self.check_url = 'http://merchant.jslife.com.cn/merchant/randomCode?rand=0.9961618814336237'
		# 提交用户名、密码、验证码的网址
		self.post_url = 'http://merchant.jslife.com.cn/merchant/login'
		# 获取优惠卷领用次数的网址
		self.gift_url = 'http://merchant.jslife.com.cn/merchant/couponReceive/queryCouponReceiveTotal'
		# 获取优惠卷使用次数的网址
		self.used_url = 'http://merchant.jslife.com.cn/merchant/couponUsed/queryCouponUsedTotal'
		# 维持会话
		self.session = requests.session()
	# 模拟登录
	def login(self, userName, password):
		response = {'respMsg': ''}
		while response['respMsg'] != '成功':
			post_data = {
				'userCode': userName,
				'pwd': password,
				'rand': 0.9961618814336237,
				'validateCode': self.check()
			}
			response = self.session.post(self.post_url, data=post_data, headers=self.headers)
			if response.status_code == 200:
				response = response.json()
		return True
	# 识别4位由英文数字组成的验证码
	def check(self):
		result = ''
		# 排除掉识别错误的验证码
		while len(result) != 4:
			response = self.session.get(self.check_url, headers=self.headers)
			# f = open('check.jpg', 'wb')
			# f.write(response.content)
			# f.close()
			# image = Image.open('check.jpg')
			f = BytesIO()
			f.write(response.content)
			image = Image.open(f)
			image = image.convert('L')
			threshold = 125
			table = []
			for i in range(256):
				if i < threshold:
					table.append(0)
				else:
					table.append(1)
			image = image.point(table, '1')
			# image.show()
			# 设置识别规则
			with PyTessBaseAPI() as api:
				api.SetVariable('tessedit_char_whitelist', digits + ascii_letters)
				api.SetImage(image)
				result = api.GetUTF8Text().replace(' ', '').replace('\n', '')
		return result
	# 获取日期、优惠卷领用次数、使用次数的字典列表
	def get_dicts(self, startTime='2018-05-01'):
		# 获取统计起始日期
		try:
			wb = openpyxl.load_workbook('优惠券统计报表.xlsx')
			ws = wb["Sheet1"]
			a = 2
			while ws.cell(row=a, column=2).value != None:
				a += 1
			startTime = str(datetime.date.fromordinal(ws.cell(row=a, column=1).value+693594)) + ' 00:00:00'
		except IOError:
			print('找不到要更新的Excel,统计从2018-01-01开始')
			startTime += ' 00:00:00'
		finally:
			datas = []
			# 将格式化日期转为时间戳
			a = time.mktime(time.strptime(startTime, '%Y-%m-%d %H:%M:%S'))
			while time.time() - a > 24*60*60:
				data = {}
				data['date'] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(a)).replace(' 00:00:00', '')
				gift_params = {
					'userId': '328a0c0583cb4bed82bf6a95348d4cdb',
					'couponReceivePage': json.dumps({
					    "planName": "",
					    "useStatus": "",
					    "couponNumber": "",
					    "channel": "",
					    "cheapType": "",
					    "useTimeBegin": "",
					    "useTimeEnd": "",
					    "receiveTimeBegin": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(a)),
					    "receiveTimeEnd": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(a+24*60*60-1))
					})
				}
				response = self.session.post(url = self.gift_url, headers = self.headers, data = gift_params).json()
				data['getTimes'] = response['respData']['totalHoursTime']
				used_params = {
					'userId': '328a0c0583cb4bed82bf6a95348d4cdb',
					'couponUsePage': json.dumps({
					    "planName": "",
					    "parkName": "",
					    "cheapType": "",
					    "status": "",
					    "storeName": "",
					    "useTimeBegin": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(a)),
					    "useTimeEnd": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(a+24*60*60-1))
					})
				}
				response = self.session.post(url = self.used_url, headers = self.headers, data = used_params).json()
				data['usedTimes'] = response['respData']['totalHoursTime']
				datas.append(data)
				a += 24*60*60
			return(datas)
	# 将数据导出为excel文档
	def new_excel(self, dicts):
		# 新建excel
		wb = xlwt.Workbook()
		# 添加工作薄
		sh = wb.add_sheet('Sheet1')
		style = xlwt.XFStyle()
		style.num_format_str = 'M月D日'
		# 写入数据
		for i in range(len(dicts)):
			sh.write(i, 0,datetime.datetime.strptime(dicts[i]['date'], '%Y-%m-%d').toordinal()-693594, style)
			sh.write(i, 1, dicts[i]['getTimes'])
			sh.write(i, 2, dicts[i]['usedTimes'])
		# 保存文件
		wb.save('Jcount.xls')
	def change_excel(self, dicts):
		try:
			# input('请确定要更新的Excel文件名为：优惠券统计报表.xlsx，并且在同一文件夹下。\n按回车键继续！')
			wb = openpyxl.load_workbook('优惠券统计报表.xlsx')
			ws = wb["Sheet1"]
			a = 2
			while ws.cell(row=a, column=2).value != None:
				a += 1
			dictsCopy = dicts.copy()
			for item in dictsCopy:
				if datetime.datetime.strptime(item['date'], '%Y-%m-%d').toordinal()-693594 == ws.cell(row=a, column=1).value:
					break
				dicts.remove(item)
			# 拷贝单元格边框与对齐样式
			border = copy.copy(ws.cell(row=2, column=2).border)
			alignment = copy.copy(ws.cell(row=2, column=2).alignment)
			for item in dicts:
				b = ws.cell(row=a, column=2)
				c = ws.cell(row=a, column=3)
				b.value = item['getTimes']
				b.border = border
				b.alignment = alignment
				c.value = item['usedTimes']
				c.border = border
				c.alignment = alignment
				a += 1
			wb.save('优惠券统计报表.xlsx')
			print('更新Excel成功！')
		except IOError:
			print('找不到要更新的Excel.将新建一个Excel')
			self.new_excel(dicts)
			
login = Login()
# 输入登录用户名与密码
login.login(USER, PWD)
# 获取日期与优惠卷领用数量，使用数量
dicts = login.get_dicts()
login.change_excel(dicts)
