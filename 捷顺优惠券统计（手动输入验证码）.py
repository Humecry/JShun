#! /usr/bin/env python3
# -*- coding:utf-8 -*-

import requests
import json
import time
from PIL import Image
import re
import xlwt
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
				print(response)
		print('登录成功！')
		return True
	# 识别4位由英文数字组成的验证码
	def check(self):
		result = ''
		# 排除掉识别错误的验证码
		while not re.match('[a-z0-9]{4}', result) or len(result) != 4:
			response = self.session.get(self.check_url, headers=self.headers)
			f = open('check.jpg', 'wb')
			f.write(response.content)
			f.close()
			image = Image.open('check.jpg')
			image = image.convert('L')
			threshold = 125
			table = []
			for i in range(256):
				if i < threshold:
					table.append(0)
				else:
					table.append(1)
			image = image.point(table, '1')
			image.show()
			result = input('请输入验证码:')
		return result
	# 获取日期、优惠卷领用次数、使用次数的字典列表
	def get_dict(self):
		datas = []
		# 格式化开始日期
		startTime = '2018-09-01 00:00:00'
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
			print(gift_params)
			response = self.session.post(url = self.gift_url, headers = self.headers, data = gift_params).json()
			print(response)
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
		print('获取数据成功！')
		return(datas)
	# 将数据导出为excel文档
	def wt_excel(self, dict):
		# 新建excel
		wb = xlwt.Workbook()
		# 添加工作薄
		sh = wb.add_sheet('Sheet1')
		# 写入数据
		for i in range(len(dict)):
			sh.write(i, 0, dict[i]['date'])
			sh.write(i, 1, dict[i]['getTimes'])
			sh.write(i, 2, dict[i]['usedTimes'])
		print('导出Excel成功！')
		# 保存文件
		wb.save('myexcel.xls')

login = Login()
# 模拟登录
login.login(USER, PWD)
# 获取日期与优惠卷领用数量，使用数量
dict = login.get_dict()
# 将数据导出为excel文档
login.wt_excel(dict)