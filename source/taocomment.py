# -*- coding:utf-8 -*-
import urllib.request, urllib.parse, http.cookiejar
import os, time,re
import http.cookies
import json
import xlsxwriter as wx
from PIL import Image
import pymysql
import socket
from bs4 import BeautifulSoup
__author__ = 'hunterhug'
# http://python.jobbole.com/81344/
# 拆分JSON
import xml.dom.minidom
import json
from openpyxl import Workbook
from openpyxl import load_workbook

def password():
	print('请输入你的账号和密码')
	user=input('账号：')
	pwd=input('密码：')
	if user=='jinhan' and pwd=='6833066':
		print('欢迎你：'+user)
		return
	try:
		mysql = pymysql.connect(host="115.28.95.129", user="root", passwd="6833066",db='hunter', charset="utf8")
		cur = mysql.cursor()
		isuser="SELECT * FROM mtaobao where user='{0}' and pwd='{1}'".format(user,pwd)
		cur.execute(isuser)
		mysql.commit()
		if cur.fetchall():
			print('欢迎你：'+user)
			localIP = socket.gethostbyname(socket.gethostname())#这个得到本地ip
			ipList = socket.gethostbyname_ex(socket.gethostname())
			s=''
			for i in ipList:
				if i != localIP and i!=[]:
					s=s+(str)(i)
			timesss=time.strftime('%Y%m%d-%H%M%S', time.localtime())
			update="UPDATE mtaobao SET `times` = `times`+1,`dates`='{0}',`ip` ='{1}' where user='{2}'".format(timesss,s.replace("'",''),user)
			#print(update)
			cur.execute(update)
			mysql.commit()
			cur.close()
			mysql.close()
			return
		else:
			raise
	except Exception as e:
		#print(e)
		mysql.rollback()
		cur.close()
		mysql.close()
		print('密码错误')
		password()

# 找出文件夹下所有html后缀的文件
def listfiles(rootdir, prefix='.xml'):
	file = []
	for parent, dirnames, filenames in os.walk(rootdir):
		if parent == rootdir:
			for filename in filenames:
				if filename.endswith(prefix):
					file.append(rootdir + '/' + filename)
			return file
		else:
			pass

def writeexcel(path,dealcontent):
	workbook = wx.Workbook(path)
	top = workbook.add_format({'border':1,'align':'center','bg_color':'white','font_size':11,'font_name': '微软雅黑'})
	red = workbook.add_format({'font_color':'white','border':1,'align':'center','bg_color':'800000','font_size':11,'font_name': '微软雅黑','bold':True})
	image = workbook.add_format({'border':1,'align':'center','bg_color':'white','font_size':11,'font_name': '微软雅黑'})
	formatt=top
	formatt.set_align('vcenter') #设置单元格垂直对齐
	worksheet = workbook.add_worksheet()        #创建一个工作表对象
	width=len(dealcontent[0])
	worksheet.set_column(0,width,38)            #设定列的宽度为22像素
	for i in range(0,len(dealcontent)):
		if i==0:
			formatt=red
		else:
			formatt=top
		for j in range(0,len(dealcontent[i])):
			if dealcontent[i][j]:
				worksheet.write(i,j,dealcontent[i][j],formatt)
			else:
				 worksheet.write(i,j,'空',formatt)
	workbook.close()
	

def getHtml(url,host='rate.taobao.com',daili='',postdata={}):
	"""
    抓取网页：支持cookie
    第一个参数为网址，第二个为POST的数据

    """
	# COOKIE文件保存路径
	filename = 'cookie.txt'

	# 声明一个MozillaCookieJar对象实例保存在文件中
	cj = http.cookiejar.MozillaCookieJar(filename)
	# cj =http.cookiejar.LWPCookieJar(filename)

	# 从文件中读取cookie内容到变量
	# ignore_discard的意思是即使cookies将被丢弃也将它保存下来
	# ignore_expires的意思是如果在该文件中 cookies已经存在，则覆盖原文件写
	# 如果存在，则读取主要COOKIE
	if os.path.exists(filename):
		cj.load(filename, ignore_discard=True, ignore_expires=True)
	# 读取其他COOKIE
	if os.path.exists('../subcookie.txt'):
		cookie = open('../subcookie.txt', 'r').read()
	else:
		cookie='ddd'
	# 建造带有COOKIE处理器的打开专家
	proxy_support = urllib.request.ProxyHandler({'http':'http://'+daili})
	# 开启代理支持
	if daili:
		print('代理:'+daili+'启动')
		opener = urllib.request.build_opener(proxy_support, urllib.request.HTTPCookieProcessor(cj), urllib.request.HTTPHandler)
	else:
		opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

	# 打开专家加头部
	opener.addheaders = [('User-Agent',
						  'Mozilla/5.0 (iPad; U; CPU OS 4_3_3 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8J2 Safari/6533.18.5'),
						 ('Host', host),
						 ('Cookie',cookie)]

	# 分配专家
	urllib.request.install_opener(opener)
	# 有数据需要POST
	if postdata:
		# 数据URL编码
		postdata = urllib.parse.urlencode(postdata)

		# 抓取网页
		html_bytes = urllib.request.urlopen(url, postdata.encode()).read()
	else:
		html_bytes = urllib.request.urlopen(url).read()

	# 保存COOKIE到文件中
	cj.save(ignore_discard=True, ignore_expires=True)
	return html_bytes

# 去除标题中的非法字符 (Windows)
def validateTitle(title):
	rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/\:*?"<>|'
	new_title = re.sub(rstr, "", title)
	return new_title

# 递归创建文件夹
def createjia(path):
	try:
		os.makedirs(path)
	except:
		print('目录已经存在：'+path)

def timetochina(longtime,formats='{}天{}小时{}分钟{}秒'):
	day=0
	hour=0
	minutue=0
	second=0
	try:
		if longtime>60:
			second=longtime%60
			minutue=longtime//60
		else:
			second=longtime
		if minutue>60:
			hour=minutue//60
			minutue=minutue%60
		if hour>24:
			day=hour//24
			hour=hour%24
		return formats.format(day,hour,minutue,second)
	except:
		raise Exception('时间非法')

def begin():
    sangjin = '''
		-----------------------------------------
		| 欢迎使用自动抓取淘宝商品评论程序   	|
		| 时间：2016年1月7日                   |
		| 新浪微博：一只尼玛                    |
		| 微信/QQ：569929309                    |
		-----------------------------------------
	'''
    print(sangjin)

def taobao(returncomment,max=200):
	"""
	页数限制：200

	"""
	p=0
	while p<max:
		# 评论页数构造
		p=p+1
		page=str(p)

		# 网址构造
		url=urlroot+page

		# 开始抓取
		print('准备抓取'+url)
		tjson=getHtml(url).decode('gbk','ignore').strip().replace('(','').replace(')','')
		# print(tjson)

		# JSON解码
		tjson=json.loads(tjson)

		# 评论
		comments=tjson['comments']

		# 评论为空，跳出
		if comments:
			pass
		else:
			break

		# 逐条评论解析
		for c in comments:

			# 1
			comment=c['content'] # 初次评论内容
			# 2
			date=c['date'] # 评论时间
			# print(comment,date)

			userlist=c["user"] # 用户信息表

			# 3
			user=userlist['nick'] # 用户昵称
			# 4
			usergrade=userlist['displayRatePic'] # 用户等级
			# print(user,usergrade,'\n')

			appendc=c['appendList'] # 追加评论列表

			# 列表不为空
			if appendc:
				# 5
				acomment=appendc[0]['content'] # 追评内容
				# 6
				aday=appendc[0]['dayAfterConfirm'] # 几天后追评
			else:
				acomment=''
				aday=''
			# print('\n',acomment,aday)

			replay=c['reply'] # 商家回复
			# 如果有回复
			if replay:
				# 7
				replay=replay['content']
			else:
				replay=''
			# print(replay)

			returncomment.append([url1,'淘宝',title,p,user,usergrade,comment,date,acomment,aday,replay])
			# print('-'*20)

		print('抓取成功'+url)
	return returncomment

def tmall(returncomment,max=99):
	"""
	页数限制：200

	"""
	p=0
	while p<max:
		# 评论页数构造
		p=p+1
		page=str(p)

		# 网址构造
		url=urlroot+page

		# 开始抓取
		print('准备抓取'+url)
		tjson='{'+getHtml(url).decode('gbk','ignore').strip()+'}'
		# print(tjson)

		# JSON解码
		tjson=json.loads(tjson)
		# print(tjson)

		# 失败证明没有评论
		try:
			# 评论
			tmallc=tjson['rateDetail']['rateList']
		except:
			return returncomment
		# print(tmallc)

		for tc in tmallc:
			# 评论时间
			tdate=tc['rateDate']
			# 初次评论
			tc1=tc['rateContent']
			# 评论用户
			tname=tc['displayUserNick']
			# 用户级别
			tgrade=tc['displayRatePic']

			# 追加评论
			tappendc=tc['appendComment']
			if tappendc:
				# 追评内容
				tc2=tappendc['content']
				# 几天后追评
				tappenddate=tappendc['days']
			else:
				tc2=''
				tappenddate=''

			# 商家回复
			tmallb=tc['reply']
			returncomment.append([url1,'天猫',title,p,tname,tgrade,tc1,tdate,tc2,tappenddate,tmallb])
			# print(p,tname,tgrade,tc1,tdate,tc2,tappenddate,tmallb)
		print('抓取成功'+url)
	return returncomment

if __name__ == '__main__':
	# 欢迎语
	begin()

	# 密码登陆
	password()

	# 抓取时间
	today=time.strftime('%Y%m%d', time.localtime())
	todays=time.strftime('%Y%m%d%H%M%S', time.localtime())

	# 程序测速
	a=time.clock()

	# 温馨提示
	print('请向taobao.txt文件中写入网址链接，每行一条')

	# 打开网址列表所在文件
	file=open('../taobao.txt','r')

	# 分割网址
	websites=file.read().split('\n')

	# 合理网址存放处
	temp=[]

	for i in websites:
		# 忽略#开头的网址
		if '#' in i:
			continue
		# 剔除空网址
		if i:
			temp.append(i)

	# 存储所有评论变量
	returncomment=[['最短网址','类型','商品标题','页码','用户昵称','用户等级','评论','评论时间','追评','几天后追评','商家回复']]
	print('温馨提示：每个商品的评论在一张Excel中。')
	jinhan=input('批量抓取网址评论请按1:')
	jinhan1=input('每个商品默认抓取全部评论请按1:')
	for s in range(0,len(temp)):
		if jinhan=='1':
			ok='2'
			pass
		else:
			ok=input('抓取下一个网址：不抓按数字: 1')
		if ok=='1':
			break
		else:
			# 商品唯一id标号，淘宝三种情况，天猫一种
			if '?id' in temp[s]:
				id=temp[s].split('?id=')[-1].split('&')[0]
			elif '&id' in temp[s]:
				id=temp[s].split('&id=')[-1].split('&')[0]
			elif 'itemId=' in temp[s]:
				id=temp[s].split('itemId=')[-1].split('&')[0]
			else:
				id=temp[s].split('?')[0].split('/')[-1].split('.')[0][1:]

			# 构造最短链接，天猫可通过淘宝跳回去
			url='https://item.taobao.com/item.htm?id='+id
			url1=url

			# 抓取，解码
			content=getHtml(url).decode('gbk','ignore')

			# 解析
			doc=BeautifulSoup(content,'html.parser')

			# 商品标题
			title=doc.find('title').text
			print(url+'  '+title)

			# 寻找评论入口秘钥，主要是userid
			key=doc.find('meta',attrs={'name':'microscope-data'})
			# 存在提取
			if key:
				key=key['content']
			# print(key)
			userid=key.split(';')
			# 查找名为userid的键
			for i in userid:
				if 'userid' in i:
					userid=i
					break
			userid=userid.split('=')[-1]
			# print(userid)

			# 通过标题分析属于哪种类型网址
			if '淘宝' in title:
				# 构造网址，做淘宝标志
				who=1
				urlroot="https://rate.taobao.com/feedRateList.htm?auctionNumId="+id+"&userNumId="+userid+"&showContent=1"+"&currentPageNum="
			elif '天猫' in title:
				who=2
				urlroot="https://rate.tmall.com/list_detail_rate.htm?itemId="+id+"&sellerId="+userid+"&content=1&order=3"+"&currentPage="
			else:
				# 找不到出错
				raise

			# 抓取评论
			# 淘宝链接
			if who==1:
				try:
					if jinhan1=='1':
						raise
					else:
						k=int(input('抓取几页评论（默认全部)：'))
						if k<0:
							print('出错，默认一直抓')
							raise
				except:
					k=200
				returncomment=taobao(returncomment,k)
			else:
				try:
					if jinhan1=='1':
						raise
					else:
						k=int(input('抓取几页评论（默认99)：'))
						if k>99 and k<0:
							print('出错，默认99页')
							raise
				except:
					k=99
				returncomment=tmall(returncomment,k)

			# 写入excel
			# 新建文件夹存放excel
			path='../excel/'+today+'/'+todays
			createjia(path)

			writeexcel(path+'/'+str(s)+'.xlsx',returncomment)

	# 程序测速
	b=time.clock()
	print('运行时间：'+timetochina(b-a))
	input('请关闭窗口')
