#!/usr/bin/env python
#-*- coding: utf-8 -*-

# python实现全角半角的相互转换
# http://www.cnblogs.com/kaituorensheng/p/3554571.html
# 全角半角转换说明
# 有规律（不含空格）：
#   全角字符unicode编码从65281~65374 （十六进制 0xFF01 ~ 0xFF5E）
#   半角字符unicode编码从33~126 （十六进制 0x21~ 0x7E）
# 特例：
#   空格比较特殊，全角为 12288（0x3000），半角为 32（0x20）
#
# 除空格外，全角/半角按unicode编码排序在顺序上是对应的（半角=全角-0xfee0）,所以可以直接通过用+-法来处理非空格数据，对空格单独处理。
# 注：
# 1. 中文文字永远是全角，只有英文字母、数字键、符号键才有全角半角的概念,一个字母或数字占一个汉字的位置叫全角，占半个汉字的位置叫半角。
# 2. 引号在中英文、全半角情况下是不同的

def strQ2B(ustring):
	"""全角转半角"""
	rstring = ""
	for uchar in ustring:
		inner_code = ord(uchar)
		if inner_code == 0x3000:  # 全角空格直接转换
			inner_code = 0x0020
		elif inner_code >= 0xFF01 and inner_code <= 0xFF5E:  # 全角字符（除空格）根据关系转化
			inner_code -= 0xFEE0

		rstring += chr(inner_code)
	return rstring

# # 用正则表达式删除字符串中的无用括号
# import re
# s = u'123(45)a啊速度（伤害）有限公司'
# ss = re.sub(u'[()（）]', '', s)
# print(ss)

test_str = 'ＭＯ－　ｍ０－　ｔＯ＿　Ｔ０－１２３４５６７８９０'
print(strQ2B(test_str).upper())			# 'MO- M0- TO_ T0-1234567890'
