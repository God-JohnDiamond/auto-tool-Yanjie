'''
Author: John Diamond
Date: 2020-10-19 10:12:38
LastEditors: John Diamond
LastEditTime: 2020-10-22 17:53:32
FilePath: \auto-tools\aotu_tools_2_mouth_report.py
'''
# -*- coding:utf-8 -*-
import openpyxl
from openpyxl.styles import Border, Side
from datetime import date
from datetime import datetime

class cFileOpt:

	def __init__(self):
		self.InputFil = openpyxl.load_workbook('3.xlsx')
		self.NowDate = date.today()						# 获取当前日期
		print('今天是：%s' % self.NowDate)
		self.basedate = date(1899, 12, 30)				# Excel的时间天数是从这天开始计算的
		self.Curdate = self.NowDate - self.basedate			# 计算当前日期的天数
		print('文件 3.xlsx 打开成功！')

	def CloseFiles(self):
		self.InputFil.save(filename = '3.xlsx')
		print('文件 3.xlsx 保存成功！' )
		self.InputFil.close()
		pass

	def Prepare(self):
		self.RawDatSht = self.InputFil.worksheets[0]		# 0 -- raw data sheet
		self.Mouth_Sht = self.InputFil.worksheets[2]		# 2 -- weekly report sheet		3 -- for debug

		m_list = self.Mouth_Sht.merged_cells				# 取消单元格合并
		cr = []												# 先获取已合并的所有单元格
		for m_area in m_list:
			# 合并单元格的起始行坐标、终止行坐标。。。。，
			r1, r2, c1, c2 = m_area.min_row, m_area.max_row, m_area.min_col, m_area.max_col
			# 纵向合并单元格的位置信息提取出
			if r2 - r1 > 0:
				cr.append((r1, r2, c1, c2))
		for r in cr:
				self.Mouth_Sht.unmerge_cells(start_row=r[0], end_row=r[1], start_column=r[2], end_column=r[3])

		for x in range(3, 200):
			self.Mouth_Sht.cell(row = x,column = 33).value = ''
			self.Mouth_Sht.cell(row = x,column = 33).alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
			self.Mouth_Sht.cell(row = x,column = 33).font = openpyxl.styles.Font(name='微软雅黑', size=9)
			pass

		for i in range(1, 10):								# 合并表头单元格
			self.Mouth_Sht.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
		self.Mouth_Sht.merge_cells(start_row=1, start_column=18, end_row=2, end_column=18)
		self.Mouth_Sht.freeze_panes = 'A3'					# 冻结前两行
		pass

	def FillDailySht(self, i, j, celval):
		if celval == 0:
			celval = ''
		self.Mouth_Sht.cell(row = i,column = j).value = celval
		if j == 2:
			self.Mouth_Sht.cell(row = i,column = j).number_format = 'yyyy/mm/dd'
		pass

	def FillMouth_Sht(self):

		pass

	def ParseRawDat(self):
		ProjectSet = set()									# 去重用的项目名称集合
		ProjectList = []									# 去重之后的项目名称
		ProjectNum = 1										# 项目序号
		RowFillDaily = 3									# 略过表头两行
		listmerge = []
		
		for i in range(2, self.RawDatSht.max_row+1):
			ProjectSet.add(self.RawDatSht.cell(row = i,column = 2).value)
			
		ProjectList = list(ProjectSet)						# 项目集合 ——> list
		#print(ProjectList)
		
		for j in range(len(ProjectList)):					# 项目循环
			if ProjectList[j] == None:						# 排除无内容列表
				continue
			
			JobSet = set()									# 去重用的岗位名称集合
			JobList = []									# 去重之后的岗位集合
			
			for projectloop in range(2, self.RawDatSht.max_row+1):	# 遍历以查找
				CelVal = self.RawDatSht.cell(row = projectloop,column = 2).value
				if CelVal == ProjectList[j]:				# 一个项目下
					JobSet.add(self.RawDatSht.cell(row = projectloop,column = 3).value)

			JobList = list(JobSet)
			#print(JobList)
			print('No.%d' % ProjectNum)
			print(ProjectList[j])

			self.FillDailySht(RowFillDaily, 1, ProjectNum)				# 写项目序号
			self.FillDailySht(RowFillDaily, 3, ProjectList[j])			# 写项目名
			
			self.cellborder(RowFillDaily, (RowFillDaily + len(JobList)))# 一个项目画一次框线
			if RowFillDaily != (RowFillDaily + len(JobList) - 1):		# 合并项目的某些单元格
				self.mergecells(RowFillDaily, (RowFillDaily + len(JobList) - 1))
			
			for k in range(len(JobList)):								# 当前项目下的岗位数循环
				Cnt_Recommend = 0										# 推荐数
				Cnt_Effective = 0										# 有效数
				Cnt_Oneside = 0											# 一面数
				Cnt_Endface = 0											# 终面数
				Cnt_Offer = 0											# offer数
				Cnt_Entry = 0											# 入职数

				if JobList[k] == None:									# 排除无内容列表
					continue
				print('%d:%s' % (k + 1, JobList[k]))

				for jobloop in range(2, self.RawDatSht.max_row+1):		# 遍历以查找当前项目下的岗位1， 2，...
					CelVal = self.RawDatSht.cell(row = jobloop,column = 3).value
					if CelVal == JobList[k]:							# 遍历当前项目下的岗位
						Cnt_Recommend += 1								# 当前行有岗位匹配 推荐数+1

						if self.RawDatSht.cell(row = jobloop,column = 10).value == '是':
							Cnt_Effective += 1							# 简历通过 有效数+1
							pass

						if self.RawDatSht.cell(row = jobloop,column = 12).value == '是':
							Cnt_Oneside += 1							# 简历通过 一面数+1
							pass

						if self.RawDatSht.cell(row = jobloop,column = 13).value == '是':
							Cnt_Endface += 1							# 简历通过 终面数+1
							pass

						if self.RawDatSht.cell(row = jobloop,column = 15).value == '是':
							Cnt_Offer += 1								# 简历通过 offer数+1

						if self.RawDatSht.cell(row = jobloop,column = 17).value == '是':
							Cnt_Entry += 1								# 简历通过 入职数+1
							pass
					pass
				# print('推荐:%d' % Cnt_Recommend)
				# print('有效:%d' % Cnt_Effective)
				# print('一面:%d' % Cnt_Oneside)
				# print('终面:%d' % Cnt_Endface)
				# print('offer:%d' % Cnt_Offer)
				# print('入职:%d' % Cnt_Entry)
				self.FillDailySht(RowFillDaily, 2, self.NowDate)		# 写推荐日期
				self.FillDailySht(RowFillDaily, 7, JobList[k])			# 写岗位名
				self.FillDailySht(RowFillDaily, 10, Cnt_Recommend)		# 写各个阶段数量
				self.FillDailySht(RowFillDaily, 11, Cnt_Effective)
				self.FillDailySht(RowFillDaily, 12, Cnt_Oneside)
				self.FillDailySht(RowFillDaily, 13, Cnt_Endface)
				self.FillDailySht(RowFillDaily, 14, Cnt_Offer)
				self.FillDailySht(RowFillDaily, 15, Cnt_Entry)
				
				RowFillDaily += 1										# 每完成一个岗位 行数+1 下移一行
				pass

			print('\n')
			ProjectNum += 1												# 完成一个项目 序号+1
			pass

	def cellborder(self, startrow, endrow):
		thin = openpyxl.styles.Side(style="thin", color="000000")
		for i in range(startrow, endrow):
			for j in range(1, 19):
				self.Mouth_Sht.cell(row = i,column = j).border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
		pass

	def mergecells(self,startrow, endrow):								# 合并某些单元格
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=1, end_row=endrow, end_column=1)
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=3, end_row=endrow, end_column=3)
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=4, end_row=endrow, end_column=4)
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=5, end_row=endrow, end_column=5)
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=6, end_row=endrow, end_column=6)
		self.Mouth_Sht.merge_cells(start_row=startrow, start_column=9, end_row=endrow, end_column=9)
	pass

def main():
	FileOpt = cFileOpt()
	FileOpt.Prepare()
	FileOpt.ParseRawDat()
	FileOpt.CloseFiles()
	
if __name__ == '__main__':
	main()