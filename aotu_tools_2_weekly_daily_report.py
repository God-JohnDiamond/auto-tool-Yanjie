'''
Author: John Diamond
Date: 2020-10-19 10:12:38
LastEditors: John Diamond
LastEditTime: 2020-10-19 17:29:57
FilePath: \auto_tools\aotu_tools_2_weekly_daily_report.py
'''
# -*- coding:utf-8 -*-
import openpyxl

class cFileOpt:
	OutCnt = 0
	def __init__(self):
		self.InputFil = openpyxl.load_workbook('1.xlsx')
		print('文件 1.xlsx 打开成功！')

	def CloseFiles(self):
		self.InputFil.save(filename = '1.xlsx')
		print('文件 1.xlsx 保存成功！' )
		self.InputFil.close()
		pass

	def Prepare(self):
		self.RawDatSht = self.InputFil.worksheets[0]		# 0 -- raw data sheet
		self.Daily_Sht = self.InputFil.worksheets[3]		# 1 -- daily report sheet		3 -- for debug
		self.WeeklySht = self.InputFil.worksheets[2]		# 2 -- weekly report sheet
		pass

	def ParseRawDat(self):
		ProjectSet = set()		# 去重用的项目名称集合
		ProjectList = []		# 去重之后的项目名称
		JobList = []
		#ProjectList_Full = []	# 所有的项目名称

		for i in range(2, self.RawDatSht.max_row+1):
			ProjectSet.add(self.RawDatSht.cell(row = i,column = 2).value)
			#ProjectList_Full.append(self.RawDatSht.cell(row = i,column = 2).value)
		ProjectList = list(ProjectSet)	# 项目集合 ——> list

		for j in range(len(ProjectList)):
			JobSet = set()		# 去重用的岗位名称集合
			RowJobInProject = []
			for i in range(2, self.RawDatSht.max_row+1):	# 遍历查找某项目 并解析项目详细
				CelVal = self.RawDatSht.cell(row = i,column = 2).value
				if CelVal == ProjectList[j]:
					JobSet.add(self.RawDatSht.cell(row = i,column = 3).value)
					JobList = list(JobSet)
					RowJobInProject += 1

					for cnt_job in (1, CntJobInProject):

						pass
		pass

	def FillDailySht(self):

		pass

	def FillWeeklySht(self):

		pass

def main():
	FileOpt = cFileOpt()
	FileOpt.Prepare()
	FileOpt.ParseRawDat()
	FileOpt.FillDailySht()
	FileOpt.FillWeeklySht()
	FileOpt.CloseFiles()
	
if __name__ == '__main__':
	main()