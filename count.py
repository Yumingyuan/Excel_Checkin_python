import sys
reload(sys)
sys.setdefaultencoding('utf8')
import xlrd
import xlwt
from xlutils.copy import copy
if __name__=="__main__":
	workbook1 = xlrd.open_workbook('list.xls')#open dianmingcezhengban 
	workbook2=xlrd.open_workbook('Checkin.xls')# open qiandaobiao
	sheet_names1= workbook1.sheet_names()
	sheet_names2=workbook2.sheet_names()
	sheet1=sheet_names1[0]
	sheet2=sheet_names2[0]
	sheet1 = workbook1.sheet_by_name(sheet1)
	sheet2=workbook2.sheet_by_name(sheet2)
	all_student=[]
	students_come=[]
	student_nocomes=[]
	#print sheet1.nrows
	#print sheet2.nrows
	cols1 = sheet1.col_values(1)#获得第二列
	cols2 = sheet2.col_values(1)
	
	#make_all_student_queue
	for i in range(sheet1.nrows):
		if sheet1.row_values(i)[1].isdigit():
			all_student.append(sheet1.row_values(i)[1])
			#print sheet1.row_values(i)[1]
	#print len(all_student)
	#make_studentcome_queue
	for i in range(sheet2.nrows):
		#print sheet2.row_values(i)[0]#获得学生的学号
		if sheet2.row_values(i)[0].isdigit():
			students_come.append(sheet2.row_values(i)[0])
			#print sheet2.row_values(i)[0]
		#if int(sheet2.row_values(i)[1]).isdigit():
			#print str(sheet2.row_values(i)[1])
			#for j in range(sheet1.nrows):
			#	if sheet1.row_values(j)[1]==sheet2.row_values(i)[1]:
			#		print sheet1.row_values(j)
			#		print j
			#	else:
			#		pass
			#		print sheet1.row_values(j)[1]
	student_nocomes=list(set(all_student).difference(set(students_come)))#取差集，点名册里有而签到单里没有
	#for student_nocome in student_nocomes:#打印测试
		#print student_nocome
	new_excel=copy(workbook1)
	ws=new_excel.get_sheet(0)
	for i in range(len(student_nocomes)):
		tmp_stu=student_nocomes.pop()
		for j in range(sheet1.nrows):
			if sheet1.row_values(j)[1]==tmp_stu:
				ws.write(j,18,'0')#write at column 12
			else:
				pass
	#for student in students_come:
	#	print student
	#print len(students_come)
	new_excel.save('result.xls')