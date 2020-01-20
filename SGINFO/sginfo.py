import xlrd,xlwt
from xlutils.copy import copy

def get_sg_peak(*,
                filenm
                 ):
	
	sg_dic = {}
	
	with open(filenm) as f:
		sginfo_lines = f.readlines()
	
	for i in range(len(sginfo_lines)):
		sginfo_lines[i] = sginfo_lines[i].strip().split(';')
		
	sgnames = list({line[2].strip() for line in sginfo_lines})   
	
	for sgname in sgnames:
		sg_dic[sgname] = [line for line in sginfo_lines 
		                  if line[2].strip()==sgname]
		
	for k in sg_dic.keys():
		capacity = max([ i[3].strip() for i in sg_dic[k] ])
		sg_dic[k].append(capacity)
		peak = (100 - 
		       float(min([ i[4].strip() for i in sg_dic[k][:-1] 
		       if i[3].strip() == sg_dic[k][-1]]))
		       )
		sg_dic[k].append(peak)
		
	return sg_dic

filename = 'SUMALL.D190524'
xlsname = 'SGINFO.XLS'
style = xlwt.easyxf('pattern: pattern solid, fore_colour turquoise;')

dic1 = get_sg_peak(filenm=filename)

try:
	old_excel = xlrd.open_workbook(xlsname, formatting_info = True)
except FileNotFoundError:
	print('未找到excel文件')
else:    
	new_excel = copy(old_excel)
	ws = old_excel.sheet_by_index(0)
	for i in range(1,len(ws.col(0))):
		sgname = ws.cell_value(i,0)
		new_excel.get_sheet(0).write(i,1,dic1[sgname][-2])
		if dic1[sgname][-1] < 75:
			new_excel.get_sheet(0).write(i,2,dic1[sgname][-1])
		else:
 			new_excel.get_sheet(0).write(i,2,dic1[sgname][-1],style)

new_excel.save(xlsname)

# for k,v in dic1.items():
	# print(k,"{},{:5.1f}".format(v[-2],v[-1]))
