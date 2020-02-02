import os
import xlrd,xlwt
from xlutils.copy import copy

def file_name_walk(file_dir):
	for root, dirs, files in os.walk(file_dir):
		filenames = [f for f in files if 'SUMALL' in f and len(f)==14]
	return filenames
		
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

xlsname = 'SGINFO.XLS'     #specify excel file to be updated.
sites = 'BPZA','SPZB'
style = xlwt.easyxf('pattern: pattern solid, fore_colour turquoise;')

data_dir = os.path.split(os.path.realpath(__file__))[0]

try:
	old_excel = xlrd.open_workbook(xlsname, formatting_info = True)
except FileNotFoundError:
	print('excel file not located in current location.')
else:    
        new_excel = copy(old_excel)
        for site in sites:
                filenames = file_name_walk(data_dir + '\\' + site)
                if filenames:
                        filenames.sort()
                        all_data = {}

                        for filename in filenames:
                            all_data[filename[7:14]] = get_sg_peak(filenm=(data_dir+'\\'+site+'\\'+filename))

                        ws = old_excel.sheet_by_name(site)
                        c = 2   #first day peak data start @colume 2(count from zero)
                        for k,v in all_data.items():
                                new_excel.get_sheet(site).write(0,c,k)  #date
                                for i in range(1,len(ws.col(0))):
                                        sgname = ws.cell_value(i,0)
                                        new_excel.get_sheet(site).write(i,1,v[sgname][-2])  #capacity
                                        if v[sgname][-1] < 75:
                                                new_excel.get_sheet(site).write(i,c,v[sgname][-1]) #peak
                                        else:
                                                new_excel.get_sheet(site).write(i,c,v[sgname][-1],style)
                                print(site+":"+k)
                                c = c + 1
                else: print('NO SUMALL FOUND for ' + site)
        new_excel.save(xlsname)
