# **1.处理异常
# **2.涂色调整
# **3.Log to file
# 编译推广

from shark_modules import build_dasd_list, update_dasdmap, build_dasd_dic, build_dasdmap
import datetime as dt

# 更新已有卷表/创建新卷表
update_exist_map = True

# 指定文件名
dataname = 'DASDINFO.BPZA.D191121'  # 数据源:D U,DASD,ONLINE,0000,9999 & D SMS,SG(ALL),LISTV的输出
xlsname = 'CEBMF-SW--DASD MAP 20191126.xls'  # 现有卷表名,只能是XLS格式
log = open('shark.log','a')

print("\nPROGRAM STARTED AT:", dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),file=log)
print("\nreading dasdinfo..",file=log)
dasd_list = build_dasd_list(source=dataname)  # 调用外部函数,构建一个列表，存放每个devices的信息（卷标和卷组名称）

if dasd_list and update_exist_map:  # 更新已有卷表
	print(f"\nreading dasdinfo complete,total {len(dasd_list)} devices prepared.",file=log)
	print("\nupdating dasdmap..",file=log)
	event = update_dasdmap(source=dasd_list, filenm=xlsname)  # 调用外部函数更新卷表
	if isinstance(event,list):
		for e in event:
			print(e,file=log)
		print("\ndasdmap updated.",file=log)
	else:
		print("\ndasdmap not found.",file=log)

elif dasd_list and (not update_exist_map):  # 创建新卷表
	print(f"\nreading dasdinfo complete,total {len(dasd_list)} devices prepared.",file=log)
	print("\nbuilding new dasdmap..",file=log)
	dasd_dic = build_dasd_dic(source=dasd_list)  # 调用外部函数,构建一个字典，存放每个CU（keys）和CU内的devices(values)
	build_dasdmap(source_dic=dasd_dic)  # 调用外部函数,构建一个XLS格式的DASDMAP:写入每个CU内的devices并构建相关的格式
	print("\nupdating new dasdmap..",file=log)
	event = update_dasdmap(source=dasd_list)  # 调用外部函数更新卷表
	if isinstance(event,list):
		for e in event:
			print(e,file=log)
		print("\nnew dasdmap created.",file=log)
	else:
		print("\ndasdmap not found",file=log)
        
else:
	print("\ndasd data not found",file=log)    

print("\nPROGRAM ENDED AT:", dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),file=log)
