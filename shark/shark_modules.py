import xlrd,xlwt
from xlutils.copy import copy

log = []

def build_dasd_list(
                    *,
                    source
                   ):
    """
        Func to build a list containing all device/storgrp info.
        Based on command output of "D U,DASD,ONLINE,0000,9999" &
                                   "D SMS,SG(ALL),LISTVOL" 
    """    
    try:
        with open(source) as fd_obj:
            dasdinfo_lines = fd_obj.readlines()
    except FileNotFoundError:
        return False
    else:
        for i in range(len(dasdinfo_lines)):
            dasdinfo_lines[i] = dasdinfo_lines[i].strip().split()

        dasd_lines = [i for i in dasdinfo_lines 
        if len(i) > 1 and i[1] == '3390']    #抽取ONLINE DEVICE

        sg_lines = [i for i in dasdinfo_lines 
        if len(i) > 2  and i[2] == 'ONRW']   #抽取SMS DEVICE

        sms_count = 0
        for d in dasd_lines:  #合并信息
            if len(d) > 6:    #清洗数据
                d.pop(3)
                log.append('data polished:')
                log.append(d)
            for s in sg_lines:
                if d[0] == s[1]:
                    sms_count += 1
                    d.append(s[4])
        return dasd_lines


def update_dasdmap(
        *,
        source,
        filenm='shark map.xls',
        title_row=3
                    ):
    """
        Funnction used to update dasdmap(device numbers only in xls format)
        Source:online sms or non-sms dasd info (list type) 
    """
    try:
        old_excel = xlrd.open_workbook(filenm, formatting_info = True)
    except FileNotFoundError:
        return False
    else:    
        new_excel = copy(old_excel)
        devnum_str = ''
        
        for device in source:     #遍历devices
            # print('.',end=' ')
            for ws in old_excel.sheets():     #遍历worksheets
                if 'DASD' in ws.name:         
                    cunum = int(ws.row_len(title_row-1) / 3)
                    for w in range(0,cunum*3,3):  #遍历worksheet内的CUs（每个CU块跨三列）
                        for r in range(title_row,len(ws.col(w))):  #遍历CU内的devices
                            if isinstance(ws.cell_value(r,w),float):   #判断device num数据类型
                                devnum_str = str(int(ws.cell_value(r,w)))    
                            else:
                                devnum_str = ws.cell_value(r,w)
                            #如果device number匹配#
                            if devnum_str == device[0]: 
                                if ws.cell_value(r,w+1) != device[3]: #如果需要更新卷标
                                    new_excel.get_sheet(ws.number).write(r,w+1,device[3])  #更新卷标
                                    log.append(f'updating label:{device[0]},{device[3]}')
                                if len(device) > 6 and ws.cell_value(r,w+2) != device[6]: #如果需要更新卷组名
                                    new_excel.get_sheet(ws.number).write(r,w+2,device[6]) #更新卷组名
                                    log.append(f'updating sgname:{device[0]},{device[3]},{device[6]}')
                                elif len(device) == 6 and ws.cell_value(r,w+2) != '':#如果需要清空卷组名
                                    new_excel.get_sheet(ws.number).write(r,w+2,'') #清空卷组名
                                    log.append(f'clearing sgname:{device[0]},{device[3]}')

        new_excel.save(filenm)
        return log

def build_dasd_dic(
                    *,
                    source
                   ):
    """
        Function to return a dictionary,which contain CU(key) and conresponding device(values)
    """
    cu_lines = [source[0][0][0:2]]

    for line in source:
        if line[0][0:2] != cu_lines[-1]:
            cu_lines.append(line[0][0:2])

    d_dic = {}

    for cu in cu_lines:
        d_dic[cu] = []
        for line in source:
            if line[0][0:2] == cu:
                d_dic[cu].append(line)

    return d_dic


def build_dasdmap(
                    *,
                    source_dic,
                    title_row=3
                   ):
    """
        Function to build a xls format dasdmap from a dictionary,which contain CU(key) and conresponding device(values)
    """
    wb_orig = xlwt.Workbook()
    ws_orig = wb_orig.add_sheet('-DASD-')

    style_title = xlwt.easyxf('font: color-index black, bold on;'
                       'pattern: pattern solid, fore_colour gray40;'
                      )
    style_device = xlwt.easyxf('font: color-index black, bold off;'
                        'pattern: pattern solid, fore_colour lime;'
                     )
    c = 0
    for devices in source_dic.values():
        ws_orig.write(title_row-1,c,'CU',style_title)
        ws_orig.write(title_row-1,c+1,'VOLSER',style_title)
        ws_orig.write(title_row-1,c+2,'SGNAME',style_title)
        for r in range(0,len(devices)):
            ws_orig.write(r+title_row,c,devices[r][0],style_device)
            ws_orig.write(r+title_row,c+1,'')
            ws_orig.write(r+title_row,c+2,'')
        c += 3

    wb_orig.save('shark map.xls')
