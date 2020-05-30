# -*- coding:utf-8 -*-
import subprocess
import threading
import time
import datetime
import socket, sys
import os
import xlrd
import xlsxwriter
import re
from copy import deepcopy
import pandas as pd
import tkinter
from tkinter import ttk


def getIcebergBugList(info):
    port = 51444
    host = "172.20.223.50"
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect((host, port))
        print('建立{} connecting...'.format(info))
    except Exception as e:
        print("Error connection to server: %s" % e)
        sys.exit(1)
    try:
        s.sendall(("getIcebergBugList {}".format(info)).encode())
        print("获取{}数据".format(info))
        f = open('{}.xls'.format(info), 'wb')
        while True:
            data = s.recv(4096)
            if data == b"SSEND":
                break
            elif data == b"EOF":
                break
            elif data == b"ERROR":
                return False
            file_size = data.decode()
            print(file_size)
        while True:
            data = s.recv(4096)
            if not len(data):
                break
            elif data == b"EOF":
                break
            elif data == b"ERROR":
                return False
            f.write(data)
        f.close()
        if str((os.stat('{}.xls'.format(info))).st_size) == file_size:
            print("取得{}数据".format(info))
            return True
        else:
            print("保存{}文件错误file error".format(info))
            return False
    except Exception as e:
        print(e)
        return False


# 打开一个excel文件
def open_xls(file):
    fh = xlrd.open_workbook(file)
    return fh


# 获取excel中所有的sheet表
def getsheet(fh):
    return fh.sheets()


# 获取sheet表的行数
def getnrows(fh, sheet):
    table = fh.sheets()[sheet]
    return table.nrows


# # 读取文件内容并返回行内容
# def getFilect(file, shnum,datavalue):
#     fh = open_xls(file)
#     table = fh.sheets()[shnum]
#     num = table.nrows
#     for row in range(1,num,1):
#         rdata = table.row_values(row)
#         if rdata not in datavalue:
#             datavalue.append(rdata)
#     return datavalue

# 获取sheet表的个数
def getshnum(fh):
    x = 0
    sh = getsheet(fh)
    for sheet in sh:
        x += 1
    return x


def countdata(alldata):
    alldata.remove(['issuekey', 'status', 'memo', 'title'])
    alldict = []  # 按周统计BUG状态
    countdate = []
    for item in range(len(alldata)):
        date = re.findall(r'B*(\d+)-', alldata[item][0])[0]
        this_date = datetime.datetime.strptime(date, "%y%m%d")
        year, week, day = this_date.isocalendar()
        alldict.append({'year': year, 'week': week, 'value': alldata[item]})
        this_date_str = this_date.strftime("%Y-%m")
        alldata[item].append(this_date_str)
        if this_date_str not in countdate:
            countdate.append(this_date_str)

    bugcountdate = dict([(k, []) for k in countdate])

    df = pd.DataFrame(data=alldata, columns=['issuekey', 'status', 'memo', 'title', 'date'])
    df['date'] = pd.to_datetime(df['date'])
    df = df.set_index('date')
    # cache = df.to_dict(orient='date')
    # df['2013-11']
    for i in countdate:
        cache = df[i].reset_index().to_dict(orient='issuekey')
        bugcountdate[i] = cache
    tablelist = sorted(alldict, key=lambda i: (i['year'], i['week']))
    bugcount = {}  # {年：{周：[buglist]}}
    for i in tablelist:
        if i['year'] not in bugcount.keys():
            bugcount.update({i['year']: {i['week']: [i['value']]}})
        else:
            if i['week'] not in bugcount[i['year']].keys():
                bugcount[i['year']].update({i['week']: [i['value']]})
            else:
                bugcount[i['year']][i['week']].append(i['value'])

    datetable = []
    allcount = 0
    # 理论上：
    # ISSUE_NOTDO # 不作处理
    # ISSUE_CLOSED  # 关闭
    # ISSUE_RESOLVED  # 解决待关闭
    # ISSUE_OPEN 新的
    # ISSUE_INPROGRESS  # 正在处理
    # ISSUE_REOPENED  # 重新打开
    # ISSUE_DELAYDO  # 延后处理
    # ISSUE_RETURN_REWRITE  # 退回


    allbugcount = {'ISSUE_CLOSED': 0, 'ISSUE_DELAYDO': 0, 'ISSUE_INPROGRESS': 0, 'ISSUE_NOTDO': 0, 'ISSUE_OPEN': 0,
                   'ISSUE_REOPENED': 0, 'ISSUE_RESOLVED': 0, 'ISSUE_RETURN_REWRITE': 0}
    # bugcount = {}  # {年：{周：[buglist]}}
    for n in sorted(bugcount):
        for m in sorted(bugcount[n]):
            weekcount = 0
            baseFunctionBugCount = 0
            for j in bugcount[n][m]:
                if '基本功能' in j[3]:
                    baseFunctionBugCount = baseFunctionBugCount + 1
                if j[1] == 'ISSUE_CLOSED':
                    allbugcount['ISSUE_CLOSED'] = allbugcount['ISSUE_CLOSED'] + 1
                elif j[1] == 'ISSUE_DELAYDO':
                    allbugcount['ISSUE_DELAYDO'] = allbugcount['ISSUE_DELAYDO'] + 1
                elif j[1] == 'ISSUE_INPROGRESS':
                    allbugcount['ISSUE_INPROGRESS'] = allbugcount['ISSUE_INPROGRESS'] + 1
                elif j[1] == 'ISSUE_NOTDO':
                    allbugcount['ISSUE_NOTDO'] = allbugcount['ISSUE_NOTDO'] + 1
                elif j[1] == 'ISSUE_OPEN':
                    allbugcount['ISSUE_OPEN'] = allbugcount['ISSUE_OPEN'] + 1
                elif j[1] == 'ISSUE_REOPENED':
                    allbugcount['ISSUE_REOPENED'] = allbugcount['ISSUE_REOPENED'] + 1
                elif j[1] == 'ISSUE_RESOLVED':
                    allbugcount['ISSUE_RESOLVED'] = allbugcount['ISSUE_RESOLVED'] + 1
                elif j[1] == 'ISSUE_RETURN_REWRITE':
                    allbugcount['ISSUE_RETURN_REWRITE'] = allbugcount['ISSUE_RETURN_REWRITE'] + 1
                allcount = allcount + 1
                weekcount = weekcount + 1
            atime = time.strptime('{} {} 1'.format(n, m), '%Y %W %w')
            d = datetime.datetime.fromtimestamp(time.mktime(atime)).strftime('%Y-%m-%d')
            datetable.append([d, m, len(bugcount[n][m]), allcount, baseFunctionBugCount])

    # datetable 是sheet1上的大打标格
    # allbugcount 是sheet1上第二个小表格
    # bugcountdate 是sheet3
    return datetable, allbugcount, bugcountdate


def bugstatus(data):
    alist = ['base', 'moudle', 'ROMDEVTEST', 'cream', 'maoyan', 'app', 'other']
    acount = {'ISSUE_CLOSED': 0, 'ISSUE_DELAYDO': 0, 'ISSUE_INPROGRESS': 0, 'ISSUE_NOTDO': 0, 'ISSUE_OPEN': 0,
              'ISSUE_REOPENED': 0, 'ISSUE_RESOLVED': 0, 'ISSUE_RETURN_REWRITE': 0, 'BASELINE_TEST': 0}
    abugcountdate = dict([(k, deepcopy(acount)) for k in alist])
    for i in range(len(data.keys())):
        issuekey = data[i]['issuekey']
        status = data[i]['status']
        memo = data[i]['memo']
        title = data[i]['title']
        if "ROMDEVTEST" in title and u"[应用测试部]" in title:
            abugcountdate['ROMDEVTEST'][status] += 1
        elif '基本功能' in title:
            abugcountdate['base'][status] += 1
        elif '模块稳定性' in title:
            abugcountdate['moudle'][status] += 1
        elif '相机稳定性' in title:
            abugcountdate['cream'][status] += 1
        elif '冒烟测试' in title:
            abugcountdate['maoyan'][status] += 1
        elif '应用稳定性(Iceberg)' in title:
            abugcountdate['app'][status] += 1
        else:
            abugcountdate['other'][status] += 1
    return abugcountdate


def merge_Excel():
    # 定义要合并的excel文件列表
    allxls = [os.getcwd() + '\\backup.xlsx', os.getcwd() + '\\iceberg.xls', os.getcwd() + '\\RomDevTest.xls']
    # 存储所有读取的结果
    alldata = [['issuekey', 'status', 'memo', 'title']]
    for fl in allxls:
        if os.path.exists(fl):
            fh = open_xls(fl)
            x = getshnum(fh)
            for shnum in range(x):
                print("正在读取文件：" + str(fl) + "的第" + str(shnum) + "个sheet表的内容...")
                # alldata = getFilect(fl, shnum,alldata)
                table = fh.sheets()[shnum]
                num = table.nrows
                for row in range(1, num, 1):
                    rdata = table.row_values(row)
                    if rdata not in alldata:
                        alldata.append(rdata)
                    # else:
                    #     print('去重数据')
                    #     print(rdata)
        else:
            if 'backup.xlsx' in fl:
                pass
            else:
                print(fl)
                input('本地文件不存在请检查,按任意键退出')
                exit()
    # 定义最终合并后生成的新文件
    now_date = datetime.datetime.now().strftime('%Y-%m-%d')
    endfile = os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date)
    backfile = os.getcwd() + '\\backup.xlsx'.format(now_date)
    wbback = xlsxwriter.Workbook(backfile)
    wsback = wbback.add_worksheet('旧BUG数据缓存')

    wb1 = xlsxwriter.Workbook(endfile)

    # 样式1
    # style1 = wb1.add_format({'bold': True, 'border': 1, 'border_color': 'black'})
    style1 = wb1.add_format({'border': 1, 'border_color': 'black'})
    style1.set_align('center')  # 水平对齐
    style1.set_align('vcenter')  # 垂直对齐
    style1.set_text_wrap()  # 内容换行
    # 样式2
    style2 = wb1.add_format({'bold': True, 'border': 1, 'border_color': 'black', 'bg_color': '#46A3FF'})
    style2.set_align('center')  # 水平对齐
    style2.set_align('vcenter')  # 垂直对齐
    style2.set_text_wrap()  # 内容换行

    # 创建一个sheet工作对象
    wsdata = wb1.add_worksheet('数据图表')
    wsbug = wb1.add_worksheet('bug状态')
    ws = wb1.add_worksheet('bug总数')

    ws.set_column('A:A', 14)
    ws.set_column('B:B', 16)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 40)
    ws.set_column('E:E', 12)
    for a in range(len(alldata)):
        for b in range(len(alldata[a])):
            c = alldata[a][b]
            ws.write(a, b, c)
            if b == 0 and c.lower() != 'issuekey':
                #用re模块解析出bug对应的年月日
                try:
                    this_date = datetime.datetime.strptime(re.findall(r'B*(\d+)-', c)[0], "%y%m%d")
                    year, week, day = this_date.isocalendar()
                    date = this_date.strftime("%Y-%m-%d")
                    ws.write(a, b + 4, date)
                    ws.write(a, b + 5, week)
                except Exception as e:
                    print(c, e)
            wsback.write(a, b, c)
    wbback.close()
    print("文件合并完成\n正在处理数据...")

    wsdata.set_column('A:A', 11)
    wsdata.set_column('B:B', 5)
    wsdata.set_column('C:C', 6.5)
    wsdata.set_column('D:D', 6.5)
    wsdata.set_column('E:E', 13)
    wsdata.set_column('G:G', 11)
    row0 = ['日期', '周数', 'BUG数', '总数', '基本功能bug']
    col6 = ['不作处理', '关闭', '解决待关闭', '新的', '正在处理', '重新打开', '延后处理', '退回']
    # ISSUE_NOTDO # 不作处理
    # ISSUE_CLOSED  # 关闭
    # ISSUE_RESOLVED  # 解决待关闭
    # ISSUE_OPEN 新的
    # ISSUE_INPROGRESS  # 正在处理
    # ISSUE_REOPENED  # 重新打开
    # ISSUE_DELAYDO  # 延后处理
    # ISSUE_RETURN_REWRITE  # 退回
    bugtitle = ['日期', '模块名称', '关闭', '延后处理', '正在处理', '不作处理', '新的', '重新打开', '解决待关闭', '退回', '基线测试']
    moudlename = ['基本功能', '模块稳定性', 'ROMDEVTEST', '相机稳定性', '冒烟测试', '应用稳定性测试', '其他', '合计']
    for col in range(len(row0)):
        wsdata.write(0, col, row0[col], style2)
    wsdata.write(0, 6, '当前状态', style2)
    wsdata.write(0, 7, '总数', style2)
    for col in range(len(bugtitle)):
        wsbug.write(0, col, bugtitle[col], style2)

    datetable, bugcount, bugcountdate = countdata(alldata)

    bugcountdatekeys = list(bugcountdate.keys())
    alist = ['base', 'moudle', 'ROMDEVTEST', 'cream', 'maoyan', 'app', 'other']
    acount = ['ISSUE_CLOSED', 'ISSUE_DELAYDO', 'ISSUE_INPROGRESS', 'ISSUE_NOTDO', 'ISSUE_OPEN',
              'ISSUE_REOPENED', 'ISSUE_RESOLVED', 'ISSUE_RETURN_REWRITE', 'BASELINE_TEST']
    wsbug.set_column('B:B', 14)
    wsbug.set_column('E:E', 11)
    for i in range(len(bugcountdatekeys)):
        abugcountdate = bugstatus(bugcountdate[bugcountdatekeys[i]])
        # 合并单元格，
        wsbug.merge_range('A{}:A{}'.format(str((i * 8) + 2), str((i * 8) + 9)), bugcountdatekeys[i], style1)

        for j in range(len(moudlename)):
            wsbug.write((i * 8) + j + 1, 1, moudlename[j], style1)
        wsbug.write((i * 8) + 8, 2, r'=SUM(C{}:C{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 3, r'=SUM(D{}:D{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 4, r'=SUM(E{}:E{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 5, r'=SUM(F{}:F{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 6, r'=SUM(G{}:G{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 7, r'=SUM(H{}:H{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 8, r'=SUM(I{}:I{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 9, r'=SUM(J{}:J{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 10, r'=SUM(K{}:K{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        for x in range(len(alist)):
            for y in range(len(acount)):
                # 感觉这里有问题
                wsbug.write((i * 8) + x + 1, 2 + y, int(abugcountdate[alist[x]][acount[y]]), style1)
    # wsbug.merge_range('A{}:B{}'.format(str(8*len(bugcountdatekeys)+2), str(8*len(bugcountdatekeys)+2)), '合计', style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 2, r'=SUMPRODUCT((MOD(ROW(C2:C{}),8)=1)*C2:C{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 3, r'=SUMPRODUCT((MOD(ROW(D2:D{}),8)=1)*D2:D{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 4, r'=SUMPRODUCT((MOD(ROW(E2:E{}),8)=1)*E2:E{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 5, r'=SUMPRODUCT((MOD(ROW(F2:F{}),8)=1)*F2:F{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 6, r'=SUMPRODUCT((MOD(ROW(G2:G{}),8)=1)*G2:G{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 7, r'=SUMPRODUCT((MOD(ROW(H2:H{}),8)=1)*H2:H{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 8, r'=SUMPRODUCT((MOD(ROW(I2:I{}),8)=1)*I2:I{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 9, r'=SUMPRODUCT((MOD(ROW(J2:J{}),8)=1)*J2:J{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    for row in range(len(col6)):
        wsdata.write(row + 1, 6, col6[row], style1)
    wsdata.write(1, 7, bugcount['ISSUE_NOTDO'], style1)  # 不作处理
    wsdata.write(2, 7, bugcount['ISSUE_CLOSED'], style1)  # 关闭
    wsdata.write(3, 7, bugcount['ISSUE_RESOLVED'], style1)  # 解决待关闭
    wsdata.write(4, 7, bugcount['ISSUE_OPEN'], style1)  # 新的
    wsdata.write(5, 7, bugcount['ISSUE_INPROGRESS'], style1)  # 正在处理
    wsdata.write(6, 7, bugcount['ISSUE_REOPENED'], style1)  # 重新打开
    wsdata.write(7, 7, bugcount['ISSUE_DELAYDO'], style1)  # 延后处理
    wsdata.write(8, 7, bugcount['ISSUE_RETURN_REWRITE'], style1)  # 退回
    for a in range(len(datetable)):
        for b in range(len(datetable[a])):
            c = datetable[a][b]
            wsdata.write(a + 1, b, c, style1)
    # --------2、生成图表并插入到excel---------------


    # 创建一个折线图(line chart)
    chart_col = wb1.add_chart({'type': 'line'})
    # 配置第一个系列数据
    if len(datetable) > 20:
        getcol = len(datetable) + 1 - 20
    else:
        getcol = 0
    chart_col.add_series({
        # 这里的sheet1是默认的值，因为我们在新建sheet时没有指定sheet名
        # 如果我们新建sheet时设置了sheet名，这里就要设置成相应的值
        'name': '汇总',
        'categories': '=数据图表!$B${}:$B${}'.format(getcol, len(datetable) + 1),
        'values': '=数据图表!$C${}:$C${}'.format(getcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'marker': {'type': 'diamond'},
        'data_labels': {'value': True},
    })
    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '每周BUG数统计'})
    # chart_col.set_x_axis({'name': 'BUG个数'})
    # chart_col.set_y_axis({'name': '周数'})
    # 设置图表的风格
    # chart_col.set_style(1)
    chart_col.set_size({'width': 550, 'height': 350})
    # 把图表插入到worksheet并设置偏移
    # wsdata.insert_chart('K1', chart_col, {'x_offset': 25, 'y_offset': 10})
    wsdata.insert_chart('K1', chart_col)

    if len(datetable) > 10:
        bugcol = len(datetable) + 1 - 10
    else:
        bugcol = 0



    # 创建一个柱形图(column chart)
    chart_column = wb1.add_chart({'type': 'column'})
    # 配置第一个系列数据
    chart_column.add_series({
        'name': 'BUG趋势',
        'categories': '=数据图表!$B${}:$B${}'.format(bugcol, len(datetable) + 1),
        'values': '=数据图表!$D${}:$D${}'.format(bugcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'data_labels': {'value': True},
    })
    # 设置图表的大小
    chart_column.set_size({'width': 550, 'height': 350})
    # 把图表插入到worksheet并设置偏移
    wsdata.insert_chart('K19', chart_column)




    # 创建一个扇形图(column chart)
    chart_pie = wb1.add_chart({'type': 'pie'})
    # 配置第一个系列数据
    chart_pie.add_series({
        'name': 'BUG状态统计',
        'categories': '=数据图表!$G$2:$G$9',
        'values': '=数据图表!$H$2:$H$8',
        # 'points': [
        #     {'fill': {'color': '#00CD00'}},
        #     {'fill': {'color': 'red'}},
        #     {'fill': {'color': 'yellow'}},
        #     {'fill': {'color': 'gray'}},
        # ],
        'data_labels': {'value': True, 'percentage': True, 'leader_lines': True, 'legend_key': True, 'category': True},
    })
    # 设置图表的大小
    chart_pie.set_size({'height': 500})
    # 把图表插入到worksheet并设置偏移
    wsdata.insert_chart('K37', chart_pie)

    # 创建一个折线图(line chart)
    chart_line = wb1.add_chart({'type': 'line'})
    # 配置第一个系列数据
    chart_line.add_series({
        'name': '汇总',
        'categories': '=数据图表!$B${}:$B${}'.format(getcol, len(datetable) + 1),
        'values': '=数据图表!$E${}:$E${}'.format(getcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'marker': {'type': 'diamond'},
        'data_labels': {'value': True},
    })
    chart_line.set_title({'name': '基本功能测试BUG统计'})
    chart_line.set_size({'width': 550, 'height': 350})
    wsdata.insert_chart('K63', chart_line)
    wsbug.freeze_panes(1, 1)  # # Freeze the first row and column
    wsdata.freeze_panes(1, 1)

    wb1.close()
    print("图表已生成完毕")
    input('数据已保存到:' + os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date) + '\n请按任意键退出并打开表格')
    # os.system('start "{}"'.format(os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date)))
    # 用subprocess打开excel
    filepath = os.path.join(os.getcwd(), 'BugCounter-{}.xlsx'.format(now_date))
    cmd = '"{}"'.format(filepath)
    cmddoing = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    Result = cmddoing.stdout.read()







































def regular():
    # string = ['iceberg', 'ROMDEVTEST']
    # for xlsfile in os.listdir(os.getcwd()):
    #     if '.xls' in xlsfile:
    #         if xlsfile != 'backup.xlsx':
    #             os.remove(os.path.join(os.getcwd(), xlsfile))
    # for info in string:
    #     if not getIcebergBugList(info):
    #         input('从服务器获取{}数据失败,请联系管理员,按任意键退出'.format(info))
    #         exit()
    # time.sleep(5)
    merge_Excel()




# this_date = datetime.datetime.strptime(re.findall(r'B*(\d+)-', c)[0], "%y%m%d")
# year, week, day = this_date.isocalendar()
def customBug(startyear, startmonth, endyear, endmonth,sbtn,win):
    sbtn.config(state=tkinter.DISABLED)
    thread_num = threading.active_count()
    print(thread_num)
    if thread_num<=5:
        if startyear=='':
            startyear='2016'
        if startmonth=='':
            startmonth='1'
        if endyear=='':
            endyear='2016'
        if endmonth=='':
            endmonth='1'
        print('正在在查询，请稍后....')
        startime = str(startyear) + str(startmonth)
        endtime = str(endyear) + str(endmonth)
        backupath = os.getcwd() + '\\backup.xlsx'
        backexcel = xlrd.open_workbook(backupath)
        table = backexcel.sheets()[0]
        finalrowdata = table.row_values(-1)
        this_date = datetime.datetime.strptime(re.findall(r'B*(\d+)-', finalrowdata[0])[0], "%y%m%d")
        year, week, day = this_date.isocalendar()
        date = this_date.strftime("%Y-%m-%d")  # 2020-05-20
        finalyear = year
        finalmonth = date.split('-')[1]

        # 询问用户是否原因触发更新
        if int(endyear) >= int(finalyear) and int(endmonth) >= int(finalmonth):
            pass
        #     print('执行更新合并筛选操作')

        # 不触发更新
        elif int(endyear) <= int(finalyear):
            untriggerupdate(startyear, startmonth, endyear, endmonth)

        else:
            print('无效输入，请检查日期是否输入正确！')
    else:
        print('请勿重负点击！')
    win.destroy()



def untriggerupdate(startyear, startmonth, endyear, endmonth):
    startime = str(int(startyear)*10000+int(startmonth)*100)
    endtime = str(int(endyear)*10000+int(endmonth)*100)
    backupath = os.getcwd() + '\\backup.xlsx'
    backexcel = xlrd.open_workbook(backupath)
    table = backexcel.sheets()[0]
    finalrowdata = table.row_values(-1)
    this_date = datetime.datetime.strptime(re.findall(r'B*(\d+)-', finalrowdata[0])[0], "%y%m%d")
    year, week, day = this_date.isocalendar()
    date = this_date.strftime("%Y-%m-%d")  # 2020-05-20
    finalyear = year
    finalmonth = date.split('-')[1]
    print('执行筛选操作,删除旧表，生成新表中...')
    for xlsfile in os.listdir(os.getcwd()):
        if '.xls' in xlsfile:
            if 'BugCounter' in xlsfile:
                os.remove(os.path.join(os.getcwd(), xlsfile))
    tempalldata = [['issuekey', 'status', 'memo', 'title']]
    for i in range(1, table.nrows):
        rowdata = table.row_values(i)
        bugdate = datetime.datetime.strptime(re.findall(r'B*(\d+)-', rowdata[0])[0], "%y%m%d")
        bugyear, week, day = bugdate.isocalendar()
        bugmonth = bugdate.month
        realbugdate = int(bugyear)*10000+int(bugmonth)*100
        # 这一步筛选满足日期条件的行信息
        if realbugdate <= int(endtime) and realbugdate >= int(startime):
            tempalldata.append(rowdata)

    now_date = datetime.datetime.now().strftime('%Y-%m-%d')
    endfile = os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date)

    tempbackfile = os.getcwd() + '\\tempbackup.xlsx'.format(now_date)
    wbback = xlsxwriter.Workbook(tempbackfile)
    wsback = wbback.add_worksheet('临时BUG数据缓存')
    wb1 = xlsxwriter.Workbook(endfile)
    # 样式1
    # style1 = wb1.add_format({'bold': True, 'border': 1, 'border_color': 'black'})
    style1 = wb1.add_format({'border': 1, 'border_color': 'black'})
    style1.set_align('center')  # 水平对齐
    style1.set_align('vcenter')  # 垂直对齐
    style1.set_text_wrap()  # 内容换行
    # 样式2
    style2 = wb1.add_format({'bold': True, 'border': 1, 'border_color': 'black', 'bg_color': '#46A3FF'})
    style2.set_align('center')  # 水平对齐
    style2.set_align('vcenter')  # 垂直对齐
    style2.set_text_wrap()  # 内容换行

    # 创建一个sheet工作对象
    wsdata = wb1.add_worksheet('数据图表')
    wsbug = wb1.add_worksheet('bug状态')
    ws = wb1.add_worksheet('bug总数')

    ws.set_column('A:A', 14)
    ws.set_column('B:B', 16)
    ws.set_column('C:C', 20)
    ws.set_column('D:D', 40)
    ws.set_column('E:E', 12)
    for a in range(len(tempalldata)):
        for b in range(len(tempalldata[a])):
            c = tempalldata[a][b]
            ws.write(a, b, c)
            if b == 0 and c.lower() != 'issuekey':
                try:
                    this_date = datetime.datetime.strptime(re.findall(r'B*(\d+)-', c)[0], "%y%m%d")
                    year, week, day = this_date.isocalendar()
                    date = this_date.strftime("%Y-%m-%d")
                    ws.write(a, b + 4, date)
                    ws.write(a, b + 5, week)
                except Exception as e:
                    print(c, e)
            wsback.write(a, b, c)
    wbback.close()
    print("文件合并完成\n正在处理数据...")

    wsdata.set_column('A:A', 11)
    wsdata.set_column('B:B', 5)
    wsdata.set_column('C:C', 6.5)
    wsdata.set_column('D:D', 6.5)
    wsdata.set_column('E:E', 13)
    wsdata.set_column('G:G', 11)
    row0 = ['日期', '周数', 'BUG数', '总数', '基本功能bug']
    col6 = ['不作处理', '关闭', '解决待关闭', '新的', '正在处理', '重新打开', '延后处理', '退回']
    # ISSUE_NOTDO # 不作处理
    # ISSUE_CLOSED  # 关闭
    # ISSUE_RESOLVED  # 解决待关闭
    # ISSUE_OPEN 新的
    # ISSUE_INPROGRESS  # 正在处理
    # ISSUE_REOPENED  # 重新打开
    # ISSUE_DELAYDO  # 延后处理
    # ISSUE_RETURN_REWRITE  # 退回
    bugtitle = ['日期', '模块名称', '关闭', '延后处理', '正在处理', '不作处理', '新的', '重新打开', '解决待关闭', '退回', '基线测试']
    moudlename = ['基本功能', '模块稳定性', 'ROMDEVTEST', '相机稳定性', '冒烟测试', '应用稳定性测试', '其他', '合计']
    for col in range(len(row0)):
        wsdata.write(0, col, row0[col], style2)
    wsdata.write(0, 6, '当前状态', style2)
    wsdata.write(0, 7, '总数', style2)
    for col in range(len(bugtitle)):
        wsbug.write(0, col, bugtitle[col], style2)

    datetable, bugcount, bugcountdate = countdata(tempalldata)

    bugcountdatekeys = list(bugcountdate.keys())
    alist = ['base', 'moudle', 'ROMDEVTEST', 'cream', 'maoyan', 'app', 'other']
    acount = ['ISSUE_CLOSED', 'ISSUE_DELAYDO', 'ISSUE_INPROGRESS', 'ISSUE_NOTDO', 'ISSUE_OPEN',
              'ISSUE_REOPENED', 'ISSUE_RESOLVED', 'ISSUE_RETURN_REWRITE', 'BASELINE_TEST']
    wsbug.set_column('B:B', 14)
    wsbug.set_column('E:E', 11)
    for i in range(len(bugcountdatekeys)):
        abugcountdate = bugstatus(bugcountdate[bugcountdatekeys[i]])
        # 合并单元格，
        wsbug.merge_range('A{}:A{}'.format(str((i * 8) + 2), str((i * 8) + 9)), bugcountdatekeys[i], style1)

        for j in range(len(moudlename)):
            wsbug.write((i * 8) + j + 1, 1, moudlename[j], style1)
        wsbug.write((i * 8) + 8, 2, r'=SUM(C{}:C{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 3, r'=SUM(D{}:D{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 4, r'=SUM(E{}:E{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 5, r'=SUM(F{}:F{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 6, r'=SUM(G{}:G{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 7, r'=SUM(H{}:H{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 8, r'=SUM(I{}:I{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 9, r'=SUM(J{}:J{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        wsbug.write((i * 8) + 8, 10, r'=SUM(K{}:K{})'.format(str((i * 8) + 2), str((i * 8) + 8)), style1)
        for x in range(len(alist)):
            for y in range(len(acount)):
                # 感觉这里有问题
                wsbug.write((i * 8) + x + 1, 2 + y, int(abugcountdate[alist[x]][acount[y]]), style1)
    # wsbug.merge_range('A{}:B{}'.format(str(8*len(bugcountdatekeys)+2), str(8*len(bugcountdatekeys)+2)), '合计', style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 2, r'=SUMPRODUCT((MOD(ROW(C2:C{}),8)=1)*C2:C{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 3, r'=SUMPRODUCT((MOD(ROW(D2:D{}),8)=1)*D2:D{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 4, r'=SUMPRODUCT((MOD(ROW(E2:E{}),8)=1)*E2:E{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 5, r'=SUMPRODUCT((MOD(ROW(F2:F{}),8)=1)*F2:F{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 6, r'=SUMPRODUCT((MOD(ROW(G2:G{}),8)=1)*G2:G{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 7, r'=SUMPRODUCT((MOD(ROW(H2:H{}),8)=1)*H2:H{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 8, r'=SUMPRODUCT((MOD(ROW(I2:I{}),8)=1)*I2:I{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)
    # wsbug.write(8*len(bugcountdatekeys)+1, 9, r'=SUMPRODUCT((MOD(ROW(J2:J{}),8)=1)*J2:J{})'.format(str(8*len(bugcountdatekeys)), str(8*len(bugcountdatekeys))), style1)

    for row in range(len(col6)):
        wsdata.write(row + 1, 6, col6[row], style1)
    wsdata.write(1, 7, bugcount['ISSUE_NOTDO'], style1)  # 不作处理
    wsdata.write(2, 7, bugcount['ISSUE_CLOSED'], style1)  # 关闭
    wsdata.write(3, 7, bugcount['ISSUE_RESOLVED'], style1)  # 解决待关闭
    wsdata.write(4, 7, bugcount['ISSUE_OPEN'], style1)  # 新的
    wsdata.write(5, 7, bugcount['ISSUE_INPROGRESS'], style1)  # 正在处理
    wsdata.write(6, 7, bugcount['ISSUE_REOPENED'], style1)  # 重新打开
    wsdata.write(7, 7, bugcount['ISSUE_DELAYDO'], style1)  # 延后处理
    wsdata.write(8, 7, bugcount['ISSUE_RETURN_REWRITE'], style1)  # 退回
    # 这一层循环，写入tempbackup 和BugCounter的所有满足条件的bug
    for a in range(len(datetable)):
        for b in range(len(datetable[a])):
            c = datetable[a][b]
            wsdata.write(a + 1, b, c, style1)
    # --------2、生成图表并插入到excel---------------

    # 创建一个折线图(line chart)
    chart_col = wb1.add_chart({'type': 'line'})
    # 配置第一个系列数据
    datetablelen = len(datetable)
    if len(datetable) > 20:
        getcol = len(datetable) + 1 - 20
    else:
        getcol = 1
    chart_col.add_series({
        # 这里的sheet1是默认的值，因为我们在新建sheet时没有指定sheet名
        # 如果我们新建sheet时设置了sheet名，这里就要设置成相应的值
        'name': '汇总',
        'categories': '=数据图表!$B${}:$B${}'.format(getcol, len(datetable) + 1),
        'values': '=数据图表!$C${}:$C${}'.format(getcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'marker': {'type': 'diamond'},
        'data_labels': {'value': True},
    })
    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '每周BUG数统计'})
    # chart_col.set_x_axis({'name': 'BUG个数'})
    # chart_col.set_y_axis({'name': '周数'})
    # 设置图表的风格
    # chart_col.set_style(1)
    chart_col.set_size({'width': 550, 'height': 350})
    # 把图表插入到worksheet并设置偏移
    # wsdata.insert_chart('K1', chart_col, {'x_offset': 25, 'y_offset': 10})
    wsdata.insert_chart('K1', chart_col)

    if len(datetable) > 10:
        bugcol = len(datetable) + 1 - 10
    else:
        bugcol = 0
    # 创建一个柱形图(column chart)
    chart_column = wb1.add_chart({'type': 'column'})
    # 配置第一个系列数据
    chart_column.add_series({
        'name': 'BUG趋势',
        'categories': '=数据图表!$B${}:$B${}'.format(bugcol, len(datetable) + 1),
        'values': '=数据图表!$D${}:$D${}'.format(bugcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'data_labels': {'value': True},
    })
    # 设置图表的大小
    chart_column.set_size({'width': 550, 'height': 350})
    # 把图表插入到worksheet并设置偏移
    wsdata.insert_chart('K19', chart_column)

    # 创建一个扇形图(column chart)
    chart_pie = wb1.add_chart({'type': 'pie'})
    # 配置第一个系列数据
    chart_pie.add_series({
        'name': 'BUG状态统计',
        'categories': '=数据图表!$G$2:$G$9',
        'values': '=数据图表!$H$2:$H$8',
        # 'points': [
        #     {'fill': {'color': '#00CD00'}},
        #     {'fill': {'color': 'red'}},
        #     {'fill': {'color': 'yellow'}},
        #     {'fill': {'color': 'gray'}},
        # ],
        'data_labels': {'value': True, 'percentage': True, 'leader_lines': True, 'legend_key': True,
                        'category': True},
    })
    # 设置图表的大小
    chart_pie.set_size({'height': 500})
    # 把图表插入到worksheet并设置偏移
    wsdata.insert_chart('K37', chart_pie)

    # 创建一个折线图(line chart)
    chart_line = wb1.add_chart({'type': 'line'})
    # 配置第一个系列数据
    chart_line.add_series({
        'name': '汇总',
        'categories': '=数据图表!$B${}:$B${}'.format(getcol, len(datetable) + 1),
        'values': '=数据图表!$E${}:$E${}'.format(getcol, len(datetable) + 1),
        # 'line': {'color': 'blue'},
        'marker': {'type': 'diamond'},
        'data_labels': {'value': True},
    })
    chart_line.set_title({'name': '基本功能测试BUG统计'})
    chart_line.set_size({'width': 550, 'height': 350})
    wsdata.insert_chart('K63', chart_line)
    wsbug.freeze_panes(1, 1)  # # Freeze the first row and column
    wsdata.freeze_panes(1, 1)

    wb1.close()
    print("图表已生成完毕")
    input('数据已保存到:' + os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date) + '\n请按任意键退出并打开表格:')
    # os.system('start "{}"'.format(os.getcwd() + '\\BugCounter-{}.xlsx'.format(now_date)))
    # 用subprocess打开excel
    filepath = os.path.join(os.getcwd(), 'BugCounter-{}.xlsx'.format(now_date))
    cmd = '"{}"'.format(filepath)
    cmddoing = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


#自定义的线程函数类
def thread_it(func, *args):
  '''将函数放入线程中执行'''
  # 创建线程
  t = threading.Thread(target=func, args=args)
  # 守护线程
  t.setDaemon(True)
  # 启动线程
  t.start()

sy=''
sm=''
ey=''
em=''
def guimode():
    win = tkinter.Tk()
    win.wm_attributes('-topmost', 1)
    win.title("Bug 日期选择")  # #窗口标题
    win.geometry("450x200+500+300")

    def bindsy(*args):
        global sy
        sy=str(staryearcom.get()).split(':')[1]
    startyearxVariable = tkinter.StringVar()  # #创建变量，便于取值
    staryearcom = ttk.Combobox(win, textvariable=startyearxVariable,state="readonly")  # #创建下拉菜单
    staryearcom.grid(row=1,column=1,padx=30,pady=30)
    staryearcom["value"] = (['起始年:'+str(y) for y in range(2016,int(datetime.datetime.now().year)+1)])
    staryearcom.current(0)
    staryearcom.bind("<<ComboboxSelected>>", bindsy)

    def bindsm(*args):
        global sm
        sm = str(starmonthcom.get()).split(':')[1]
    startmonthxVariable = tkinter.StringVar()
    starmonthcom = ttk.Combobox(win,textvariable=startmonthxVariable,state="readonly")  # #创建下拉菜单
    starmonthcom.grid(row=2,column=1)
    starmonthcom["value"] = (['起始月:'+str(y) for y in range(1, 13)])
    starmonthcom.current(0)
    starmonthcom.bind("<<ComboboxSelected>>", bindsm)

    def bindey(*args):
        global ey
        ey = str(endyearcom.get()).split(':')[1]
    endyearxVariable = tkinter.StringVar()  # #创建变量，便于取值
    endyearcom = ttk.Combobox(win, textvariable=endyearxVariable,state="readonly")  # #创建下拉菜单
    endyearcom.grid(row=1,column=2)
    endyearcom["value"] = (['结束年:'+str(y) for y in range(2016,int(datetime.datetime.now().year)+1)])
    endyearcom.current(0)
    endyearcom.bind("<<ComboboxSelected>>",bindey)


    def bindem(*args):
        global em
        em=str(endmonthcom.get()).split(':')[1]
    endmonthxVariable = tkinter.StringVar()
    endmonthcom = ttk.Combobox(win, textvariable=endmonthxVariable,state="readonly")  # #创建下拉菜单
    endmonthcom.grid(row=2,column=2)
    endmonthcom["value"] = (['结束月:'+str(y) for y in range(1, 13)])
    endmonthcom.current(0)
    endmonthcom.bind("<<ComboboxSelected>>",bindem)

#执行耗时操作新开一个线程
    sbtn=ttk.Button(win,text="开始查询",command=lambda :thread_it(customBug, sy,sm,ey,em,sbtn,win))
#这个是单线程模式 ，会报anr
    # sbtn = ttk.Button(win, text="开始查询", command=lambda: customBug(sy, sm, ey, em))
    sbtn.grid(row=3,column=1,padx=0,pady=20,columnspan=2)



    win.mainloop()



if __name__ == "__main__":
    while True:
        print('-**************-')
        print('1.导出所有BUG')
        print('2.导出自定义日期BUG')
        try:
            select = int(input('请选择：'))
        except Exception:
            print('无效输入')
        if select == 1:
            # 调用张尧的老方法
            regular()
            break
        if select == 2:
            guimode()
            break
            # startyear = int(input('输入开始年：').strip())
            #
            # startmonth = int(input('输入开始月：').strip())
            # endyear = int(input('输入结束年：').strip())
            # endmonth = int(input('输入结束月:').strip())
            # customBug(startyear, startmonth, endyear, endmonth)

        else:
            print('无效输入')




