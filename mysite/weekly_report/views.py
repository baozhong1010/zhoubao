# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.shortcuts import render
from django.http import HttpResponse
from django.http import StreamingHttpResponse

import requests
from bs4 import BeautifulSoup

from itertools import product
import types

import openpyxl
from openpyxl import worksheet
from openpyxl import load_workbook
from openpyxl import Workbook

import time
import datetime
from datetime import timedelta

import re
# Create your views here.

def date_time():
    time_today = datetime.date.today()
    time1 = timedelta(days=7)
    time2 = timedelta(days=2)
    time_ago = time_today - time1
    time_now = time_today - time2
    _times = '%s' ' ---- ' '%s' % (time_ago, time_now)
    dic = {"_times":_times,"time_ago":time_ago,"time_now":time_now}
    return dic


def get_data():

#登录部分
    root_url = 'http://172.16.203.12/zentao/user-login.html'

    my_url = 'http://172.16.203.12/zentao/my-project.html'
    r = requests.Session()
    UA = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36"

    header = {"User-Agent": UA,
               "referer":"http://172.16.203.12/zentao/my/"
               }
    f = r.get(root_url,headers = header)

#cookie登录
    r.cookies = requests.utils.cookiejar_from_dict({
        
        'zentaosid':'641kllhds6ae5o9fcor51u9ts2'})
    r.post(root_url,
        cookies = r.cookies,      
        headers = header
        )
#账号密码登录
    '''
    postdata = {
        'account':'baozhong',
        'password':'111111'
    }
    r.post(
        root_url,
        data = postdata,
        headers = header)
    '''
    
#抓取迭代名称
    h = r.get(my_url,headers = header)
    _soup = BeautifulSoup(h.content,'lxml')
    diedais = _soup.select("table tbody tr")
    task_names = []             #task_names是迭代名称
    ID_names = []               #ID_names是迭代ID
    planA1_names = []
    planA2_names = []
    out_plan_names = []               
    for diedai in diedais:
        lt = list(diedai.stripped_strings)
        diedai_status = lt[5]
        if lt[5] == '进行中':
            task_names.append(lt[2])
            ID_names.append(lt[0])
    for n in task_names:
        m = n[:5]                                           #取迭代名称的前5个字符串   
        print m
        x = re.match(r'^(\d{1}).(\d{1}).(\d{1})$', m)       #判断前5个字符串是不是数字（1.2.3类型的）
        if x:
            ID = m
            if x.group(1) == '1':       #如果第一个.之后是1，就是planA1计划的任务
                
                planA1_names.append(n[6:])
                print planA1_names
                
            elif x.group(1) == '2':       #如果第一个.之后是2，就是planA2计划的任务
                
                planA2_names.append(n[6:])
                print planA2_names    
        else:
            out_plan_names.append(n)
            print out_plan_names
    


#抓取任务名称
    for ID_name in ID_names:           #ID_name是迭代的ID
        print ID_name
        index_url = 'http://172.16.203.12/zentao/project-task-' + ID_name + '.html'
        f = r.get(index_url,headers = header)
        soup = BeautifulSoup(f.content,'lxml')
        plans = soup.find_all('tr',class_='text-center')
        names = []
        times = []
        jindus = []
        task_ID_names = []          #task_ID_names是任务名称的ID
    for plan in plans:          
        l = list(plan.stripped_strings)
        status = l[2]
        task_ID_names.append(l[0])
    for task_ID_name in task_ID_names:

        task_url = 'http://172.16.203.12/zentao/task-view-' + task_ID_name + '.html'     #这是某具体任务名称的网页地址
        #进入某具体任务名称网页抓取该任务名称的历史记录时间
        t = r.get(task_url,headers = header)
        t_soup = BeautifulSoup(t.content,'lxml')
        time_logs = t_soup.find_all('span', class_ = "item")
    for time_log in time_logs:
        lis = list(time_log.stripped_strings)
        c = lis[0]
        _lis = c[:10]                            
        a = time.strptime(_lis,"%Y-%m-%d")
        b = datetime.date(*a[:3])                  #转换成时间格式的历史记录时间
    b_times = date_time()
    time_ago = b_times['time_ago']
    time_now = b_times['time_now']
    if b >= time_ago and b <= time_now:
        print b
        names.append(l[1])
        times.append(l[3])
        jindus.append(l[-1])
        print l[1],status,l[3],l[-1],'\n'

    

    ret = {
    'diedai_status':diedai_status,
    'planA1_names':planA1_names,
    'planA2_names':planA2_names,
    'out_plan_names':out_plan_names,
    'names':names,
    'ID':ID,
    'status':status, 'jindus':jindus
    }
    return ret






def index(request):
    data = get_data()
    times = date_time()
    time = times['_times']
    names = data['names']
    jindus = data['jindus']
    status = data['status']
    ID = data['ID']
    planA1_names = data['planA1_names']
    planA2_names = data['planA2_names']
    out_plan_names = data['out_plan_names']
    diedai_status = data['diedai_status']
    context = {
    'diedai_status':diedai_status,
    'time':time,
    'status':status,
    'planA1_names':planA1_names,
    'planA2_names':planA2_names,
    'out_plan_names':out_plan_names,
    'names':names,
    'jindus':jindus,
    'ID':ID
        }

    return render(request,'weekly_report/index.html',context)



def downloadFile(request):
    data = get_data()
    times = date_time()
    time = times['_times']
    names = data['names']
    jindus = data['jindus']
    status = data['status']
    ID = data['ID']
    planA1_names = data['planA1_names']
    planA2_names = data['planA2_names']
    out_plan_names = data['out_plan_names']
    diedai_status = data['diedai_status']
    wb = load_workbook("zhoubao.xlsx")
    ws = wb.active
    num1 = 21
    num2 = 21
    num3 = 9
    num4 = 12
    for m in planA1_names:
        num3 = num3 + 1
        cell3 = 'B' + str(num3)
        ws[cell3] = m
        cell5 = 'D' + str(num3)
        ws[cell5] = ID
        cell6 = 'E' + str(num3)
        ws[cell6] = diedai_status
    for n in planA2_names:
        num4 = num4 + 1
        cell4 = 'B' + str(num4)
        ws[cell4] = n
        num5 = num5 + 1
        cell5 = 'D' + str(num4)
        ws[cell5] = ID
    for i in out_plan_names:
        num1 = num1 +1
        cell1 = 'B'+ str(num1)
        ws[cell1] = i
    for j in jindus:
        num2 = num2 + 1
        cell2 = 'E' + str(num2)
        ws[cell2] = status + j
    ws['G7'] = time
    wb.save('zhoubao.xlsx')
    file_name = 'zhoubao.xlsx'
    def file_iterator(file_name, chunk_size=512):#用于形成二进制数据  
        with open(file_name,'rb') as f:  
            while True:  
                c = f.read(chunk_size)
                if c:  
                    yield c 
                else:
                    break  
    the_file_name ="zhoubao.xlsx"#要下载的文件路径  
    response =StreamingHttpResponse(file_iterator(the_file_name))#这里创建返回  
    response['Content-Type'] = 'application/vnd.ms-excel'#注意格式   
    response['Content-Disposition'] = 'attachment;filename="zhoubao.xlsx"'#注意filename 这个是下载后的名字  
    return response


#这是一个解决openpyxl打开表格部分边框缺失的补丁
def patch_worksheet():
    """This monkeypatches Worksheet.merge_cells to remove cell deletion bug
    https://bitbucket.org/openpyxl/openpyxl/issues/365/styling-merged-cells-isnt-working
    Thank you to Sergey Pikhovkin for the fix
    """

    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        """ Set merge on a cell range.  Range is a cell range (e.g. A1:E1)
        This is monkeypatched to remove cell deletion bug
        https://bitbucket.org/openpyxl/openpyxl/issues/365/styling-merged-cells-isnt-working
        """
        if not range_string and not all((start_row, start_column, end_row, end_column)):
            msg = "You have to provide a value either for 'coordinate' or for\
            'start_row', 'start_column', 'end_row' *and* 'end_column'"
            raise ValueError(msg)
        elif not range_string:
            range_string = '%s%s:%s%s' % (get_column_letter(start_column),
                                          start_row,
                                          get_column_letter(end_column),
                                          end_row)
        elif ":" not in range_string:
            if COORD_RE.match(range_string):
                return  # Single cell, do nothing
            raise ValueError("Range must be a cell range (e.g. A1:E1)")
        else:
            range_string = range_string.replace('$', '')

        if range_string not in self._merged_cells:
            self._merged_cells.append(range_string)


        # The following is removed by this monkeypatch:

        # min_col, min_row, max_col, max_row = range_boundaries(range_string)
        # rows = range(min_row, max_row+1)
        # cols = range(min_col, max_col+1)
        # cells = product(rows, cols)

        # all but the top-left cell are removed
        #for c in islice(cells, 1, None):
            #if c in self._cells:
                #del self._cells[c]

    # Apply monkey patch
    m = types.MethodType(merge_cells, None, worksheet.Worksheet)
    worksheet.Worksheet.merge_cells = m


patch_worksheet()

