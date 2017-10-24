# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.shortcuts import render
from django.http import HttpResponse


import requests
from bs4 import BeautifulSoup
from itertools import product
import types
import openpyxl
from openpyxl import worksheet
from openpyxl import load_workbook
from openpyxl import Workbook


# Create your views here.

def index(request):
   
#登录部分

    root_url = 'http://172.16.203.12/zentao/user-login.html'
    index_url = 'http://172.16.203.12/zentao/project-task-206.html'
    UA = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36"

    header = {"User-Agent": UA,
               "referer":"http://172.16.203.12/zentao/my/"
               }

    
    r = requests.Session()
    f = r.get(root_url,headers = header)
    '''
    r.cookies = requests.utils.cookiejar_from_dict({
        
        'zentaosid':'buppolukrhfdeefbce93rjigc7'})
    r.post(root_url,
        cookies = r.cookies,        
        headers = header
        )
    '''
    postdata = {
        'account':'baozhong',
        'password':'111111'
    }
    r.post(
        root_url,
        data = postdata,
        headers = header)
    

#抓取数据部分

  

    f = r.get(index_url,headers = header)

    soup = BeautifulSoup(f.content,'lxml')

    plans = soup.find_all('tr',class_='text-center')

    names = []
    times = []
    jindus = []

    for plan in plans:
        l = list(plan.stripped_strings)
        if l[2] == '进行中':
            names.append(l[1])
            status = l[2]
            times.append(l[3])
            jindus.append(l[-1])
            print l[1],status,l[3],l[-1],'\n'
    context = {
        'names1':names[0],'status':status,'jindus1':jindus[0],
        'names2':names[1],'status':status,'jindus2':jindus[1],
        'names3':names[2],'status':status,'jindus3':jindus[2]
        } 

    


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



    wb = load_workbook("项目周报.xlsx")
    ws = wb.active
    wb.save('项目周报.xlsx')
    return render(request,'weekly_report/index.html',context)



