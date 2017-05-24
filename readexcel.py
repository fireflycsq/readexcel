#!/usr/bin/env python
#-*- coding:utf-8 -*-
import xlrd
import re
import sys
data = xlrd.open_workbook(r'D:\learning\pylearn\test.xlsx')
datares=open(r'D:\learning\pylearn\test518.txt','w')
table = data.sheets()[3]
# cell_A1 = table.cell(4,2).value
# cell_A1 = table.row(2609)[0].value
nrows = table.nrows
print nrows
# # print cell_A1.encode('utf-8')
# (weibo, num) = re.subn('\n', " ", cell_A1.encode('utf-8'))
# (weibo,num)=re.subn('\r', " ", weibo)
# # weibo=cell_A1.encode('utf-8').replace('\r\n',' ')
# # print (weibo,num)
# print weibo
# label=table.row(1)[4].value
# print label
for i in range(1,nrows):
    line=table.row(i)[0].value.encode('utf-8')
    # score=table.row(i)[1].value
    label=table.row(i)[4].value
    score=table.row(i)[1].value
    if label=="":
        label=0
    # re.subn(, "", weibo)
    (weibo, num) = re.subn('\n', " ", line)
    (weibo, num) = re.subn('\r', " ", weibo)
    datares.write(str(label)+'\t'+str(score)+'\t'+weibo+'\n')
    if i%10000==0:
        print >> sys.stderr, 'loading %s(%s)' % ('test.xlsx', i) 
