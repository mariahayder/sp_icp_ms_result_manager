#-*- coding: utf-8 -*-
import os

from frequency2size_chart_creator import *
from statistic_table_creator import *

#os.chdir("files")
import string
project_names=[]
all_names = [ name.split('.')[0] for name in os.listdir('files') ]
for name in os.listdir('./files'):
    if("table" in name):
        TABLE_FILENAME=name
print("table: {}".format(TABLE_FILENAME))

for name in os.listdir('./files'):
  #  if ("CHART" in name):
#        os.remove(name)
    if("xlsx" in name and "table" not in name and "lock" not in name and "CHART" not in name and "results" not in name):
        project_names.append(name)
    #if("table" in name):
     #   TABLE_FILENAME=name
    if ("EMF" in name):
        try:
            os.chdir("files")
        except(FileNotFoundError):
            pass
        SMPL=find_smpl_number(name)
        emf_content=find_true_emf_name(TABLE_FILENAME, SMPL)
        os.rename('{}'.format(name), '{}_{}.EMF'.format(SMPL, emf_content))
print(project_names)
print(TABLE_FILENAME)
#os.chdir("files")
#dic={'key':value, 'key':value}
size_data={}
median_data={}
for name in project_names:
    FILENAME = '%s' % name
    TABLE_FILENAME = 'table.xlsx'
    smpl_content, size=create_chart(FILENAME, TABLE_FILENAME)
    max_size=size[0]
    median_size=size[1]
    if(smpl_content not in size_data):
        size_data.update({smpl_content:max_size})
    else:
        previous=size_data[smpl_content]
        for size in max_size:
            previous.append(size)
        size_data.update({smpl_content:previous})
    if (smpl_content not in median_data):
        median_2_list=[]
        median_2_list.append(median_size)
        median_data.update({smpl_content: median_2_list})
    else:
        previous = median_data[smpl_content]

        previous.append(median_size)
        median_data.update({smpl_content: previous})

print(size_data)
print(median_data)

create_table(size_data, median_data)





