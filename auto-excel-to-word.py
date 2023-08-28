import pandas as pd
from docxtpl import DocxTemplate
import re

data = pd.read_excel('meetingtest.xlsx')
dict_list = data.to_dict(orient="records")
"""
data_dict[2]开始是王贺
"""
people=['王贺','刘婧恬','杨彬彬','桂羽','李峥','胡梦恬','卜苗苗','邢玉婷']
# 这里添加的上周工作小结
mission=list()
missionwithname=list()
for i in range(len(people)):
    mission.append(dict_list[2+i]['Unnamed: 3'])
for i in range(len(mission)):
    missionwithname.append(re.sub('\n','【'+people[i]+'】'+'\n',mission[i])+'【'+people[i]+'】')
    # for j in range(len(mission[i])):
    #     missionwithname.append(mission[i])
    # mission[i].sub('\n',people[i])
for i in dict_list:
    tpl = DocxTemplate('temp.docx')
    tpl.render(i)
    tpl.save(i['姓名']+'.docx')
