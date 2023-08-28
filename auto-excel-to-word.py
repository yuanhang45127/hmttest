import pandas as pd
from docxtpl import DocxTemplate,RichText
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt,RGBColor
import re
# *一些基础设置与函数
def number(str_list):
    # *将一开始的编号删除
    test=list()
    for i in range(len(str_list)):
        for j in range(len(str_list[i])):
            pattern=re.compile(r'[0-9]+.')
            test.append(pattern.sub("",str_list[i][j]))
    return test
def addnumber(listresult,doc,title):
    # *将对应数值按照列开始放入
    result=list()
    doc.add_paragraph(title)
    for i in range(len(listresult)):
        doc.add_paragraph(str(i+1)+'.'+listresult[i])
def readdata(dict_list,number):
    # * 最开始读取数据的位置，将某一列的数值进行读取，然后将每一条目分开变成数组,分开的数组为missions
    mission=list()
    missionwithname=list()
    missions=list()
    for i in range(len(people)):
        mission.append(dict_list[2+i]['Unnamed: '+str(number)])
        #这里放number3与5
    for i in range(len(mission)):
        missionwithname.append(re.sub('\n','【'+people[i]+'】'+'\n',mission[i])+'【'+people[i]+'】')
    for i in range(len(missionwithname)):
        missions.append(re.split('\n',missionwithname[i]))
        # for j in range(len(mission[i])):
        #     missionwithname.append(mission[i])
        # mission[i].sub('\n',people[i])
    return missions
    # *这里就变成了split
def splitbyprandfengzhuang(missions): 
    # *将pr和封装分开放进去
    fengzhuang=list()
    pr=list()
    for i in range(len(missions)):
        if i==0 or i==3 or i==7:
            fengzhuang.append(missions[i]) 
        else:
            pr.append(missions[i])
    return fengzhuang,pr

if __name__ == '__main__':
    doc=Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(10.5)
    doc.styles['Normal'].font.color.rgb = RGBColor(0,0,0)

    # * 这里开始使用
    data = pd.read_excel('meetingtest.xlsx')
    dict_list = data.to_dict(orient="records")
    """
    data_dict[2]开始是王贺
    """
    wordname='result.docx'
    people=['王贺','刘婧恬','杨彬彬','桂羽','李峥','胡梦恬','卜苗苗','邢玉婷']
    # 这里添加的上周工作小结
    # 王贺、刑玉婷、桂羽 封装,对应0，3，7
    missions=readdata(dict_list,3)
    fengzhuang,pr=splitbyprandfengzhuang(missions)
    listpr=number(pr)
    listfengzhuang=number(fengzhuang)
    addnumber(listpr,doc,'PR与innovus脚本调试方面：')
    addnumber(listfengzhuang,doc,'封装工作方面：')
    doc.save(wordname)
    missions=readdata(dict_list,5)
    fengzhuang,pr=splitbyprandfengzhuang(missions)
    listpr=number(pr)
    listfengzhuang=number(fengzhuang)
    # 注意不能有空白段落
    doc.add_paragraph('下周计划')
    addnumber(listpr,doc,'PR与innovus脚本调试方面：')
    addnumber(listfengzhuang,doc,'封装工作方面：')
    doc.save(wordname)

