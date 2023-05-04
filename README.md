'''
import openpyxl as vb
from docx import Document # 导入docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT # 导入表格对齐方式 
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT # 导入单元格垂直对齐 
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT # 导入段落对齐
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn # 中文字体
from docx.shared import Pt,RGBColor #字号，颜色
from docx.shared import Cm # 导入单位转换函数
import time
import datetime as dt
import random


路径='/Users/redtea/Desktop/上海服务单生成-230413.xlsx'

工作簿 =vb.load_workbook(路径) 
服务质量 = 工作簿 ['平台服务质量统计表']
技术支持 = 工作簿 ['技术支持反馈']

def generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置):

    文件 = Document(模板位置)
    表=文件.tables[0]

    单元格 = 表.cell(4,1) # 指定单元格
    单元格.text=客户名称
    for 段落 in 单元格.paragraphs:
        for 块 in 段落.runs:
            块.font.name = 'Arial' # 英文字体设置 
            块._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 设置中文字体
            块.font.size = Pt(12)#小四-12号  四号-14 小五-9号

    单元格 = 表.cell(4,5) # 指定单元格
    单元格.text=客户联系人和手机
    for 段落 in 单元格.paragraphs:
        for 块 in 段落.runs:
            块.font.name = 'Arial' # 英文字体设置 
            块._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 设置中文字体
            块.font.size = Pt(12) #小四-12号  四号-14 小五-9号

    表.cell(8,0).text=需求时间
    表.cell(8,2).text=响应时间
    表.cell(8,4).text=完成时间    
    for 单元格 in 表.rows[8].cells:
        单元格.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER 
        单元格.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # 中部居中对齐
        for 段落 in 单元格.paragraphs:
            for 块 in 段落.runs:
                块.font.name = 'Arial' # 英文字体设置 
                块._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置中文字体
                块.font.size = Pt(9) #小四-12号  四号-14 小五-9号

    文件.save(保存位置)

def get_template_and_save_location(需求内容,客户名称,年月份):
    模板位置=r'/Users/redtea/Desktop/服务单模板/'+需求内容+'.docx'
    保存位置前缀=r'/Users/redtea/Desktop/服务单生成保存位置/'
    文件名称后缀=需求内容
    保存位置=保存位置前缀+客户名称+'-'+年月份+'-'+文件名称后缀+'.docx'   
    return 模板位置,保存位置

def 培训文件处理(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置):

    文件 = Document(模板位置)
    表=文件.tables[0]

    单元格 = 表.cell(4,1) # 指定单元格
    单元格.text=客户名称
    for 段落 in 单元格.paragraphs:
        for 块 in 段落.runs:
            块.font.name = 'Arial' # 英文字体设置 
            块._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 设置中文字体
            块.font.size = Pt(12)#小四-12号  四号-14 小五-9号

    单元格 = 表.cell(4,5) # 指定单元格
    单元格.text=客户联系人和手机
    for 段落 in 单元格.paragraphs:
        for 块 in 段落.runs:
            块.font.name = 'Arial' # 英文字体设置 
            块._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 设置中文字体
            块.font.size = Pt(12) #小四-12号  四号-14 小五-9号

    表.cell(9,0).text=需求时间
    表.cell(9,2).text=响应时间
    表.cell(9,4).text=完成时间    
    for 单元格 in 表.rows[9].cells:
        单元格.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER 
        单元格.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # 中部居中对齐
        for 段落 in 单元格.paragraphs:
            for 块 in 段落.runs:
                块.font.name = 'Arial' # 英文字体设置 
                块._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置中文字体
                块.font.size = Pt(9) #小四-12号  四号-14 小五-9号

    文件.paragraphs[3].text ='单位名称：'+客户名称
    文件.paragraphs[5].text = f'时间：{需求时间.split()[0]}    地点：会议室'
    for i in [文件.paragraphs[3],文件.paragraphs[5]]:
        for 块 in i.runs:
            # 块.font.bold = True # 加粗
            # 块.font.italic = True # 斜体
            # 块.font.underline = True # 下划线
            # 块.font.strike = True # 删除线
            # 块.font.shadow = True # 阴影
            块.font.size = Pt(14) #四号-14
            # 块.font.color.rgb = RGBColor(255,0,0) # 颜色
            块.font.name = 'Arial' # 英文字体设置 
            块._element.rPr.rFonts.set(qn('w:eastAsia'),'宋体') # 设置中文字体

    表=文件.tables[1]
    单元格 = 表.cell(1,0) # 指定单元格
    单元格.text='深圳自由连接科技有限公司'
    单元格 = 表.cell(6,0) # 指定单元格
    单元格.text=客户名称
    for 行 in 表.rows:
        for 单元格 in 行.cells:
            单元格.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER 
            单元格.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            # 中部居中对齐
            for 段落 in 单元格.paragraphs:                
                for 块 in 段落.runs:
                    块.font.name = 'Arial' # 英文字体设置 
                    块._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') # 设置中文字体
                    块.font.size = Pt(12) #四号-14  小四-12号  小五-9号

    文件.save(保存位置)



for 行 in 技术支持.rows:
    if 行[1].value=='需求内容':
        pass
    else:
        客户名称=行[0].value
        客户联系人=行[12].value
        手机号=行[13].value
        客户联系人和手机=行[14].value
        需求内容=行[1].value
        需求时间=行[7].value
        响应时间=行[9].value
        完成时间=行[11].value
        年月份=行[15].value

        if 需求内容=='综合解决方案':
            print('综合解决方案')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='销售支撑':
            print('销售支撑')
            模板位置=r'/Users/redtea/Desktop/服务单模板/销售支持-模组网络测试'+str(random.randint(1,7))+'.docx'
            保存位置前缀=r'/Users/redtea/Desktop/服务单生成保存位置/'
            文件名称后缀='销售支持-模组网络测试'
            保存位置=保存位置前缀+客户名称+'-'+年月份+'-'+文件名称后缀+'.docx'   
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='相关培训-网络安全介绍':
            print('相关培训-网络安全介绍')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            培训文件处理(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='相关培训-业务发展规划介绍':
            print('相关培训-业务发展规划介绍')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            培训文件处理(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='相关培训-AEP平台介绍':
            print('相关培训-AEP平台介绍')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            培训文件处理(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)


for 行 in 服务质量.rows:

    if 行[2].value=='需求内容':
        pass
    else:
        客户名称=行[0].value
        客户联系人=行[1].value
        手机号=行[14].value
        客户联系人和手机=行[15].value
        需求内容=行[2].value
        需求时间=行[8].value
        响应时间=行[10].value
        完成时间=行[12].value
        年月份=行[16].value              

        if 需求内容=='告警通知':
            print('告警通知')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='物联网卡连接管理':
            print('物联网卡连接管理')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='应用接入-套餐订购':
            print('应用接入-套餐订购')
            模板位置=r'/Users/redtea/Desktop/服务单模板/应用接入-套餐订购.docx'
            保存位置前缀=r'/Users/redtea/Desktop/服务单生成保存位置/'
            文件名称后缀='应用接入-套餐订购'
            保存位置=保存位置前缀+客户名称+'-'+年月份+'-'+文件名称后缀+'.docx'  
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)               

        elif 需求内容=='应用接入-流量查询':
            print('应用接入-流量查询')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)   

        elif 需求内容=='eSIM卡远程下发':
            print('eSIM卡远程下发')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='周期数据分析':
            print('周期数据分析')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)

        elif 需求内容=='客户终端和管理平台对接':
            print('客户终端和管理平台对接')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)            

        elif 需求内容=='物联网卡流量池管理':
            print('物联网卡流量池管理')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)     

        elif 需求内容=='账务数据校验':
            print('账务数据校验')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)   

        elif 需求内容=='智能交付':
            print('智能交付')
            模板位置, 保存位置 =get_template_and_save_location(需求内容,客户名称,年月份)
            generate_service_order(模板位置, 客户名称, 客户联系人和手机, 需求时间, 响应时间, 完成时间, 保存位置)   

print('over了')

'''
