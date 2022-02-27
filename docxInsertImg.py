#!/usr/bin/env python
# -*- encoding: utf-8 -*-
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm
import sys
import os
import pandas as pd

'''
整体思路：
    1.根据文件夹 下的文件个数多少，先插入 {{<img_T1_1_xxxx}},生成中间文件
    2.根据中间文件再插入img
'''



def getImgFiles():
    # 读取 img/1-课题 下有多少文件目录情况
    target_paths = 'img/1-课题/'
    files_arr = []
    for root, dirs, files in os.walk(target_paths):
        parent_paths = root.replace(target_paths,'')
        for file in files:
            if file.endswith('.png'):
                fileName = file.replace('.png','')
                temp_arr = []
                temp_arr.append(parent_paths)
                temp_arr.append(fileName)
                temp_arr.append(root+"/"+file)
                files_arr.append(temp_arr)

    df = pd.DataFrame(data=files_arr,columns=['title','file_name','file_paths'])
    return df


# 向docx 模版中拆入普通文本
def renderVarDocx(df,tpl_docx):
    doc = DocxTemplate(tpl_docx) # 读取模板
    var_context = {}
    for title in df['title'].drop_duplicates():
        df_temp = df[df['title'] == title]
        df_temp = df_temp.sort_values(by=['file_name'])
        file_names = df_temp['file_name'].values.tolist()
        # 配置变量
        var_context[title] = [ '{{'+title+'_'+i+'}}' for i in file_names]
    
    
    # 将变量渲染到模版中
    doc.render(var_context) # 渲染到模板中
    temp_docx = 'tpl_temp.docx'
    doc.save(temp_docx)
    print(temp_docx,'  finish')
    return temp_docx
    
# 向docx 模版中 插入图片
def renderImageDocx(df,tpl_docx):
    doc = DocxTemplate(tpl_docx) # 读取模板
    image_context = {}
    for title in df['title'].drop_duplicates():
        df_temp = df[df['title'] == title]
        # 配置图片
        for t,file_name,file_paths in df_temp.values.tolist():
            image_var = t + '_'+file_name
            image_path = r''+file_paths # 要生成的图片地址
            insert_image = InlineImage(doc, image_path,width=Mm(150))
            image_context[image_var] = insert_image
     # 将图片渲染到模版中
    doc.render(image_context) # 渲染到模板中
    temp_docx = 'tpl_final.docx'
    doc.save(temp_docx)
    print(temp_docx,'  finish')
    return temp_docx

# 获取文件
df = getImgFiles()
# 生成变量模版文件
temp_tpl_docx = renderVarDocx(df,'tpl.docx')
# 生成最终文件
renderImageDocx(df,temp_tpl_docx)

print('finish')









