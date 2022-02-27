#!/usr/bin/env python
# -*- encoding: utf-8 -*-
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm

doc = DocxTemplate("my_word_tpl.docx") # 读取模板


# 配置插入的图片


image1_path = r'img/0_01.png' # 要生成的图片地址
image2_path = r'img/0_02.png' 
insert_image1 = InlineImage(doc, image1_path,width=Mm(140))
insert_image2 = InlineImage(doc, image2_path, width=Mm(140))

# 作为图片的替换
img_context = {
    'company_name': insert_image1,
    'seq':['{{a}}','{{b}}','{{c}}']
}


# context = {**img_context} # 需要传入的字典， 需要在word对应的位置输入 {{ company_name }}
# context = {'company_name':img_context}
doc.render(img_context) # 渲染到模板中
doc.save("generated_image_doc.docx") # 生成一个新的模板
print('finish')