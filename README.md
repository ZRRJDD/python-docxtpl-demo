
# 一、docxtpl

docxtpl 一个很强大的包，其主要通过对docx文档模板加载，从而对其进行修改。
主要依赖两个包

- python-docx ：读写doc文本
- jinja2：管理插入到模板中的标签

因为模板标签主要来自jinja2，可以了解其语法
[http://docs.jinkan.org/docs/jinja2/templates.html](http://docs.jinkan.org/docs/jinja2/templates.html)


docxtpl 英文参考文档 

[https://docxtpl.readthedocs.io/en/latest/index.html](https://docxtpl.readthedocs.io/en/latest/index.html)


## 安装

```bash
pip install docxtpl
```

基本使用示例
```python
from docxtpl import DocxTemplate
doc = DocxTemplate("my_word_template.docx")
context = { 'company_name' : "World company" }
doc.render(context)
doc.save("generated_doc.docx")
```

## 参考文档
[https://www.cnblogs.com/c-keke/p/14831821.html](https://www.cnblogs.com/c-keke/p/14831821.html)



