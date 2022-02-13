---
layout: post
title: 操作word
date: 2021-01-27 
tags: Python    
---
# python write word
## 需要安装的库
```shell
pip3 install python-docx
```
## 需要引用的包
```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
#创建文档，()中含路径则为打开文档
document = Document()
'''
对文档的操作
'''
#保存文档
document.save('./demo.docx')
```
## 文档添加标题及内容
```python
#添加文档标题
document.add_heading('Document Title', 0)
#插入第一段内容
document.add_paragraph('This is my first paragraph.')
#插入第二段内容
p = document.add_paragraph('This is my first paragraph.')
#在第二段前面插入一段
p.insert_paragraph_before('This paragraph is between first and second')
#加一个分页符
document.add_page_break()

#add_run()在第二段中追加内容，bold是粗体，italic是斜体
p.add_run('bold').bold = True
p.add_run(' and ')
p.add_run('italic.').italic = True
```
## 文档中字体的设定
* 颜色设定
```python
a = 255
white = RGBColor(a, a, a)
yellow = RGBColor(a, a, 0)
red = RGBColor(a, 0, 0)
blue = RGBColor(0, 0, a)
black = RGBColor(0, 0, 0)
green = RGBColor(0, a, 0)
p = document.add_paragraph('')
run = p.add_run('hello world')
#设定"hello world"为红色
run.font.color.rgb = red
```
* 字体字号设定
```python
#正文字号设置
document.styles['Normal'].font.size = Pt(12)
```
word字号对应关系如下
|  单位  | | | | | | | | | | | | | | | |
| :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
| 字号  | 八号 | 七号 | 小六 | 小五 | 五号 | 小四 | 四号 | 小三 | 三号 | 小二 | 二号 | 小一 | 一号 | 小初 | 初号 |
| 磅值 (Pt)  | 5 | 5.5 | 6.5 | 7.5 | 9 | 10.5 | 12 | 14 | 15 | 16 | 18 | 22 | 24 | 26 | 36 | 42 |
* 字体颜色设定
```python
#英文正文字体设置
document.styles['Normal'].font.name = u'Times new Roman'
#中文正文字体设置
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
```
* 段落设定
```python
from docx.enum.text import WD_LINE_SPACING
p = document.add_paragraph('正文第一段 Hello world')
#设定该段落为单倍行距
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
#整个段落左侧缩进
p = document.add_paragraph('this is a paragraph with 0.5 inches on the left.this is a paragraph with 0.5 inches on the left.')
paragraph_format = p.paragraph_format
paragraph_format.left_indent = Inches(0.5)
#首行缩进
p = document.add_paragraph('this is a paragraph with 2 inches on the left of the first line.this is a paragraph with 2 inches on the left of the first line.')
paragraph_format = p.paragraph_format
paragraph_format.first_line_indent = Inches(2)
```
## 文档中的表格
```python
#用来设置表格位置
from docx.enum.table import WD_TABLE_ALIGNMENT
records = (
    (3, '101', 'Spam'),
    (7, '422', 'Eggs'),
    (4, '631', 'Spam and Eggs')
)
#新建表格,风格是三线风格
table = document.add_table(rows=1, cols=3, style='Light Shading')
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'

for qty, id, desc in records:
    row_cells = table.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc
#表格位置居中
table.alignment=WD_TABLE_ALIGNMENT.CENTER
```
更多表格格式[点击这里](https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template)
## 文档中的图片
```python
#宽度设定为2英尺，1英尺=2.54CM，固定长宽比
def docx_add_picture(path, width):
	document.add_picture('monty-truth.jpeg', width=Inches(2))
	last_paragraph = document.paragraphs[-1]
	last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER      #图片居中设置
```

### **批量设置题注**
* **Windows word 版本**
1. 选择题注，按下【Alt+F9】快捷键，切换到域代码状态。

2. 选中域代码，这里不包含后面的图片注释文字。按【Ctrl+C】键复制，再按【Ctrl+H】键打开“查找和替换”对话框。

3. 在查找文本框中输入【^g】；在替换文本框中输入【^&^p^c】，然后点击【全部替换】按钮。

4. 此时，所有图片下方均添加了编号。按下【Ctrl+A】键全选所有内容，再按【F9】刷新，最后，再按【Alt+F9】切回正常状态，这样就完成了图片的编号

* **Mac word 版本**
1.  选择题注，按下【Option+Fn+F9】快捷键，切换到域代码状态。

2. 选中域代码，这里不包含后面的图片注释文字。按【Command+C】键复制，再按【Command+H】键打开“查找和替换”对话框。

3. 在查找文本框中输入【^g】；在替换文本框中输入【^&^p^c】，然后点击【全部替换】按钮。

4. 此时，所有图片下方均添加了编号。按下【Command+A】键全选所有内容，再按【Fn+F9】刷新，最后，再按【Option+Fn+F9】切回正常状态，这样就完成了图片的编号

* **说明**
^g 表示图片；^& 指要查找的内容；^p 表示换行符（也就是换一行）；^c 表示剪切版的内容（即第二步中复制的代码内容