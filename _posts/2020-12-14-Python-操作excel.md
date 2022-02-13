---
layout: post
title: 操作excel
date: 2020-12-14 
tags: Python    
---

# Python操作excel

## 安装必要库

```shell
pip3 install openpyxl
pip3 install pandas
```

## 读取excel
```python
import pandas as pd

def read_raw1(sheet_path):
    excel_content=pd.read_excel(sheet_path,sheet_name='Sheet')
    excel_header = excel_content.columns.tolist()
    return excel_content[excel_header[1]].values.tolist()

if __name__ == '__main__':
    raw1 = read_raw1('excel路径')
    print(len(raw1))
```

## 写excel

```python
def write_excel_xlsx(workbook, data):
    #取第一个表
    sheet = workbook.active
    #赋值操作
    sheet.cell(row=1, column=1, value=str(data))

def save_close_excel(workbook, sheet_path):
    workbook.save(sheet_path)
    workbook.close()
```