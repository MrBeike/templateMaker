# %%
import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment

# %%
f = pd.ExcelFile('bill.xlsx')
sheet = f.sheet_names
writer = pd.ExcelWriter('newbill.xlsx')

# %%
dataSet = []
for i in sheet:
    data = pd.read_excel('bill.xlsx', sheet_name=i)
    # 解除单元格合并，前置填充
    data["字段中文名称"] = data["字段中文名称"].ffill()
    # 利用groupby合并同样分组，并汇总填写规则（合并单元格实质还是多个单元格，此处合并为一个单元格，sort=False保持原来顺序）
    dataNew = data.groupby(["字段中文名称"],sort=False)['填写规则'].apply(list).to_frame()
    # 前面利用list合并填写规则，此处将list转换为字符串【lambda x: ''.join(x)）报错，暂时没弄明白】
    dataNew['填写规则'] = dataNew['填写规则'].apply(lambda x:str(x).replace("[","").replace("]","").replace("'",""))
    # 转置
    dataTransport = dataNew.T
    # 写入文件，注意index=False，不写入索引
    dataTransport.to_excel(writer, sheet_name=i,index=False)
writer.save()
writer.close()

# %%
f = pd.ExcelFile('newbill.xlsx')
sheet = f.sheet_names
wb = load_workbook('newbill.xlsx')
# 将第二行的数据当成批注写入标题行
for i in sheet:
    ws = wb[i]
    for column in range(1,ws.max_column+1):
        # 获取第二行单元格内容
        commentContent = ws.cell(row=2,column=column).value
        # 构建Comment对象,（内容，作者）
        comment = Comment(commentContent, 'xeroxYor')
        # 添加批注
        ws.cell(row=1,column=column).comment = comment
    # 删除第二行内容   
    ws.delete_rows(2)
    # 恢复第二行行高到Excel默认值
    ws.row_dimensions[2].height = 13.5
wb.save('template.xlsx')
# %%
