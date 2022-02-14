# 金数报文EXCEL模板生成工具

由字段说明文件生成EXCEL模板文件,便于无报文系统的小金融机构制作Excel模板填写报文内容,结合报文生成小程序生产最终Dat、log文件上报系统。

## 使用方法
 + 从字段说明word文件中复制表格到excel文件中(手动,data/bill.xlsx)
 + 打开`main.py`,输入目标文件路径(data/bill.xlsx)【支持拖拽】
 + 去除合并单元格，整理需要的字段(程序实现,data/newbill.xlsx)【仅为展示而保存,生成后可删除】
 + 生成带批注的模板文件(程序实现,data/【template】bill.xlsx)【为了美观，简单的调整了背景色和边框等样式】
  
**`bill.ipynb`为刚开始快速测试写的notebook。运行可大致了解处理思路及过程。**【菜鸟选手，大神轻喷】