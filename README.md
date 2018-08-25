# excel导出xml工具

## 脚本

convert.py

## 使用方法

修改`convert.ini`文件,
分别配置excel目录,xml目录,和导出文件格式目录
(导出文件格式目录如果为空,则不生成导出文件格式)

## 生成exe的方法

`pyinstaller -F -p xlrd convert.py`

`-p xlrd`为包含的xlrd的库

