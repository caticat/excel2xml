# excel导出xml工具

## 脚本

convert.py

## 使用方法

修改脚本配置(或打包的exe)同名配置文件(扩展名为`.ini``)(默认为`convert.ini`)文件,
分别配置excel目录,xml目录,和导出文件格式目录
(导出文件格式目录如果为空,则不生成导出文件格式)

### 参数(可选)

可通过参数的方式设置excel,xml和fmt的路径

参数:

- `-e`,`--excel=`:excel的路径配置
- `-x`,`--xml=`:xml的路径配置
- `-f`,`--fmt=`:fmt的路径配置

## 生成exe的方法

`pyinstaller -F -p xlrd convert.py`

`-p xlrd`为包含的xlrd的库

## 其他

### 换行符

python在使用`w`,`r`方式读写文件时,
强制会按照系统平台的默认换行符替换.
比如在windows下写`\n`会被直接替换成`\r\n`.
可以通过使用`wb`,`rb`(读写二进制流)的方式来
过滤掉这个替换.
例子:

```python
f = open("/tmp/tmp.txt", 'wb')
f.write("字符串数据\n".encode("utf8"))
f.close()
```

### 控制台编码

脚本调用的print,在控制台输出时,中文windows下是GBK编码
在其他程序调用时,想要不乱吗,需要手动转编码为合适的编码.

