使用python3.7并且安装插件xlrd

安装方法:

在有python\scripts路径情况下在cmd中输入下面的命令安装xlrd插件。

>pip install xlrd

![image](/Temp/1.png)
类似这样的列表，确定开始位置，在里面添加“[start]”标签且确保全表只有一个” [start]”标签，转化标志位后面一列为键，后面第二列之后转化为值。

用法在cmd里面使用“python Excel2Json.py -h” 查看具体参数

使用范例：

>python Excel2Json.py -i test.xlsx -o out.json -s 0

![image](/Temp/2.png)

Develop Log:

2019.8.10

**添加了新的特性**

1. [ignore]标签和空白行可以被忽略。
2. 现在[ignore] 和 space 可以被省略了。