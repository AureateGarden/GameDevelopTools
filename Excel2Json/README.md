使用python3.7并且安装插件openpyxl

>pip install openpyxl

![](https://github.com/AureateGarden/GameDevelopTools/blob/master/Excel2Json/Temp/1.png)

用法在cmd里面使用“python Excel2Json.py -h” 查看具体参数

使用范例：

>python Excel2Json.py -i test.xlsx -o Output.json -s Sheet1

Develop Log:

2022.7.9

1. 放弃使用xlrd插件，现在使用openpyxl库进行xlsx读取。
2. 修复了之前输出路径的Bug。
3. 现在不再支持随机开始位置。[start]标签被遗弃。

**添加了新的特性**
1. 现在单元格可以支持Json数据格式了，但是暂时还不支持Json数据格式嵌套。也就是说可以有用户自定义数据了。

2019.8.10

**添加了新的特性**
1. [ignore]标签和空白行可以被忽略。
2. 现在[ignore] 和 space 可以被省略了。
