# UnityTools-Excel
## 说明
1. 使用EPPlus.dll实现简单的Excel创建数据表格，导出导入等功能。
2. 通过Unity编辑器扩展，添加两个工具按钮：创建表格以及导出表格数据
## 使用方式
1. 点击Tools/Create Table：选择创建的数据类型，以及名称，点击确定即可生成对应的Excel
2. 点击Tools/Export Table：选择对应数据类进行单个解析或者点击解析所有Excel，之后会生成对应的.byte以及.json文件。
## 注释
1. byte文件用于实际运行时的数据读取
2. json文件用于方便查看每次的修改数据
3. 创建的Excel名称为xxx@type，多个同类型的Excel在解析时都会解析成一个整体，但是主键不能存在重复的
4. 只提供了默认Editor下的解析帮助类，实际使用需要自行加载资源