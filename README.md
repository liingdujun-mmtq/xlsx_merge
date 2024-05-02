# xlsx_merge

#### 介绍
xlsx_merge：将多个xlsx文件中指定位置的数值累加并写入另一个xlsx文件中。

#### 使用说明

1.  Merge Files：需要合并的xlsx文件所在的文件夹。所有需要合并的xlsx文件必须在同一个文件夹，且暂时不支持加载子文件夹中的文件。
2.  Template File：数据合并后写入的xlsx文件
3.  Merge Location：需要合并的数据在xlsx文件中的位置。

#### Merge Location支持的写法

Merge Location已经完全支持Excel风格的定位方法，具体如下写法：

- 单独指明每个需要合并单元格的位置（以英文逗号分隔），例如：A1,A2,C3,D9
- 按Excel的方式一次指定同列的多个单元格，例如：A3:6，此时等同于A3,A4,A5,A6
- Excel风格跨列写法，例如A3:B5,此时等同于A3,A4,A5,B3,B4,B5
- 混合写入以上三种种方式，以英文逗号隔开

注意：英文字母必须全大写！！！


