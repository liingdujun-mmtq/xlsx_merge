# xlsx_merge

#### 介绍
xlsx_merge：将多个xlsx文件中指定位置的数值累加并写入另一个xlsx文件中。

#### 使用说明

1.  Merge Files：需要合并的xlsx文件所在的文件夹。所有需要合并的xlsx文件必须在同一个文件夹，且暂时不支持加载子文件夹中的文件。
2.  Template File：数据合并后写入的xlsx文件
3.  Merge Location：需要合并的数据在xlsx文件中的位置。

#### Merge Location支持的写法

Merge Location支持如下形式的写法：

- 单独指明每个需要合并单元格的位置（以英文逗号分隔），例如：A1,A2,C3,D9(目前存在一个bug，形如“A1”这样的形式需要写成“A1:1”否则报错，下个版本会修正问题)
- 按Excel的方式一次指定同列的多个单元格，例如：A3:9
- 混合写入以上两种方式，以英文逗号隔开


