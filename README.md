# xlsx_merge

xlsx_merge is a APP used to merge multiple xlsx file with the same format,sum the specified positions and write them into a new table
xlsx_merge 是一个将多个xlsx文件中指定位置的数值累加并写入另一个xlsx文件中的APP。

#### How to use/使用说明

1.  Merge Files：xlsx files need to be merged, all files should be located in a singel folder and sub-folder is not support
2.  Template File：merged data will write into the Template File
3.  Merge Location：the location in Template File (suce as A1, A2, C3 or A3:A7). Only the data in Merge Location will be merged.

1.  Merge Files：需要合并的xlsx文件所在的文件夹。所有需要合并的xlsx文件必须在同一个文件夹，且暂时不支持加载子文件夹中的文件。
2.  Template File：数据合并后写入的xlsx文件
3.  Merge Location：需要合并的数据在xlsx文件中的位置。

#### Merge Location

You can use the following format in Merge Location

- Write out each cell separately, and separate them with ",", such as: A1,A2,C3,D9
- Write out multiple cells in the same column in Excel format, such as: A3:6, which is equivalent to A3,A4,A5,A6 
- Excel style multiple cells, such as A3:B5, which is equivalent to A3,A4,A5,B3,B4,B5 
- Mix the above three methods, and separate them with ",", such as: A3:9, B2:5, C8

Note! All letters must be capitalized!

Merge Location已经完全支持Excel风格的定位方法，具体如下写法：

- 单独指明每个需要合并单元格的位置（以英文逗号分隔），例如：A1,A2,C3,D9
- 按Excel的方式一次指定同列的多个单元格，例如：A3:6，此时等同于A3,A4,A5,A6
- Excel风格跨列写法，例如A3:B5,此时等同于A3,A4,A5,B3,B4,B5
- 混合写入以上三种种方式，以英文逗号隔开

注意：英文字母必须全大写！！！

## Relasea/下载

#### Windows user/Windows用户
You can download a exe file build for windows system at Release Page.
直接从发行版页面下载即可

#### Other user/其他用户
If the all used module has been installed, you can execute the py file directly.
在所有python3 module都被安装的情况下，可以直接执行py文件

The list of used module is here:
* openpyxl

以下模组是必须的：
* openpyxl

## Acknowledgements/致谢
Thanks for the testing and support from Hefei University of Technology (HFUT, Hefei, China).
感谢合肥工业大学提供的测试与支持
