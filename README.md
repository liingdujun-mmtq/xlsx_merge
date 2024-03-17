# xlsx_merge
xlsx_merge is a APP used to merge multiple xlsx file with the same format,sum the specified positions and write them into a new table

#### How to use

1.  Merge Files：xlsx files need to be merged, all files should be located in a singel folder and sub-folder is not support
2.  Template File：merged data will write into the Template File
3.  Merge Location：the location in Template File (suce as A1, A2, C3 or A3:A7). Only the data in Merge Location will be merged.

#### Merge Location

You can use the following format in Merge Location

- Write out each cell separately, and separate them with ",", such as: A1,A2,C3,D9
- Write out multiple cells in the same column in Excel format, such as: A3:9
- Mix the above two methods, and separate them with ",", such as: A3:9, B2:5, C8

Note! The Excel-like format such as "A3:C9" is not support now.

## Release

#### Windows user
You can download a exe file build for windows system at [release page](https://github.com/liingdujun-mmtq/xlsx_merge/releases).

##### System compatibility
Works on：
+ Windows 10 2004 64bit
+ Windows 8.1 64 bit
+ Windows 7 SP1 32bit (need KB2533623 update)

* For win7/8 user: If *.dll lost, you may need install [Visual C++ Redistributable for Visual Studio 2015](https://www.microsoft.com/en-us/download/details.aspx?id=48145)

#### Other user
If the all used module has been installed, you can execute the py file directly.

Note: Crystal_calculator_linux_mac.py is recommended for use in Linux/Mac 

The list of used module is here:
* openpyxl

## Acknowledgements
Thanks for the testing and support from Hefei University of Technology (HFUT, Hefei, China).
