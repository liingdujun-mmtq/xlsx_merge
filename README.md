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


