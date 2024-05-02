import tkinter as tk
import sys
from tkinter import scrolledtext
from tkinter import messagebox
from tkinter import END
from tkinter import filedialog
import os
import openpyxl
#v0.99
# Full support excel style location! @ 2024.5.2


## A self_add function for Excel style 26 base num
# where [A-Z] was used instead of [0-9] and [A-F]
def self_add_ABC(st):
    if len(st) == 0:
        return "A"
    s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Inverse_st = st[::-1]
    for i in range(26):
        if s[i] == Inverse_st[0]:
            if i < 25:
                return_st = st[:-1]+s[i+1]
                return return_st
            elif i == 25:
                return_st = self_add_ABC(st[:-1])+"A"
                return return_st

## A excel style location tansfrom function
# Full support excel style
# eg:
# A1:4 to A1,A2,A3,A4
# A2:B4 to A2,A3,A4,B2,B3,B4
# Replace old "location_split" at 2024.5.2
def location_split(input_str):
    A=input_str.split(',')
    location=[]
    for j in A:
        if ':' in j:
            i=j.split(':')
            location_start=i[0]
            location_end=i[1]
            location_start_split=re.split("([0-9])", location_start, maxsplit=1) 
            location_start_ABC=location_start_split[0]
            location_start_NUM=location_start_split[1]
            if location_end[0] in '1234567890':
                location_end_ABC=location_start_ABC
                location_end_NUM=location_end
            else:
                location_end_split=re.split("([0-9])", location_end, maxsplit=1)
                location_end_ABC=location_end_split[0]
                location_end_NUM=location_end_split[1]

            max_loop_num=1000
            location_now_ABC=location_start_ABC
            for loop_num in range(max_loop_num):
                if int(location_end_ABC,base=36)>= int(location_now_ABC,base=36):
                    for NUM in list(range(int(location_start_NUM),int(location_end_NUM)+1)):
                        location.append(location_now_ABC+str(NUM))
                    location_now_ABC=self_add_ABC(location_now_ABC)          
            break
        else:
            location.append(j)
    return location


#read Merge Location
# tansfrom “A1:4” to ['A1','A2,'A3','A4]
# Replaced by new function at 2024.5.2
'''
def location_split(A):
    location=[]
    for j in A:
        if ':' in j:
            for i in range(len(j)):
                if j[i] in '1234567890':
                    location_ABC=j[0:i]
                    [location_NUM_start,location_NUM_end]=j[i:].split(':')
                    location_NUM=list(range(int(location_NUM_start),int(location_NUM_end)+1))
            
                    for i in location_NUM:
                        location.append(location_ABC+str(i))
                    break
        else:
            location.append(j)
    return location
'''

def main_window():
    root=tk.Tk()
    root.title('Xlsx_merge by Lingdujun')

    main_frame=tk.Frame(root)
    main_frame.pack(anchor='w')

    merge_files_path=tk.StringVar(root,value="")
    template_file_path=tk.StringVar(root,value="")
    set_merge_locatons=tk.StringVar(root,value="A1,A2")

    tk.Label(main_frame,text="Merge Files:",width=16).grid(row=0,column=0)
    tk.Entry(main_frame,textvariable=merge_files_path,state='readonly',width=50,justify='left').grid(row = 0, column = 1)
    tk.Label(main_frame,text="Template_file:",width=16).grid(row=1,column=0)
    tk.Entry(main_frame,textvariable=template_file_path,state='readonly',width=50,justify='left').grid(row = 1, column = 1)
    tk.Label(main_frame,text="Merge Location:",width=16).grid(row=2,column=0)
    tk.Entry(main_frame,textvariable=set_merge_locatons,width=50,justify='left').grid(row = 2, column = 1)

    def select_merge_files():
        merge_files_path_=filedialog.askdirectory()
        merge_files_path.set(merge_files_path_)

    def select_template_file():
        template_file_path_=filedialog.askopenfilename()
        template_file_path.set(template_file_path_)


    def merge():
        xlsx_path=merge_files_path.get()
        template_path=template_file_path.get()
        all_files=os.listdir(xlsx_path)
        xlsx_files=[f for f in all_files if '.xlsx' in f]

        input_merge_data=set_merge_locatons.get()
        merge_location=location_split(input_merge_data.split(','))
        merge_data=[]

        for i in merge_location:
            merge_data.append(0)

        for filename in xlsx_files:
            workbook=openpyxl.load_workbook(xlsx_path+'/'+filename,data_only=True)
            sheet= workbook.worksheets[0]
            for i in range(len(merge_location)):
                if sheet[merge_location[i]].value != None:
                    merge_data[i]=sheet[merge_location[i]].value+merge_data[i]
                    workbookmerge=openpyxl.load_workbook(template_path)
                    sheetmerge= workbookmerge.worksheets[0]
        
        for i in range(len(merge_location)):
            sheetmerge[merge_location[i]].value=merge_data[i]
            workbookmerge.save(template_path)
            workbookmerge.close()
        messagebox.showinfo('','All Done!')

    tk.Button(main_frame, text = "Select Merge Files", command = select_merge_files,width=15).grid(row = 0, column = 2)
    tk.Button(main_frame, text = "Select Template_File", command = select_template_file,width=15).grid(row = 1, column = 2)
    tk.Button(main_frame, text = "Merge!", command = merge,width=15).grid(row = 2, column = 2)
    root.mainloop()

main_window()
