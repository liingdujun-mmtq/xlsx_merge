import tkinter as tk
import sys
from tkinter import scrolledtext
from tkinter import messagebox
from tkinter import END
from tkinter import filedialog
import os
import openpyxl
#v0.92
#fix: location_split(A)
#fix: open xlsx files with data_only


#read Merge Location
#tansfrom “A1:4” to ['A1','A2,'A3','A4]
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
        print(merge_data)
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
