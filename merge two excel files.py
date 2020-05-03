''' In this we can merge the two diffrent excel file with same formate it's merge them and create a new one''''

import pandas as pd
import os
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as msg

root=Tk()
root.title('Excel satish-Made by Monu Sharma')
root.configure(bg='lightblue')


def reset():
    #print('this is reset button') #prints msg on console
    s=entry_browse.get()
    ln=len(s)
    entry_browse.delete(0,ln) #to clear username text field

    s=entry_name.get()
    ln=len(s)
    entry_name.delete(0,ln)
    #Lb1.delete(0,END)

def browse():
    path=filedialog.askdirectory()
    path+='/'
    entry_browse.insert(0,path)
    
def work(path,new_filename):
    list=os.listdir(path)
    count=0
    l=[]
    for i in list:
        if(i.endswith('.xlsx')):
            ready_path=path+i
            df=pd.read_excel(ready_path)
           # print(df.head())
            l.append(df)
    df=pd.DataFrame()
    for i in range(len(l)):
        if(count!=0):
            df=df.append(l[i],ignore_index=True)
        else:
            df=pd.DataFrame(l[i])
           # print(df)
            count+=1

    df.to_excel(path+new_filename+'.xlsx',index=False)
    reset()
    msg.showinfo('Info',f'File is created in same folder\n file name is {path+new_filename}.xlsx')
    
lbl_browse=Label(root,text="Directory Path",font=('cambria',15,'italic'),bg="lightblue")
lbl_browse.place(x=50,y=50)

entry_browse=Entry(root,font=('cambria',15),bg='lightblue',bd=5)
entry_browse.place(x=200,y=50)

btn_browse=Button(root,command=browse,text='Browse',font=('Book Antiqua',15),bg='lightblue')
btn_browse.place(x=450,y=50)

filename=Label(root,text="File Name",font=('cambria',15,'italic'),bg="lightblue")
filename.place(x=50,y=100)

entry_name=Entry(root,font=('cambria',15),bg='lightblue',bd=5)
entry_name.place(x=200,y=100)

sub_btn=Button(root,command=lambda: work(entry_browse.get(),entry_name.get()),text='Submit',font=('Book Antiqua',15),bg='lightblue')
sub_btn.place(x=350,y=150)
    
root.geometry('600x200+350+250')
root.mainloop()