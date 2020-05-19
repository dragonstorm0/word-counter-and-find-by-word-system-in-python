from tkinter import *
import slate3k
from tkinter import messagebox
from tkinter import ttk
from collections import Counter
from tkinter import filedialog
import re
import docx
import os
import docxpy




win=Tk()
win.state('zoomed')
win.configure(bg='cyan')
win.title('Word counting system')
win.resizable(width=False,height=False)

lbl_title=Label(win,text='Welcome To Word counting system',font=('',40,'bold'),bg='cyan',fg='red',pady='10px')
lbl_title.pack()
lbl_title_1=Label(win,text='- By Aayush Kumar',font=('',20,''),bg='cyan',fg='red',justify='right')
lbl_title_1.place(x=950,y=90)

txtFrame1=StringVar()
txtFrame2=StringVar()
txtFrame3=StringVar()
txtFrame4=StringVar()

def browse_files():
    file_path = filedialog.askopenfilename()
    txtFrame1.set(file_path)

def count():
    txt_value=txtFrame2.get()
    txt_value=txt_value.lower()
    
    path=txtFrame1.get()
    file=open(f'{path}','rb')
    if(path.endswith('.pdf')):
        pages=slate3k.PDF(file)
        text_pdf=pages.text()
        words=re.findall('[\w]+',text_pdf)
        lenght=len(words)
        print(words)
        for x in range(lenght):
            words[x]=words[x].lower()
                
        ctr=Counter(words)
        for x,y in ctr.items():
            if (x==txt_value):
                pointer=float(result_text.index(INSERT))
                result_text.insert(pointer,(f' \nthe word "{txt_value}" appeared : \n\n\t{x} \t\t:\t {y} times\n\n'))
                pointer +=1.0
            elif(txt_value not in words):
                messagebox.showwarning('NO WORD FOUND','No such word exist in this file.')
                break
            
        top_ten=ctr.most_common(11)
        highest_count=top_ten[0][1]
        x=(f'most occured word in the file is:\n\n \t {top_ten[0][0]} \t:\t {highest_count} times\n\n')
        pointer=float(result_text.index(INSERT))
        pointer +=1.0
        result_text.insert(pointer,x)
        pointer=float(result_text.index(INSERT))
        pointer +=1.0
        result_text.insert(pointer,('Top ten numbers on the basis of occurences are :\n\n'))
        main_len=len(top_ten)
        for n in range(main_len):
            pointer=float(result_text.index(INSERT))
            pointer +=1.0
            result_text.insert(pointer,(f' \t{top_ten[n][0]} \t:\t {top_ten[n][1]}\n\n'))
        result_text.insert(pointer,('\n\nCount of all the words that appeared in the file :\n\n'))
        for x,y in ctr.items():
            pointer=float(result_text.index(INSERT))
            pointer +=1.0
            result_text.insert(pointer,(f' \t{x} \t\t:\t {y}\n'))
            
             
    elif(path.endswith('.txt')):
        frequency = {}
        document_text = open((f'{path}'), 'r')
        text_string = document_text.read().lower()
        match_pattern = re.findall(r'\b[a-z]{3,15}\b', text_string)
         
        for word in match_pattern:
            count = frequency.get(word,0)
            frequency[word] = count + 1
        
            
        frequency_list = frequency.keys()
        result_text.insert(1.0,'The words appear on the txt file.\n\n') 
        for words in frequency_list:
             pointer=float(result_text.index(INSERT))
             pointer +=1.0
             result_text.insert(pointer,(f'{words}, {frequency[words]} time\n\n'))

        pointer=float(result_text.index(INSERT))
        pointer +=1.0    
        result_text.insert(pointer,f'The word {txt_value} appear on the txt file: \n\n')
        for words in frequency_list:
            if (words==txt_value):
                pointer=float(result_text.index(INSERT))
                pointer +=1.0
                result_text.insert(pointer,(f'{words}, {frequency[words]} time\n\n'))
            else:
                messagebox.showwarning('NO WORD FOUND','No such word exist in this file.')
                return

    elif(path.endswith('.docx')):
        doc = docx.Document(f'{path}')
        all_paras = doc.paragraphs
        main_list=[]
        count=0
        for para in all_paras:
            main_list.append(para.text.split())
        
        for x in range(len(main_list)):
            for y in range(len(main_list[x])):
                main_list[x][y]=re.sub(r"[^a-zA-Z0-9]+","",main_list[x][y]).lower()
        
        for x in range(len(main_list)):
            for y in range(len(main_list[x])):
                if(main_list[x][y]==txt_value):
                    count+=1

        result_text.insert(1.0,('the context of searched value is : \n\n'))
        pointer=float(result_text.index(INSERT))
        pointer +=1.0
        result_text.insert(pointer,(f'{txt_value}, {count} time\n\n'))
       
            
    else:
        messagebox.showwarning('WRONG SOURCE FILE.','sourse file is not of correct fomat pls select again.')
  
    
def reset(e1,e2):
    txtFrame1.set('')
    e2.delete(0,END)
    result_text.delete(1.0,10000000.0)

def browse_directory():
    dict_path = filedialog.askdirectory()
    txtFrame4.set(dict_path)

def find_files():
    word=txtFrame3.get()
    path=txtFrame4.get()
    print(path)
    all_file_names=os.listdir(path)#returns a list of all file names present in given dir
    pointer=float(result_text_word.index(INSERT))
    result_text_word.insert(pointer,("Total Files\t:\t"+(f'{len(all_file_names)}')+"\n\n"))
    pointer +=1.0

    docx_file_names=[]
    for filename in all_file_names:
        if(filename.endswith('.docx')):
            docx_file_names.append(filename)
    result_text_word.insert(pointer,("Docx Files\t:\t"+(f'{len(docx_file_names)}')+"\n\n"))
    pointer=float(result_text_word.index(INSERT))
    pointer +=1.0


    result_file_names=[]
    for filename in docx_file_names:
        text=docxpy.process(path+'/'+filename)
        if(word in text):
            result_file_names.append(filename)
    result_text_word.insert(pointer,("Matched Files\t:\t"+(f'{len(result_file_names)}')+"\n\n"))
    pointer=float(result_text_word.index(INSERT))
    pointer +=1.0
    for file in result_file_names:
         result_text_word.insert(pointer,(f'\t{file}\n'))
    

def  search_by_word(wfrm):
    wfrm.destroy()
    frm=Frame(win,bg='cyan')
    frm.place(relwidth=1,relheight=1)

    lbl_title=Label(win,text='Find The Files That Are Needed.',font=('',40,'bold'),bg='cyan',fg='red')
    lbl_title.place(x=440,y=20)

    lbl_directory=Label(frm,text='Select Directory : ',font=('',20,'bold'),bg='cyan')
    lbl_directory.place(relx=.1,rely=.2)
    entry_dict_name=Entry(frm, textvariable=txtFrame4,state="readonly",font=('',15,'bold'),bd=5)
    entry_dict_name.place(relx=.27,rely=.2)
    entry_dict_name.focus()

    btn_browse_file=Button(frm,command=lambda:browse_directory(),text='Browse',font=('',15,'bold'),bd=5)
    btn_browse_file.place(relx=.36,rely=.3)
    
    lbl_word=Label(frm,text='Enter Keyword : ',font=('',20,'bold'),bg='cyan')
    lbl_word.place(relx=.1,rely=.5)
    entry_word=Entry(frm, textvariable=txtFrame3,font=('',15,'bold'),bd=5)
    entry_word.place(relx=.27,rely=.5)
    entry_word.focus()

    lbl_word_result=Label(frm,text='The Result',font=('',20,'bold'),bg='cyan',fg='dark red')
    lbl_word_result.place(relx=.70,rely=.15)

    global result_text_word
    result_text_word=Text(frm,font=('','20','bold'),height=20,width=50)
    scrollbar=Scrollbar(frm, orient=VERTICAL, command=result_text_word.yview)
    result_text_word['yscroll']=scrollbar.set
    scrollbar.pack(side=RIGHT,fill=Y)
    result_text_word.place(relx=.50,rely=.2)
    scrollbar.place(in_=result_text_word,relx=0.98,relheight=1.0,bordermode='inside')


    btn_find=Button(frm,command=lambda:find_files(),text='find files',font=('',15,'bold'),bd=5)
    btn_find.place(relx=.25,rely=.7)
    

    btn_reset=Button(frm,command=lambda:reset_find(entry_word),text='Reset',font=('',15,'bold'),bd=5)
    btn_reset.place(relx=.36,rely=.7) 
def reset_find(e1):
    txtFrame4.set('')
    e1.delete(0,END)
    result_text_word.delete(1.0,10000000.0)

def main_screen(wfrm):
    wfrm.destroy()
    frm=Frame(win,bg='cyan')
    frm.place(relwidth=1,relheight=1)

    lbl_title=Label(win,text='Let\'s count some words ',font=('',40,'bold'),bg='cyan',fg='red')
    lbl_title.place(x=480,y=20)
    
    lbl_word_1=Label(frm,text='Select File : ',font=('',20,'bold'),bg='cyan')
    lbl_word_1.place(relx=.1,rely=.2)
    entry_file_name=Entry(frm, textvariable=txtFrame1,state="readonly",font=('',15,'bold'),bd=5)
    entry_file_name.place(relx=.27,rely=.2)
    entry_file_name.focus()

    btn_browse_file=Button(frm,command=lambda:browse_files(),text='Browse',font=('',15,'bold'),bd=5)
    btn_browse_file.place(relx=.36,rely=.3)


    lbl_word_2=Label(frm,text='To count any specific word from file: ',font=('',20,'bold'),bg='cyan')
    lbl_word_2.place(relx=.1,rely=.45)
    lbl_word_2=Label(frm,text='enter the word: ',font=('',20,'bold'),bg='cyan')
    lbl_word_2.place(relx=.1,rely=.55)
    entry_word=Entry(frm, textvariable=txtFrame2,font=('',15,'bold'),bd=5)
    entry_word.place(relx=.27,rely=.55)
    

    

    btn_count=Button(frm,command=lambda:count(),text='Count word',font=('',15,'bold'),bd=5)
    btn_count.place(relx=.25,rely=.7)
    

    btn_reset=Button(frm,command=lambda:reset(entry_file_name,entry_word),text='Reset',font=('',15,'bold'),bd=5)
    btn_reset.place(relx=.36,rely=.7)

    btn_reset=Button(frm,command=lambda:reset_result_text(),text='Reset Result box',font=('',10,'bold'),bd=5)
    btn_reset.place(relx=.42,rely=.9) 

    lbl_word_2=Label(frm,text='The Result',font=('',20,'bold'),bg='cyan',fg='dark red')
    lbl_word_2.place(relx=.70,rely=.15)

    global result_text
    result_text=Text(frm,font=('','20','bold'),height=20,width=50)
    scrollbar=Scrollbar(frm, orient=VERTICAL, command=result_text.yview)
    result_text['yscroll']=scrollbar.set
    scrollbar.pack(side=RIGHT,fill=Y)
    result_text.place(relx=.50,rely=.2)
    scrollbar.place(in_=result_text,relx=0.98,relheight=1.0,bordermode='inside')

    btn_search_word=Button(frm,command=lambda:search_by_word(frm),text='click to enter search by word program',font=('',15,'bold'),bd=5)
    btn_search_word.place(relx=.1,rely=.9)


def reset_result_text():
    result_text.delete(1.0,10000000.0)

    
def homescreen():
    frm=Frame(win,bg='cyan')
    frm.place(x=0,y=170,relwidth=1,relheight=1)

    lbl_title_1=Label(win,text='count the word of the given word file and if the \nfile is not of appropriate formate then proceed as commanded.',font=('',20,''),bg='cyan',fg='dark red',justify='center')
    lbl_title_1.place(x=430,y=300)
    
    btn_login=Button(frm,command=lambda:main_screen(frm),text='Let\'s count the Words ',font=('',15,'bold'),bd=5,height=2)
    btn_login.place(x=670,y=350)

homescreen()
win.mainloop()
