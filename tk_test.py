from tkinter import *
from tkinter import messagebox
import tkinter as tk
root = Tk()
root.title("hwp 수식 문법 검사기")
def pass_through(a):
    print(a)
start_comment_label = Label(root, text="시작하려면 버튼을 눌러 주세요.")
start_comment_label.grid(row=0, column=0)
start_btn = Button(root, text="START", width=5, command= lambda : pass_through('ab'))
# btn.grid(row=1, column=2)
start_btn.grid(row = 0, column = 1, ipadx=25, ipady=15)
start_btn_explain_label = Label(root, text="버튼을 누른 후 뜨는 팝업창에서 접근 허용 또는 모두 허용을 클릭")
start_btn_explain_label.grid(row=0, column=2)

hwp_name_label = Label(root, text="현재 작업중인 hwp 파일 이름")
hwp_name_label.grid(row=1, column=0)
hwp_name_entry = Entry(root)
hwp_name_entry.grid(row=1,column=1,padx=100,pady=1,ipadx=80,ipady=1)

next_btn = Button(root, text="NEXt", width=5, command=pass_through, bg = "white", fg = "red")
next_btn.grid(row=3,column=2,padx=100,pady=5,ipadx=80,ipady=30)

next_btn_explain_label = Label(root, text="다음 에러를 보려면 밑의 버튼을 눌러주세요.")
next_btn_explain_label.grid(row=2, column=2)

count_label = Label(root, text="count")
count_label.grid(row=2, column=0)

expression_label = Label(root, text="수식")
expression_label.grid(row=3, column=0)

fix_label = Label(root, text="고쳐야할 사항")
fix_label.grid(row=4, column=0)

count_entry_value = StringVar()
count_entry = Entry(root,textvariable=count_entry_value)
count_entry.grid(row=2,column=1,padx=1,pady=1,ipadx=1,ipady=1)

expression_entry_value = StringVar()
expression_entry = Entry(root,textvariable=expression_entry_value)
expression_entry.grid(row=3,column=1,padx=100,pady=5,ipadx=80,ipady=30)

fix_entry_value = StringVar()
fix_entry = Entry(root,textvariable=fix_entry_value)
fix_entry.grid(row=4,column=1,padx=100,pady=5,ipadx=80,ipady=30)


root.mainloop()