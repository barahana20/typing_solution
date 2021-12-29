"""모든 수식 텍스트 차례로 dict로 얻기.
키는 (List, Para, Pos), 값은 eqn_string"""
from tkinter import *
from tkinter import messagebox
import tkinter as tk
import re
from typing import Counter
import win32com.client as win32
import pyperclip as cb
from win32com.client.makepy import GenerateFromTypeLibSpec
import os
import glob
import time

global hwp
global eqn_dict
global next_error_count
global error_find_log
global error_save_log
global count
count = 0
error_save_log = []
next_error_count = 0
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)
def make_log_file():  #  로그파일 생성
    
    now = time.localtime()

    now_time = "%04d%02d%02d_%02d%02d%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
    parent_path = os.path.dirname(os.path.realpath(__file__))
    createFolder(parent_path+f'\log')
    f = open(parent_path+f'\log\{now_time}_log.txt', 'w')
    f.close()
    return parent_path+f'\log\{now_time}_log.txt'

def extract_eqn(hwp):  # 이전 포스팅에서 소개한, 수식 추출방법을 함수로 정의
    Act = hwp.CreateAction("EquationModify")
    Set = Act.CreateSet()
    Pset = Set.CreateItemSet("EqEdit", "EqEdit")
    Act.GetDefault(Pset)
    return Pset.Item("String")


def select_error(hwp, key, value, comment):  # error를 tkinter에 출력
    global error_find_log
    error_find_log = False
    return comment
    # print(f"{count}번째\nposition: {key}, expression: {value} \n*Error_comment: {comment}\n") # 에러 발견 count, 좌표 key, 전체 수식 value, 무엇이 잘못됐는지 comment 출력


def return_hwp_files():  # hwp 파일들 조사 후 glob data 반환
    hwp_names = []
    parent_path = os.path.dirname(os.path.realpath(__file__))
    data = glob.glob(parent_path+'\*')
    for i in data:
        if i.find('.hwp')!=-1:
            hwp_names.append(i)
    return hwp_names


def Open_hwp(hwp_name):  # hwp 파일 열기
    global hwp
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.Open(hwp_name)
    hwp.XHwpWindows.Item(0).Visible = True


def error_find(key, value):  # 이상한 점 찾는 함수
    global error_find_log
    error_find_log = True
    const_value = value
    if(const_value.find('lim _{')!=-1 and error_find_log):
        if const_value.find('`->`')!=-1 or const_value.find('` rarrow `')!=-1:
            pass
        elif const_value.find('`->')!=-1 and const_value.find('->`')==-1:
            comment = select_error(hwp, key, const_value, "->' 얖에 약한 공백을 넣어주세요.")
            
        elif const_value.find('->`')!=-1:
            comment = select_error(hwp, key, const_value, "->' 앞에 약한 공백을 넣어주세요.")
            
        elif const_value.find('->')!=-1:
            comment = select_error(hwp, key, const_value, "->' 얖 옆에 약한 공백을 넣어주세요.")
            
        elif const_value.find('` rarrow ')!=-1 and const_value.find(' rarrow `')==-1:
            comment = select_error(hwp, key, const_value, "'rarrow(->)' 얖에 약한 공백을 넣어주세요.")
            
        elif const_value.find(' rarrow `')!=-1:
            comment = select_error(hwp, key, const_value, "'rarrow(->)' 앞에 약한 공백을 넣어주세요.")
            
        elif const_value.find(' rarrow ')!=-1:
            comment = select_error(hwp, key, const_value, "'rarrow(->)' 얖 옆에 약한 공백을 넣어주세요.")
    #x`vert `x`=` root {b} of {a} ,`a IN A,`b IN B
    if const_value.find('rm')==-1 and re.findall(r'[a-zA-Z]{2,}', const_value)==[] and error_find_log: # rm이 아닌 it대문자 뒤에 `이 들어가 있지 않으면 경고하기 위한 조건문
        index_bak = 0
        value = const_value
        for i in const_value:
            if(i.isupper() and i!='A'): # 대문자를 발견하면
                if(value.index(i)+1==len(value)): # 만약 i가 문장 마지막 문자라면
                    comment = select_error(hwp, key, const_value, f"{value.index(i)+index_bak}번 째({i}) 뒤에 약한 공백(`)을 넣어주세요.")
                    
                elif(value[value.index(i)+1]!='`'):
                    comment = select_error(hwp, key, const_value, f"{value.index(i)+index_bak}번 째({i}) 뒤에 약한 공백(`)을 넣어주세요.")
                    
            value = value[value.index(i)+1:]
            index_bak+=1
    if(const_value.find('log _{')!=-1 and error_find_log):
        if(const_value.count('log _{')==1):
            if const_value[const_value.index('log _{'):const_value.index('}')].find('`')==-1:
                comment = select_error(hwp, key, const_value, "'log 밑' 앞 에 약한 공백을 넣어주세요.")
                
            value = const_value
            if(const_value[const_value.index('log _{')+len('log _{'):const_value.index('}')].find('{')!=-1): # 'log 밑'에 sqrt같은 명령어가 오면 뒤에 {}붙음(ex> sqrt {2}). 이거 건너뛰기 위한 코드.
                value = value[value.index('}')+len('}'):]
            if value[value.index('}')+1]!='`' and value[value.index('}')+2]!='`' and value[value.index('}')+3]!='`':
                comment = select_error(hwp, key, const_value, "'log 지수' 앞 에 약한 공백을 넣어주세요.")
                
        elif(const_value.count('log _{')>1):
            #4 ^{log _{`2} `x} BULLET 2 ^{log _{` sqrt {2}} `y}
            value = const_value # value 초기화
            while(1): # value에 'log _{'가 없어질 때까지 반복
                value_index = value.index('log _{')
                if value[value_index+len('log _{'):value.index('}')].find('`')==-1:
                    comment = select_error(hwp, key, const_value, f"'log 밑'(위치: {value_index}번 째 log) 앞 에 약한 공백을 넣어주세요.")
                
                if(value[value_index+len('log _{'):value.index('}')].find('{')!=-1): # 'log 밑'에 sqrt같은 명령어가 오면 뒤에 {}붙음(ex> sqrt {2}). 이거 건너뛰기 위한 코드.
                    value = value[value.index('}')+len('}'):]
                else:
                    value = value[value.index('}'):]
                if(len(value)-1==value.index('}')):
                    break
                elif value[value.index('}')+1]!='`' and value[value.index('}')+2]!='`' and value[value.index('}')+3]!='`':
                    comment = select_error(hwp, key, const_value, f"'log 지수'(위치: {value_index}번 째 log) 앞 에 약한 공백을 넣어주세요.")
                    #log _{`6} `a _{`1} +log _{`6} `a _{`2} +log _{`6} a _{`3} + BULLET  BULLET  BULLET  +log _{`6} `a _{`12} 에러 발생
                if value.count('log _{')==0:
                    break
                else:
                    value = value[value.index('log _{'):]
    elif(const_value.find('log')!=-1 and error_find_log):
        if(const_value.count('log')==1):
            if(const_value[const_value.index('log')+len('log')].isdigit()):
                comment = select_error(hwp, key, const_value, f"'log'({const_value.index('log')}번 째 log) 뒤 에 약한 공백을 넣어주세요.")
        elif(const_value.count('log')>1):
            value = const_value
            first_sentence_len = 0
            while(1):
                if(value[value.index('log')+len('log')].isdigit()):
                    comment = select_error(hwp, key, const_value, f"'log'({len(value[:value.index('log')])+len('log')+first_sentence_len}번 째 log) 뒤 에 약한 공백을 넣어주세요.")
                first_sentence_len += len(value[:value.index('log')])
                value = value[value.index('log')+len('log'):]
                if value.count('log')==0:
                    break
    if(re.findall(r"\([0-9]+,`[0-9]+\)", const_value)!=[] or re.findall(r"\(-[0-9]+,`[0-9]+\)", const_value)!=[] or re.findall(r"\([0-9]+,`-[0-9]+\)", const_value)!=[] or re.findall(r"\(-[0-9]+,`-[0-9]+\)", const_value)!=[] and error_find_log): # 만약 range 안에 약한 공백 1개밖에 없다면
        comment = select_error(hwp, key, const_value, f"',`' 뒤에 약한 공백(`)을 하나 더 넣어주세요.")
    elif(re.findall(r"\([0-9]+,[0-9]+\)", const_value)!=[] or re.findall(r"\(-[0-9]+,[0-9]+\)", const_value)!=[] or re.findall(r"\([0-9]+,-[0-9]+\)", const_value)!=[] or re.findall(r"\(-[0-9]+,-[0-9]+\)", const_value)!=[] and error_find_log): # 만약 약한 공백이 아예 없다면
            comment = select_error(hwp, key, const_value, f"',' 뒤에 약한 공백(`)을 두 개 넣어주세요.")
    oper_dic = {'mul' : 'TIMES', 'div' : '/', 'per' : '%', 'plus' : '+', 'minus' : '-', 'greater' : '>', 'less' : '<', 'greater_equal' : '>:',\
    'less_equal' : '<:', 'equal' : '='}
    if(error_find_log):
        for oper_name, sign in oper_dic.items():
            if(const_value.find(sign)!=-1):
                if(const_value.count(sign)==1):
                    judge_last = const_value.index(sign)+len(sign)
                    if(judge_last == len(const_value)): # 만약 i가 문장 마지막 문자라면 
                        comment = select_error(hwp, key, const_value, f"{len(const_value[:const_value.index(sign)])+len(sign)}번 째('{sign}') 뒤에 약한 공백(`)을 넣어주세요.")
                            
                elif(const_value.count(sign)>1):
                    value = const_value
                    first_sentence_len = -1
                    while(1):
                        judge_last = value.index(sign)+len(sign)
                        try:
                            if(judge_last == len(value)): # 만약 i가 문장 마지막 문자라면 
                                comment = select_error(hwp, key, const_value, f"{len(value[:value.index(sign)])+len(sign)+first_sentence_len}번 째('{sign}') 뒤에 약한 공백(`)을 넣어주세요.")
                            elif(value[value.index(sign)+len(sign)]!='`' and value[value.index(sign)+len(sign)+1]!='`'):
                                comment = select_error(hwp, key, const_value,  f"{len(value[:value.index(sign)])+len(sign)+first_sentence_len}번 째('{sign}') 뒤에 약한 공백(`)을 넣어주세요.")
                        except:
                            if(judge_last == len(value)): # 만약 i가 문장 마지막 문자라면 
                                comment = select_error(hwp, key, const_value, f"{len(value[:value.index(sign)])+len(sign)+first_sentence_len}번 째('{sign}') 뒤에 약한 공백(`)을 넣어주세요.")
                            elif(value[value.index(sign)+len(sign)]!='`'):
                                comment = select_error(hwp, key, const_value,  f"{len(value[:value.index(sign)])+len(sign)+first_sentence_len}번 째('{sign}') 뒤에 약한 공백(`)을 넣어주세요.")
                        
                        first_sentence_len += len(value[:value.index(sign)])
                        value = value[value.index(sign)+len(sign):]
                        if value.count(sign)==0:
                            break
    if(error_find_log==False):
        return False, comment
    else:
        return True, "True"

def adventure_hwp():  # 모든 수식의 좌표와 값을 딕셔너리 eqn_dict에 저장
    global eqn_dict
    eqn_dict = {}  # 사전 형식의 자료 생성 예정
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = position.Item("List"), position.Item("Para"), position.Item("Pos")
            hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            eqn_string = extract_eqn(hwp)  # 문자열 추출
            eqn_dict[position] = eqn_string  # 좌표가 key이고, 수식문자열이 value인 사전 생성
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.Run("Cancel")  # 완료했으면 선택해제


def tk_start_work(hwp_names):  # hwp 파일 조사 후 hwp 수식 조사 후 저장
    global next_error_count
    # for hwp_name in hwp_names: # 여러 파일들을 한꺼번에 볼 수 있도록 하려 하였으나 귀찮아서 안함.
    hwp_name = hwp_names[0]
    next_error_count = 0
    hwp_name_entry_value.set(f'{os.path.basename(hwp_name)}')
    Open_hwp(hwp_name)
    adventure_hwp()


def the_end():  # 마지막을 알리는 함수
    global f_add
    expression_entry_value.set('끝')
    fix_entry_value.set('끝')
    f_add.close()


def tk_next_error_find():  # 다음 에러 찾는 함수
    global next_error_count
    global error_save_log
    global count
    global f_add
    
    while(1):
        if(next_error_count>=len(eqn_dict)):
            the_end()
            break
        key, value = list(eqn_dict.items())[next_error_count][0], list(eqn_dict.items())[next_error_count][1]
        error_find_result = error_find(key, value)
        if(error_find_result[0]==False):
            comment = error_find_result[1]
            count+=1
            print(count)
            hwp.SetPos(*key)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            count_entry_value.set(f"{count}번 째")
            expression_entry_value.set(value)
            fix_entry_value.set(comment)
            f_add.write(f'{count}번 째\nposition: {key}, expression: {value} \n*Error_comment: {comment}\n\n')
            error_save_log.append((count, next_error_count))
            next_error_count+=1
            return
        next_error_count+=1

def tk_before_error_find():  # 이전 에러 찾는 함수
    global next_error_count
    global error_save_log
    global count
    error_save_log.pop()
    error_save_log_pop = error_save_log.pop()
    count -= 2
    next_error_count = error_save_log_pop[1]
    tk_next_error_find()
if __name__ == "__main__":
    log_path = make_log_file()
    global f_add
    f_add = open(log_path, 'a')

    root = Tk()
    root.title("hwp 수식 문법 검사기")
    start_comment_label = Label(root, text="    시작하려면 버튼을 눌러 주세요.")
    start_comment_label.grid(row=0, column=0)
    start_btn = Button(root, text="START", width=5, command= lambda : tk_start_work(return_hwp_files()))
    # btn.grid(row=1, column=2)
    start_btn.grid(row = 0, column = 1, ipadx=25, ipady=15)
    start_btn_explain_label = Label(root, text="버튼을 누른 후 뜨는 팝업창에서 접근 허용 또는 모두 허용을 클릭")
    start_btn_explain_label.grid(row=0, column=2)

    hwp_name_label = Label(root, text="    현재 작업중인 hwp 파일 이름")
    hwp_name_label.grid(row=1, column=0)
    hwp_name_entry_value = StringVar()
    hwp_name_entry = Entry(root,textvariable=hwp_name_entry_value)
    hwp_name_entry.grid(row=1,column=1,padx=100,pady=1,ipadx=80,ipady=1)

    btn_explain_label = Label(root, text="다음 에러를 보려면 'next' 버튼을 눌러주세요.")
    btn_explain_label.grid(row=1, column=2)
    btn_explain_label1 = Label(root, text="이전 에러를 보려면 'before' 버튼을 눌러주세요.")
    btn_explain_label1.grid(row=2, column=2)


    before_btn = Button(root, text="BEFORE", width=5, command= lambda : tk_before_error_find(), bg = "white", fg = "blue")
    before_btn.grid(row=3,column=2,padx=100,pady=5,ipadx=80,ipady=30)

    next_btn = Button(root, text="NEXt", width=5, command= lambda : tk_next_error_find(), bg = "white", fg = "red")
    next_btn.grid(row=4,column=2,padx=100,pady=5,ipadx=80,ipady=30)


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


# for key, value in eqn_dict.items():
#     print(value)
# exit(1)





"""
검사하는 에러 목록
!=1 -> != `1
=1 -> =`1
- -> -`
+ -> +`
< -> <`
> -> >`
TIMES
%
/
"""

"""
x->INF
x rarrow INF
rm 없을때 A빼고 전부 뒤에 `붙이기 it붙은거 포함

log _{`a} `a ^{7} b ^{3}
log _{`2} `x+`log _{`4} `y ^{2} =`3

log _{`a} `x=`2
log`rootx
log`2=`0.3013


!=1 -> != `1
=1 -> =`1
- -> -`
+ -> +`
< -> <`
> -> >`
``vert `` 나중에 팔요하면 구현
(-1,``3)
"""