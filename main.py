import tkinter
import tkinter.ttk as ttk
from tkinter import *
from tkinter import filedialog
import excel_trans
from functools import partial
import openpyxl


def btncmd():
    # progressbar.pack()
    files = []
    # 수정
    if len(txt_dest_path1.get()) != 0:
        files.append({"type": "쿠팡", "file": txt_dest_path1.get()})
    if len(txt_dest_path2.get()) != 0:
        files.append({"type": "11번가", "file": txt_dest_path2.get()})
    if len(txt_dest_path3.get()) != 0:
        files.append({"type": "위메프", "file": txt_dest_path3.get()})
    if len(txt_dest_path4.get()) != 0:
        files.append({"type": "네이버", "file": txt_dest_path4.get()})
    if len(txt_dest_path5.get()) != 0:
        files.append({"type": "티몬", "file": txt_dest_path5.get()})
    if len(txt_dest_path6.get()) != 0:
        files.append({"type": "롯데온", "file": txt_dest_path6.get()})
    if len(txt_dest_path7.get()) != 0:
        files.append({"type": "ESM+", "file": txt_dest_path7.get()})

    if len(files) == 0:
        tkinter.messagebox.showwarning("경고", "파일을 한 개 이상 추가하세요.")
        return
    if len(txt_dest_path.get()) == 0:
        tkinter.messagebox.showwarning("경고", "저장경로를 선택하세요.")
        return
    try:
        excel_trans.excel_trans_print(files, folder_selected, p_var, progress_bar, product_info_txt_dest_path.get())
        tkinter.messagebox.showinfo("알림", "변환 완료")
    except Exception as e:
        tkinter.messagebox.showerror("에러", e)


def open_file(type):
    print("open file : ", type)
    file = filedialog.askopenfilename(title=type + " 엑셀 파일을 선택하세요.",
                                      filetypes=(("excel 파일", ".xlsx .xls"), ("모든 파일", "*.*")))
                                      #initialdir=r"C:/Users")
    print("file : ", file)
    # 수정
    if type == "쿠팡":
        txt_dest_path1.delete(0, END)
        txt_dest_path1.insert(0, file)
    elif type == "11번가":
        txt_dest_path2.delete(0, END)
        txt_dest_path2.insert(0, file)
    elif type == "위메프":
        txt_dest_path3.delete(0, END)
        txt_dest_path3.insert(0, file)
    elif type == "네이버":
        txt_dest_path4.delete(0, END)
        txt_dest_path4.insert(0, file)
    elif type == "티몬":
        txt_dest_path5.delete(0, END)
        txt_dest_path5.insert(0, file)
    elif type == "롯데온":
        txt_dest_path6.delete(0, END)
        txt_dest_path6.insert(0, file)
    elif type == "ESM+":
        txt_dest_path7.delete(0, END)
        txt_dest_path7.insert(0, file)
    elif type == "상품정보":
        product_info_txt_dest_path.delete(0, END)
        product_info_txt_dest_path.insert(0, file)

def browse_dest_path():
    global  folder_selected
    folder_selected = filedialog.askdirectory()
    if folder_selected is None:
        return
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected)

root = Tk()
root.title("TOGGLE 변환 프로그램")

root.geometry("640x800")  # 가로 * 세로
root.resizable(False, False)

main_frame = Frame(root, height=20)
main_frame.pack(side=TOP, fill=BOTH, expand=YES, padx=5, pady=5)

# 쿠팡
path_frame1 = LabelFrame(main_frame, text="쿠팡")
path_frame1.pack(fill="both")
txt_dest_path1 = Entry(path_frame1)
txt_dest_path1.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path1 = Button(path_frame1, text="찾아보기", width=10, command=partial(open_file, "쿠팡"))
btn_dest_path1.pack(side="right", padx=5, pady=5)

# 11번가
path_frame2 = LabelFrame(main_frame, text="11번가")
path_frame2.pack(fill="both")
txt_dest_path2 = Entry(path_frame2)
txt_dest_path2.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path2 = Button(path_frame2, text="찾아보기", width=10, command=partial(open_file, "11번가"))
btn_dest_path2.pack(side="right", padx=5, pady=5)

# 위메프
path_frame3 = LabelFrame(main_frame, text="위메프")
path_frame3.pack(fill="both")
txt_dest_path3 = Entry(path_frame3)
txt_dest_path3.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path3 = Button(path_frame3, text="찾아보기", width=10, command=partial(open_file, "위메프"))
btn_dest_path3.pack(side="right", padx=5, pady=5)

# 네이버
path_frame4 = LabelFrame(main_frame, text="네이버")
path_frame4.pack(fill="both")
txt_dest_path4 = Entry(path_frame4)
txt_dest_path4.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path4 = Button(path_frame4, text="찾아보기", width=10, command=partial(open_file, "네이버"))
btn_dest_path4.pack(side="right", padx=5, pady=5)

# 티몬
path_frame5 = LabelFrame(main_frame, text="티몬")
path_frame5.pack(fill="both")
txt_dest_path5 = Entry(path_frame5)
txt_dest_path5.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path5 = Button(path_frame5, text="찾아보기", width=10, command=partial(open_file, "티몬"))
btn_dest_path5.pack(side="right", padx=5, pady=5)

# 롯데온
path_frame6 = LabelFrame(main_frame, text="롯데온")
path_frame6.pack(fill="both")
txt_dest_path6 = Entry(path_frame6)
txt_dest_path6.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path6 = Button(path_frame6, text="찾아보기", width=10, command=partial(open_file, "롯데온"))
btn_dest_path6.pack(side="right", padx=5, pady=5)

# ESM+
path_frame7 = LabelFrame(main_frame, text="ESM+")
path_frame7.pack(fill="both")
txt_dest_path7 = Entry(path_frame7)
txt_dest_path7.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)
btn_dest_path7 = Button(path_frame7, text="찾아보기", width=10, command=partial(open_file, "ESM+"))
btn_dest_path7.pack(side="right", padx=5, pady=5)

# 수정
product_info_path_frame = LabelFrame(root, text="상품파일경로")
product_info_path_frame.pack(fill="both", padx=5, pady=5)

product_info_txt_dest_path = Entry(product_info_path_frame)
product_info_txt_dest_path.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)

product_info_btn_dest_path = Button(product_info_path_frame, text="찾아보기", width=10, command=partial(open_file, "상품정보"))
product_info_btn_dest_path.pack(side="right", padx=5, pady=5)




path_frame = LabelFrame(root, text="저장경로")
path_frame.pack(fill="both", padx=5, pady=5)

txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)

btn_dest_path = Button(path_frame, text="찾아보기", width=10, command=browse_dest_path)
btn_dest_path.pack(side="right", padx=5, pady=5)

frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x", padx=5, pady=5)

p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x", padx=5, pady=5)

frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_conv = Button(frame_run, padx=5, pady=5, text="변환", width=12, command=btncmd)
btn_conv.pack(side="right", padx=5, pady=5)

root.mainloop()
