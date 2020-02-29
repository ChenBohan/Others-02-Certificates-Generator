from pptx import Presentation
from os import rename
import sys
from tkinter import filedialog, messagebox, Tk, Label, Button, StringVar, mainloop

txt_path = ''
ppt_path = ''
ppt_filename = 'certificates'


sys.stdout = open("./log.txt", "w")
print ("test sys.stdout")

def browse_ppt_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global show_ppt_path
    global ppt_path
    ppt_path = filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("txt files","*.pptx"),("all files","*.*")))
    show_ppt_path.set(ppt_path)
    print(ppt_path)

def browse_txt_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global show_txt_path
    global txt_path
    txt_path = filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("ppt files","*.txt"),("all files","*.*")))
    show_txt_path.set(txt_path)
    print(txt_path)

def browse_rename_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global show_rename_path
    global txt_path
    global ppt_filename
    rename_path = filedialog.askdirectory()
    show_txt_path.set(rename_path)
    print(rename_path)

    with open(txt_path, 'r') as f:
        name_list = f.readlines()

    for i in range(len(txt_path) - 1):
        if i <= 8:
            os.rename("{}/{}_0{}.png".format(rename_path, ppt_filename, i+1), "{}.png".format(name_list[i].strip()))
        else:
            os.rename("{}/{}_{}.png".format(rename_path, ppt_filename, i+1), "{}.png".format(name_list[i].strip()))

    messagebox.showinfo("提示", "重命名完成")

def browse_export_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global ppt_path
    global txt_path
    global ppt_filename

    with open(txt_path, 'r') as f:
        name_list = f.readlines()
    
    prs = Presentation(ppt_path)
    i = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    print(run.text)
                    if run.text == 'xxx':
                        run.text = name_list[i]
                        print(name_list[i])
                        i = i + 1

        prs.save('{}.pptx'.format(ppt_filename))
    messagebox.showinfo("提示", "生成成功")

row = 0
root = Tk()

lbl10 = Label(master=root, text="1.准备名单（每行一个名字，无标点无空格）：")
lbl10.grid(row=row, column=0)
row = row + 1

show_txt_path = StringVar()
lbl1 = Label(master=root, text="2.选择名单路径：")
lbl1.grid(row=row, column=0)
lbl2 = Label(master=root, textvariable=show_txt_path)
lbl2.grid(row=row, column=1)
button2 = Button(text="选择名单", command=browse_txt_button)
button2.grid(row=row, column=2)
row = row + 1

lbl11 = Label(master=root, text="3.准备.pptx模板（名字预填'xxx'，先复制好足够的页数）：")
lbl11.grid(row=row, column=0)
row = row + 1

show_ppt_path = StringVar()
lbl3 = Label(master=root, text="4.选择.pptx路径：")
lbl3.grid(row=row, column=0)
lbl4 = Label(master=root, textvariable=show_ppt_path)
lbl4.grid(row=row, column=1)
button3 = Button(text="选择PPT", command=browse_ppt_button)
button3.grid(row=row, column=2)
row = row + 1

export_path = StringVar()
lbl5 = Label(master=root, text="5.生成PPT")
lbl5.grid(row=row, column=0)
button4 = Button(text="一键生成", command=browse_export_button)
button4.grid(row=row, column=2)
row = row + 1

lbl12 = Label(master=root, text="6.手动逐页导出为图片，图片文件名应为'PPT名_序号'")
lbl12.grid(row=row, column=0)
row = row + 1

show_rename_path = StringVar()
lbl6 = Label(master=root, text="7.选择导出图片所在的文件夹目录")
lbl6.grid(row=row, column=0)
lbl7 = Label(master=root, textvariable=show_rename_path)
lbl7.grid(row=row, column=1)
button5 = Button(text="一键改名", command=browse_rename_button)
button5.grid(row=row, column=2)
row = row + 1

mainloop()



