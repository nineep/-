# -*- coding: utf-8 -*-

from configparser import ConfigParser

import tkfilebrowser
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk

from insert_images import dirs_images_insert_excels


cfg = ConfigParser(allow_no_value=True)
cfg.read('config.ini', encoding='utf-8')

img_root_dir = cfg['DEFAULT']['img_root_dir']
img_dir_name_list = cfg['DEFAULT']['img_dir_name_list'].split()
excel_file_name = cfg['DEFAULT']['excel_file_name']
sheet_name = cfg['DEFAULT']['sheet_name']
label_template = cfg['DEFAULT']['label_template'].split()

# 开启窗口
window = tk.Tk()
window.title('imageXexcel')
window.geometry('800x500')

# 定义几个字体格式
ft = tkfont.Font(family='Arial', size=10, weight=tkfont.BOLD)
ft1 = tkfont.Font(size=20, slant=tkfont.ITALIC)
ft2 = tkfont.Font(size=30, weight=tkfont.BOLD, underline=1, overstrike=1)


def update_config_file(cfg_key, cfg_value, config_section='DEFAULT'):
    """更新配置文件"""
    cfg[config_section][cfg_key] = cfg_value

    with open('config.ini', 'w', encoding='utf-8') as configfile:
        cfg.write(configfile)


def list_to_str(ls):
    """
    将元素为tuple的list，转换为元组元素以空格隔开的一个字符串
    """
    new_tup = ()
    for i in range(len(ls)):
        tup = ls[i]
        # print(tup)
        new_tup += tup
    # print(new_tup, type(new_tup))
    new_set = set(new_tup)
    # print(new_set, type(new_set))

    new_str = ''
    for s in new_set:
        ss = s + ' '
        new_str += ss
    # print(new_str)
    return new_str


# 定义GUI界面

# 定义 LabelFrame1 部件
lf1 = ttk.LabelFrame(window, text='输入文件：', height=20, width=100)
lf1.grid(row=0, column=0, padx=1, pady=15)


# 定义excel模板文件选择部件
def select_file():
    file_path_selected = filedialog.askopenfilename()
    file_path.set(file_path_selected)
    print('选中文件：', file_path_selected, type(file_path_selected))
    update_config_file('excel_file_name', file_path_selected)


file_path = tk.StringVar()
tk.Label(window, text='Excel模板文件路径:', width=15, height=2).grid(row=3, column=0)
e = tk.Entry(window, textvariable=file_path, width=80)
e.grid(row=3, column=1)
tk.Button(window, text='选择文件', command=select_file, width=10, foreground='green').grid(row=3, column=2)


# 定义excel模板文件sheet name选择部件
def input_worksheet_name():
    ws_name = worksheet_name.get()
    # worksheet_name.set(ws_name)
    print('工作表名：', ws_name, type(ws_name))
    update_config_file('sheet_name', ws_name)


worksheet_name = tk.StringVar()
worksheet_name.set('勘察照片')

tk.Label(window, text='Excel模板工作表名:', width=15, height=2).grid(row=4, column=0)
e = tk.Entry(window, textvariable=worksheet_name, width=80)
e.grid(row=4, column=1)
tk.Button(window, text='确认工作表名', command=input_worksheet_name, width=10, foreground='green').grid(row=4, column=2)


# 定义照片文件夹选择部件
dir_path_selected = []


def select_dir_path():
    dir_path_selected.append(tkfilebrowser.askopendirnames())
    dir_path.set(dir_path_selected)
    print('选中文件夹：', dir_path_selected, type(dir_path_selected))
    # 选中文件夹为一个列表(列表元素为每次选择的元组)，需要将其转换为字符串，再更新到配置文件
    transform_dir_path_selected = list_to_str(dir_path_selected)
    update_config_file('img_dir_name_list', transform_dir_path_selected)


dir_path = tk.StringVar()
tk.Label(window, text='图片文件夹路径:', width=15, height=2).grid(row=1, column=0)
lb = tk.Listbox(window, listvariable=dir_path, width=80, height=3)
lb.grid(row=1, column=1)
tk.Button(window, text='选择文件夹', command=select_dir_path, width=10, foreground='green').grid(row=1, column=2)


# 定义运行脚本部件
def run():
    print('执行图片插入excel脚本。')
    dirs_images_insert_excels(img_dir_name_list, label_template)


run_button = tk.Button(window, text='开始运行', command=run, width=10, foreground='red', background='pink')
run_button.grid(row=5, column=2, padx=0, pady=15)


# 定义 LabelFrame2 部件
lf2 = ttk.LabelFrame(window, text='输出文件：', height=20, width=100)
lf2.grid(row=6, column=0, padx=1, pady=10)






window.mainloop()
