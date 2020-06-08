# -*- coding: utf-8 -*-

from configparser import ConfigParser

import tkinter as tk
import tkfilebrowser
from tkinter import filedialog, ttk

cfg = ConfigParser(allow_no_value=True)
cfg.read('config.ini', encoding='utf-8')

window = tk.Tk()
window.title('导入图片至表格')
window.geometry('800x500')


def update_config_file(cfg_key, cfg_value, config_section='DEFAULT'):
    """更新配置文件"""
    cfg[config_section][cfg_key] = cfg_value

    with open('config.ini', 'w', encoding='utf-8') as configfile:
        cfg.write(configfile)


# 定义excel模板文件选择部件
def select_file():
    file_path_selected = filedialog.askopenfilename()
    file_path.set(file_path_selected)
    print('选中文件：', file_path_selected, type(file_path_selected))
    update_config_file('excel_file_name', file_path_selected)


file_path = tk.StringVar()
tk.Label(window, text='Excel模板文件路径:', width=15, height=2).grid(row=0, column=0)
e = tk.Entry(window, textvariable=file_path, width=80)
e.grid(row=0, column=1)
tk.Button(window, text='选择文件', command=select_file, width=10).grid(row=0, column=2)


# 定义照片文件夹选择部件
dir_path_selected = []


def select_dir_path():
    dir_path_selected.append(tkfilebrowser.askopendirnames())
    dir_path.set(dir_path_selected)
    print('选中文件夹：', dir_path_selected, type(dir_path_selected))
    # 选中文件夹为一个列表(列表元素为每次选择的元组)，需要将其转换为字符串，再更新到配置文件
    update_config_file('img_dir_name_list', str(dir_path_selected))


dir_path = tk.StringVar()
tk.Label(window, text='图片文件夹路径:', width=15, height=2).grid(row=2, column=0)
lb = tk.Listbox(window, listvariable=dir_path, width=80, height=3)
lb.grid(row=2, column=1)
tk.Button(window, text='选择文件夹', command=select_dir_path, width=10).grid(row=2, column=2)


window.mainloop()
