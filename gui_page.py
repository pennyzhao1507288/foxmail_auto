import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import os
import subprocess


def browse_excel():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_path_entry.delete(0, tk.END)
    excel_path_entry.insert(0, excel_file_path)

def browse_app():
    app_path = filedialog.askopenfilename(filetypes=[("Executable Files", "*.exe")])
    app_path_entry.delete(0, tk.END)
    app_path_entry.insert(0, app_path)

def run_script():
    excel_path = excel_path_entry.get()
    app_path = app_path_entry.get()
    shortcut_key = shortcut_key_entry.get()  # 获取客户输入的快捷键
    search_box_type = search_box_var.get()

    script_path = "main.py"  # 替换为您的自动化脚本路径

    # 构造命令行
    command = ["python", script_path, excel_path, app_path, shortcut_key,search_box_type]  # 将快捷键传递给自动化脚本

    # 执行命令行
    subprocess.run(command)

    # 运行完成后弹出窗口
    messagebox.showinfo("运行完成", "程序已运行完成！", parent=root)  # 将 parent 参数设置为 root

# 创建主窗口
root = tk.Tk()
root.title("自动化脚本")
root.geometry("400x300")  # 设置窗口大小为 400x300

# 设置背景颜色为天空蓝
root.configure(bg="#6ED2EC")

# Excel 文件路径输入框和浏览按钮
excel_frame = tk.Frame(root, bg="#6ED2EC")  # 创建一个框架，设置背景颜色
excel_frame.pack(pady=20)  # 设置框架上下边距
excel_path_label = tk.Label(excel_frame, text="Excel 文件路径:", bg="#6ED2EC")
excel_path_label.pack(side=tk.LEFT)
excel_path_entry = tk.Entry(excel_frame)
excel_path_entry.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=10)  # 增加水平填充和边距
browse_button = tk.Button(excel_frame, text="浏览", command=browse_excel, bg="#3883BE", fg="white", relief=tk.GROOVE)  # 设置按钮背景颜色、前景颜色和边框样式
browse_button.pack(side=tk.LEFT)

# Foxmail 路径输入框和浏览按钮
app_frame = tk.Frame(root, bg="#6ED2EC")  # 创建一个框架，设置背景颜色
app_frame.pack(pady=20)  # 设置框架上下边距
app_path_label = tk.Label(app_frame, text="Foxmail 路径:", bg="#6ED2EC")
app_path_label.pack(side=tk.LEFT)
app_path_entry = tk.Entry(app_frame)
app_path_entry.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=10)  # 增加水平填充和边距
browse_app_button = tk.Button(app_frame, text="浏览", command=browse_app, bg="#3883BE", fg="white", relief=tk.GROOVE)  # 设置按钮背景颜色、前景颜色和边框样式
browse_app_button.pack(side=tk.LEFT)

# 快捷键输入框
shortcut_frame = tk.Frame(root, bg="#87CEEB")
shortcut_frame.pack(pady=20)
shortcut_key_label = tk.Label(shortcut_frame, text="自定义快捷键:", bg="#87CEEB")
shortcut_key_label.pack(side=tk.LEFT)
shortcut_key_entry = tk.Entry(shortcut_frame)
shortcut_key_entry.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=10)

# 搜索框类型选择
search_box_frame = tk.Frame(root, bg="#87CEEB")
search_box_frame.pack(pady=20)
search_box_label = tk.Label(search_box_frame, text="搜索框类型:", bg="#87CEEB")
search_box_label.pack(side=tk.LEFT)
search_box_var = tk.StringVar()
search_box_var.set("主题搜索框")  # 默认选择主题搜索框
search_box_menu = tk.OptionMenu(search_box_frame, search_box_var, "主题搜索框", "全文搜索框")
search_box_menu.pack(side=tk.LEFT)

# 运行按钮，设置为椭圆形状
run_button = tk.Button(root, text="运行脚本", command=run_script, bg="#3883BE", fg="white", relief=tk.GROOVE, bd=0, padx=20, pady=10, borderwidth=0, highlightthickness=0, activebackground="#3883BE")  # 设置按钮样式
run_button.config(width=10, height=2)
run_button.pack()

# 启动主循环
root.mainloop()
