import tkinter as tk
from tkinter import filedialog
from tkinter import font
import data_processor
import locale
import os
import sys
from tkinter import messagebox

# 定义一个用于重定向标准输出的辅助类
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ''

    def write(self, message):
        self.buffer += message
        lines = self.buffer.split('\n')
        for line in lines[:-1]:
            self.text_widget.insert(tk.END, line + '\n')
            self.text_widget.see(tk.END)  # 自动滚动到最底部
        self.buffer = lines[-1]

    def flush(self):
        pass

# 设置locale为中文（简体）
locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')

# 创建窗口
window = tk.Tk()
window.title("Excel自动化工具")
# 设置字体
chinese_font = font.Font(family='微软雅黑', size=12)  # 选择合适的字体
window.option_add("*Font", chinese_font)

# 设置窗口大小和位置
window_width = 600
window_height = 350
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
window.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 创建选择结果标签
result_label = tk.Label(window, text="")
result_label.pack()

# 记录选择的文件路径
selected_file_path = ""

# 新增：创建文本控件用于显示日志
log_text = tk.Text(window, height=10, width=60)
log_text.pack(pady=10)

# 创建选择文件函数
def open_file_dialog():
    global selected_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_name = os.path.basename(file_path)  # 获取文件名
        result_label.config(text="当前选择的文件：" + file_name)
        selected_file_path = file_path
        return file_path
    else:
        result_label.config(text="")
        selected_file_path = ''
        return None

# 创建开始数据操作函数
def start_data_processing():
    global selected_file_path
    if not selected_file_path:
        messagebox.showinfo("提示", "请先选择一个文件")
        return
    
    # 清空日志文本控件
    log_text.delete('1.0', tk.END)

    # 重定向标准输出流到日志文本控件
    old_stdout = sys.stdout
    sys.stdout = TextRedirector(log_text)
    
    data_processor.process_excel(selected_file_path)
    result_label.config(text="数据操作完成")

    # 恢复标准输出流
    sys.stdout = old_stdout

# 创建按钮
open_button = tk.Button(window, text="选择Excel文件", command=open_file_dialog)
start_button = tk.Button(window, text="开始", command=start_data_processing)

# 布局
open_button.pack(pady=10)
start_button.pack(pady=10)

# 进入主循环
window.mainloop()
