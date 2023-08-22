import tkinter as tk
from tkinter import scrolledtext
import logging
import threading
import time

# 定义一个用于重定向标准输出的辅助类
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)  # 滚动到文本末尾

def worker():
    for i in range(10):
        logger.info(f"Processing step {i}")
        time.sleep(1)

def start_processing():
    # 启动一个工作线程，模拟数据处理过程
    threading.Thread(target=worker, daemon=True).start()

# 创建窗口
window = tk.Tk()
window.title("实时日志输出示例")

# 创建文本控件用于显示日志
log_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=40, height=10)
log_text.pack()

# 创建日志记录器并设置级别
logger = logging.getLogger('RealTimeLog')
logger.setLevel(logging.DEBUG)

# 创建日志处理器并将其添加到记录器
text_handler = TextHandler(log_text)
logger.addHandler(text_handler)

# 创建按钮
start_button = tk.Button(window, text="开始处理", command=start_processing)
start_button.pack(pady=10)

# 进入主循环
window.mainloop()
