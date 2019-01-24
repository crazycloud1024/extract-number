# # @Time    : 2019/1/16 11:05
# # @Author  : luoqian
# # @Email   : vampire.luo@foxmail.com
# # @File    : 号码提取.py
# # @Software: PyCharm
#

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import re, datetime
import openpyxl

# regex = re.compile(
#     '(1(?:3\d{3}|5[^4\D]\d{2}|8\d{3}|7(?:[01356789]\d{2}|4(?:0\d|1[0-2]|9\d))|9[189]\d{2}|6[567]\d{2}|4(?:[14]0\d{3}|[68]\d{4}|[579]\d{2}))\d{6})')

regex = re.compile(
    """[0\-_]?(1(?:3\d{3}|5[^4\D]\d{2}|8\d{3}|7(?:[01356789]\d{2}|4(?:0\d|1[0-2]|9\d))|9[189]\d{2}|6[567]\d{2}|4(?:[14]0\d{3}|[68]\d{4}|[579]\d{2}))\d{6})[^0-9]""")


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.open_fi = tk.Button(self, width=23)
        self.open_fi["text"] = u"打开需要提取号码的文件夹"
        self.open_fi["command"] = self.open_file
        self.open_fi.pack(side="top")

        self.quit = tk.Button(self, text="退出", width=23, command=root.destroy)
        self.quit.pack(side="bottom")

    def open_file(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        results = []
        tar = {'zip', 'rar', '7z'}
        row = 'ABCDEF'
        first = ['号码', '脚本执行率', '电话营销语速', '电话营销态度', '电话营销技巧', '营销准确性']
        file_path = filedialog.askopenfilenames()
        for f in file_path:
            _, x = f.split('.', 1)
            if x not in  tar:
                fd = Path(f).name
                result = regex.findall(fd)
                if result:
                    results.append(result[0])
        length = len(results)
        for i, j in enumerate(row):
            sheet[str(j) + str(1)] = first[i]
        for i in range(length):
            sheet['A' + str(i + 2)] = results[i]
        workbook.save('D:/upload/{}.xlsx'.format(int(datetime.datetime.now().timestamp())))
        workbook.close()
        tk.messagebox.showinfo('提示', '保存成功，请去D盘寻找文件')


root = tk.Tk()
root.title(u"提取号码")
root.geometry('400x180')

theLabel = tk.Label(root,
                    text="\n手机号码提取小工具\n",
                    justify=tk.LEFT,
                    compound=tk.CENTER,
                    font=("华文行楷", 20),
                    fg="black")
theLabel.pack()

app = Application(master=root)
app.mainloop()
