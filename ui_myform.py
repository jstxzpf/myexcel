"""
本代码由[Tkinter布局助手]生成
当前版本:3.1.2
官网:https://www.pytk.net/tkinter-helper
QQ交流群:788392508
"""
from tkinter import *
from tkinter.ttk import *

"""
全局通用函数
"""
# 自动隐藏滚动条
def scrollbar_autohide(bar,widget):
    def show():
        bar.lift(widget)
    def hide():
        bar.lower(widget)
    hide()
    widget.bind("<Enter>", lambda e: show())
    bar.bind("<Enter>", lambda e: show())
    widget.bind("<Leave>", lambda e: hide())
    bar.bind("<Leave>", lambda e: hide())
    
class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_label_lf0ug3rp = self.__tk_label_lf0ug3rp()
        self.tk_label_lf0uhbkt = self.__tk_label_lf0uhbkt()
        self.tk_select_box_lf0ui9ef = self.__tk_select_box_lf0ui9ef()
        self.tk_select_box_lf0ukjen = self.__tk_select_box_lf0ukjen()
        self.tk_label_lf0um2xa = self.__tk_label_lf0um2xa()
        self.tk_button_lf0unqii = self.__tk_button_lf0unqii()
        self.tk_button_lf0untrd = self.__tk_button_lf0untrd()
        self.tk_button_lf0unw0h = self.__tk_button_lf0unw0h()

    def __win(self):
        self.title("Tkinter布局助手")
        # 设置窗口大小、居中
        width = 600
        height = 500
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)

    def __tk_label_lf0ug3rp(self):
        label = Label(self,text="请选择",anchor="center")
        label.place(x=10, y=10, width=50, height=24)
        return label

    def __tk_label_lf0uhbkt(self):
        label = Label(self,text="年",anchor="center")
        label.place(x=125, y=10, width=20, height=24)
        return label

    def __tk_select_box_lf0ui9ef(self):
        cb = Combobox(self, state="readonly")
        cb['values'] = ("列表框","Python","Tkinter Helper")
        cb.place(x=70, y=10, width=50, height=24)
        return cb

    def __tk_select_box_lf0ukjen(self):
        cb = Combobox(self, state="readonly")
        cb['values'] = ("列表框","Python","Tkinter Helper")
        cb.place(x=150, y=10, width=50, height=24)
        return cb

    def __tk_label_lf0um2xa(self):
        label = Label(self,text="月",anchor="center")
        label.place(x=210, y=10, width=20, height=24)
        return label

    def __tk_button_lf0unqii(self):
        btn = Button(self, text="按钮")
        btn.place(x=10, y=70, width=50, height=24)
        return btn

    def __tk_button_lf0untrd(self):
        btn = Button(self, text="按钮")
        btn.place(x=70, y=70, width=50, height=24)
        return btn

    def __tk_button_lf0unw0h(self):
        btn = Button(self, text="按钮")
        btn.place(x=150, y=70, width=50, height=24)
        return btn

class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.__event_bind()

    def __event_bind(self):
        pass
        
if __name__ == "__main__":
    win = Win()
    win.mainloop()
