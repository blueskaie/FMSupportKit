import tkinter as Tk
import keyboard
import win32com.client as win32
import win32ui
import os
from datetime import datetime, timedelta
from tkcalendar import Calendar, DateEntry
from babel.numbers import *

class MyApp(object):

    def __init__(self, parent):

        self.root = parent
        self.root.title("Main frame")
        self.root.protocol("WM_DELETE_WINDOW", self.cancelCallBack)
        
        todaystr = datetime.today().strftime("%Y-%m-%d")
        yesterdaystr = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")

        self.todayLabel = Tk.Label(self.root)
        self.todayLabel.configure(height=1, width=3)
        self.todayLabel.configure(font=("Calibri", 8))
        self.todayLabel.configure(text="Received Today")
        self.todayLabel.configure(compound="left")
        self.todayLabel.place(relx=0.05, rely=0.02, relheight=0.1, relwidth=0.4)

        self.todayBtn = Tk.Button(self.root, command = self.todayCallBack)
        self.todayBtn.configure(height=1, width=3)
        self.todayBtn.configure(font=("Calibri", 8))
        self.todayBtn.configure(text=todaystr + " - FileName")
        self.todayBtn.place(relx=0.45, rely=0.02, relheight=0.1, relwidth=0.5)
        
        self.yesterdayLabel = Tk.Label(self.root)
        self.yesterdayLabel.configure(height=1, width=3)
        self.yesterdayLabel.configure(font=("Calibri", 8))
        self.yesterdayLabel.configure(text="Received Yesterday")
        self.yesterdayLabel.place(relx=0.05, rely=0.14, relheight=0.1, relwidth=0.4)
        
        self.yesterdayBtn = Tk.Button(self.root, command = self.yesterdayCallBack)
        self.yesterdayBtn.configure(height=1, width=3)
        self.yesterdayBtn.configure(font=("Calibri", 8))
        self.yesterdayBtn.configure(text=yesterdaystr + " - FileName")
        self.yesterdayBtn.place(relx=0.45, rely=0.14, relheight=0.1, relwidth=0.5)
        
        self.cal = Calendar(self.root, font="Calibri 8", selectmode='day', locale='en_US', cursor="hand2", showweeknumbers=False)
        self.cal.place(relx=0.1, rely=0.27, relheight=0.6, relwidth=0.8)
   
        self.curdayBtn = Tk.Button(self.root, command = self.curdayCallBack)
        self.curdayBtn.configure(height=1, width=3)
        self.curdayBtn.configure(font=("Calibri", 8))
        self.curdayBtn.configure(text="Use date selected from calendar")
        self.curdayBtn.place(relx=0.05, rely=0.88, relheight=0.1, relwidth=0.7)
        
        self.cancelBtn = Tk.Button(self.root, command = self.cancelCallBack)
        self.cancelBtn.configure(height=1, width=3)
        self.cancelBtn.configure(font=("Calibri", 8))
        self.cancelBtn.configure(text="Cancel")
        self.cancelBtn.place(relx=0.78, rely=0.88, relheight=0.1, relwidth=0.17)
        
        keyboard.add_hotkey('ctrl+shift+v', self.show)
        self.hide()

    def changeFileNames(self, template):
        files = self.explorer_fileselection()
        if files:
            for f in files:
                dirpath = os.path.dirname(os.path.abspath(f))
                filename = os.path.basename(os.path.abspath(f))
                newfilename = template + " - " + filename
                src = dirpath + "\\" + filename
                dst = dirpath + "\\" + newfilename
                os.rename(src, dst)
        else:
            win32ui.MessageBox("No seleted Files!", "Error")

    def todayCallBack(self):
        todaystr = datetime.today().strftime("%Y-%m-%d")
        self.changeFileNames(todaystr)
        self.hide()
    
    def yesterdayCallBack(self):
        yesterdaystr = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        self.changeFileNames(yesterdaystr)
        self.hide()

    def curdayCallBack(self):
        curdatestr = self.cal.selection_get().strftime("%Y-%m-%d")
        self.changeFileNames(curdatestr)
        self.hide()

    def cancelCallBack(self):
        self.hide()
    
    def explorer_fileselection(self):
        clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}' #Valid for IE as well!
        shellwindows = win32.Dispatch(clsid)
        files = []
        try:
            for window in range(shellwindows.Count):
                window_URL = shellwindows[window].LocationURL
                window_dir = window_URL.split("///")[1].replace("/", "\\")
                selected_files = shellwindows[window].Document.SelectedItems()
                for file in range(selected_files.Count):
                    files.append(selected_files.Item(file).Path)
        except:
            win32ui.MessageBox("Close IE!", "Error")
        del shellwindows
        return files

    def show(self):
        self.root.attributes('-topmost', True)
        self.root.update()
        self.root.attributes('-topmost', False)
        self.root.deiconify()
        

    def hide(self):
        self.root.update()
        self.root.withdraw()

def checkRegister():
    from datetime import datetime
    limited = datetime(2019, 5, 28, 0, 0)
    today = datetime.today()
    if (today > limited):
        return False
    else:
        return True

if __name__ == "__main__":

    root = Tk.Tk()
    w = 300
    h = 420
    x = (root.winfo_screenwidth() - w) / 2
    y = (root.winfo_screenheight() - h) / 2
    root.geometry("%dx%d+%d+%d" % (w,h,x,y))
    
    app = MyApp(root)
    root.mainloop()
