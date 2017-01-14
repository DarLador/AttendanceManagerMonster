'''
Created on Feb 8, 2016

@author: Dar1
'''
from tkinter import *
from tkinter.ttk import *
import os
from ManagerView import main
import pandas as pd
import smtplib
from email.mime.text import MIMEText


class Intro():
    def __init__(self, root, path):
        self.root=root
        self.root.title("(c) Attendance Manager Monster")
        self.root.geometry("275x120")
        self.frame = Frame(root)
        self.mypath = path
        
        style = Style()
        style.configure("BW.TLabel", foreground="black", font=('Verdana', '10', 'roman'))        
        
        """Logo"""
        try:
            self.logo = PhotoImage(file= os.path.join(self.mypath,"Intro.gif"))
        except IOError:
            self.my_path()
            self.mypath = self.mypath.read()
            self.logo = PhotoImage(file= os.path.join(self.mypath,"Intro.gif"))
        label1 = Label(self.frame, image=self.logo)
        label1.image = self.logo 
        label1.grid(row = 0, column = 0, columnspan = 3, sticky=NW)
        
        self._manager = Label(self.frame, text = "Manager name: ", style="BW.TLabel")
        self._intro = Label(self.frame, text = "Enter password: ", style="BW.TLabel")
        self._manager.grid(row=1, sticky=W)
        self._intro.grid(row=2, sticky=W)
        
        self.frame.grid(row=0, column=0, sticky=(N, S, E, W))
        
        self.frame.columnconfigure(0, weight=2)
        self.frame.rowconfigure(0, weight=2)
        
        self.password = Entry(self.frame, show="*")
        self.mname = Entry(self.frame)
        self.mname.grid(row=1, column=1, sticky=(N, S, E, W))
        self.password.grid(row=2, column=1, sticky=(N, S, E, W))
        self.submit_button = Button(self.frame, text= 'Submit', command = self.submit_it, style='Fun.TButton')
        
#         self.last_saved_path = BooleanVar()
#         self.check1 = Checkbutton(self.frame, text='Use the last saved path', variable=self.last_saved_path, command = self.my_path, 
#                                   style ='TCheckbutton').grid(row=2, column=0, sticky=W)
                                  
        self.new_path = BooleanVar()
        Checkbutton(self.frame, text='Add a new path', variable=self.new_path, command = self.save_path, 
                    style ='TCheckbutton').grid(row=3, column=0, sticky=W) 
        
        self.submit_button.grid(row=4, column=0, columnspan=1, sticky=(N, S, E, W))
        self.quit_b = Button(self.frame, text= 'Quit', command = self.quit, style='Fun.TButton')
        self.quit_b.grid(row=4, column=1, sticky=(N, S, E, W))
    

    
    def my_path(self):
        #Delete
        if not os.path.exists('mypath.txt'):
            self.toplevel3 = Toplevel()
            self.lab30 = Label(self.toplevel3, text="Can't find your path. Please submit a new path.")
            self.lab30.grid(row=0)
            self.b30 = Button(self.toplevel3, text= 'OK', command = self.send_password_email, style='Fun.TButton')
            self.b30.grid(row=1)
        else:
            if os.stat("mypath.txt").st_size == 0:
                self.toplevel4 = Toplevel()
                self.lab4 = Label(self.toplevel4, text="Can't find your path. Please submit a new path.")
                self.lab4.grid(row=0)
                self.b40 = Button(self.toplevel4, text= 'OK', command = self.save_path, style='Fun.TButton')
                self.b40.grid(row=1)
                
            else:
                self.mypath=open('mypath.txt')         
           
    def save_path(self):
        #Delete
        self.toplevel1 = Toplevel()
        self.lab2 = Label(self.toplevel1, text="Enter your path")
        self.lab2.grid(row=0, sticky=(S,W))
        self.e2 = Entry(self.toplevel1, width=50)
        self.e2.grid(row=1, columnspan=3, sticky=(N, S, E, W))
        self.lab3 = Label(self.toplevel1, text="Format: C:\\\\Users\\\\...\\\\")
        self.lab3.grid(row=2, sticky=(S,W))
        self.b1 = Button(self.toplevel1, text= 'Save', command = self.save_new_path, style='Fun.TButton')
        self.b1.grid(row=3, sticky=(S,W))
        try: self.toplevel3.destroy()
        except AttributeError: pass
        try: self.top2.destroy()
        except AttributeError: pass
        try: self.toplevel4.destroy()
        except AttributeError: pass
        self.last_saved_path.set(None)
        self.new_path.set(None)
    
    def save_new_path(self):
        file = open("mypath.txt", "w")
        try:
            pat = self.e2.get()
            file.write(pat)
            file.close()
            self.mypath = open('mypath.txt', 'r')
            self.toplevel1.destroy()
        except AttributeError:
            self.top2 = Toplevel()
            self.lab1 = Label(self.top2, text="Please submit a new path.")
            self.lab1.grid(row=0)
            self.self.b2 = Button(self.top2, text= 'OK', command = self.save_path, style='Fun.TButton')
            self.b2.grid(row=1)
                    
    def quit(self):
        self.frame.quit()
    
    def submit_it(self):
        df = pd.read_pickle(os.path.join(self.mypath, 'ManagersInfo.dat'))
        content = self.password.get()
        mname = self.mname.get()
        if content == 'shecodes' and mname in df.index: 
            if mname == 'Dar Lador' or mname=='Toot Moran':
                self.topi = Toplevel()
                label1i = Label(self.topi, image=self.logo)
                label1i.image = self.logo 
                label1i.grid(row = 0, column = 0, columnspan = 3, sticky=NW)
                
                self._track_s = Label(self.topi, text = "Select a track: ", style="BW.TLabel")
                self._track_s.grid(row=2, sticky=W)
                self.br = IntVar()
                self.br.set(2)
                for i,m in enumerate(['android', 'web']):
                    self.choose_b = Radiobutton(self.topi, text=m, variable=self.br, value = i, command = self.get_track)
                    self.choose_b.grid(row=3, column=i, sticky=W+E)
                
                #self.topi.grid(row=0, column=0, sticky=(N, S, E, W))
                
                self.topi.columnconfigure(0, weight=2)
                self.topi.rowconfigure(0, weight=2)
                
                self.submit_next = Button(self.topi, text= 'Submit', command = self.open_mview, style='Fun.TButton')
                self.submit_next.grid(row=4, column=0, sticky=W+E)
                self.speciality = df.loc[mname, 'Specialty']
                self.me = mname
                
        else:
            self.errormanager = Toplevel()
            self.merror = Label(self.toplevel1, text="Unable to identify manager")
            self.merror.grid(row=0, sticky=(S,W))
            self.qm = Button(self.toplevel1, text= 'Quit', command = self.quit, style='Fun.TButton')
            self.bm.grid(row=3, sticky=(S,W))
    def get_track(self):
        self.my_sp = ['android', 'web'][self.br.get()]
    def open_mview(self):
        self.frame.destroy()
        self.topi.destroy()
        root.destroy()
        main(self.mypath, self.my_sp, self.me)
# if __name__ == '__main__':
#     root = Tk()
#     path = os.getcwd() if not os.path.exists('C:/Users/me') else 'C:/Users/me'
#     root.iconbitmap(os.path.join(path, 'CookieMonster.ico'))
#     b = Intro(root, path)
#     root.mainloop()
root = Tk()
root.geometry("275x120")
root.title("(c) Attendance Manager Monster")
#path = os.getcwd() if not os.path.exists('C:/Users/me') else 'C:/Users/me'
path = 'C:/Users/Dar1/Documents/she_codes'
root.iconbitmap(os.path.join(path, 'CookieMonster.ico'))
b = Intro(root, path)
root.mainloop()