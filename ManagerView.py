# coding=utf-8
'''
Created on Feb 8, 2016

@author: Dar1
'''
from tkinter import *
from tkinter.ttk import *
import tkinter.font as tkFont
from pandas import read_excel, read_pickle, DataFrame, to_pickle, Series, \
                isnull, ExcelWriter, to_datetime
import os
from config import create_lesson_state_file, update_table, xlsx_attendance, \
                excel_saver
import numpy as np
import tkinter.scrolledtext as tkst
import smtplib
from email.mime.text import MIMEText
from datetime import date
import datetime as dtime

class ManagerView():
    def __init__(self, root, my_path, specialty, me):
        self.root=root
        self.me = me
        self.path = my_path
        self.notebook = Notebook(root)
        self.specialty = specialty
        self.lab_opt = { 'background':'darkgreen', 'foreground':'white'  }
        self.helv12 = tkFont.Font(family='Helvetica', size=12)       
        self.root.geometry("592x520")
        self.root.title("(c) Attendance Manager Monster")
        #Resize rules
        root.columnconfigure(0, weight=1)
        self.notebook.grid(row=0, column=0, sticky=(N, S, E, W))
        
        #All notebooks (f1-f5)
        f1=Frame(self.notebook)
        f2 = Frame(self.notebook)
        f3 = Frame(self.notebook)
        f8 = Frame(self.notebook)
        f4 = Frame(self.notebook)
        f5 = Frame(self.notebook)
        f6 = Frame(self.notebook)
        f7 = Frame(self.notebook)
        
        f1.grid(row=0, column=0, sticky=(N, S, E, W))
        f2.grid(row=0, column=0, sticky=(N, S, E, W))
        f3.grid(row=0, column=0, sticky=(N, S, E, W))
        f8.grid(row=0, column=0, sticky=(N, S, E, W))
        f4.grid(row=0, column=0, sticky=(N, S, E, W))
        f5.grid(row=0, column=0, sticky=(N, S, E, W))
        f6.grid(row=0, column=0, sticky=(N, S, E, W))
        f7.grid(row=0, column=0, sticky=(N, S, E, W))
        
        self.notebook.add(f1, text='Summary')
        self.notebook.add(f2, text='Detailed Report')
        self.notebook.add(f3, text='Meetups')
        self.notebook.add(f8, text='Progress report')
        self.notebook.add(f4, text='Complete lessons')
        self.notebook.add(f5, text='Delete/restore')
        self.notebook.add(f6, text='Weekly email')
        if self.me=='Dar Lador':
            self.notebook.add(f7, text='Notes')
        
        '''All images'''
        self.logo = PhotoImage(file = os.path.join(self.path, "shecodegif.gif"))
        self.xl_photo = PhotoImage(file = os.path.join(self.path, "excel.gif"))
        self.exitdoor_photo = PhotoImage(file = os.path.join(self.path, "image009.gif"))
        self.email_icon = PhotoImage(file = os.path.join(self.path, 'email_icon_small.gif'))
        self.mini_shecodes = PhotoImage(file = os.path.join(self.path, 'dar.gif'))
        
        """Logo"""
        self.label1 = Label(f1, image=self.logo)
        self.label2 = Label(f2, image=self.logo)
        self.label3 = Label(f3, image=self.logo)
        self.label8 = Label(f8, image=self.logo)
        self.label5 = Label(f5, image=self.logo)
        self.label7 = Label(f7, image=self.logo)
        self.label1.image = self.logo 
        self.label1.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label2.image = self.logo 
        self.label2.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label3.image = self.logo 
        self.label3.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label8.image = self.logo 
        self.label8.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label5.image = self.logo 
        self.label5.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label5.configure(width=30)
        self.label7.grid(row = 0, column = 0, columnspan = 30, sticky=NW)
        self.label7.configure(width=30)
        # track_contacts.xlsx
        self.w = read_excel(os.path.join(self.path,'track_contacts.xlsx'), sheetname = 'tracks', header = 0)
        try:
            self.attlist = read_pickle(os.path.join(self.path, 'attlist.dat'))
            dar = ['Dar Android', 'Dar Web', 'Dar Neww', 'Dar Newa']
            for col in self.attlist.columns:
                if self.attlist[col].count() <=3:
                    for d in dar:
                        if not isnull(self.attlist.loc[d, col]):
                            self.attlist =self.attlist.drop(col,1)
            
        except (IOError, KeyError):
            pass
        try:
            self.LessonState = read_pickle(os.path.join(self.path, 'LessonState.dat'))
            self.just_dates = self.attlist.drop(['Studying', 'Starting month', 'Lesson'],1)
            self.debug_dar = [s for s in self.LessonState.index if 'Dar' in s]
        except:
            create_lesson_state_file(self.path)
            self.LessonState = read_pickle(os.path.join(self.path, 'LessonState.dat'))
        # manager info
        self.manager_inf = read_pickle(os.path.join(self.path, 'ManagersInfo.dat'))
        # former members:
        self.former_att = read_pickle(os.path.join(self.path, 'FormerMembersAtt.dat'))
        self.former_inf = read_pickle(os.path.join(self.path, 'FormerMembersInf.dat'))
        self.former_attlist = read_pickle(os.path.join(self.path, 'FormerLessonState.dat'))
        #=========================================================
        #                    (f1) Summary table
        #=========================================================
        
        #creating pandas df
        if self.specialty == 'android':
            self.track = ['android', 'android_new', 'Android', 'Android (New)']
            self.folder_track = 'android_'
            self.t_name_old, self.t_name_new = 'Android', 'Android (New)'
            self.summary_table()
        if self.specialty == 'web':
            self.track = ['web', 'web_new', 'Web', 'Web (New)']
            self.folder_track = 'web_'
            self.t_name_old, self.t_name_new = 'Web', 'Web (New)'
            self.summary_table()
        
        #subset the data and convert to giant list of strings (rows) for viewing        
        self.sub_data = self.summery.ix[self.dat_rows, self.dat_cols].sort_values(by='Name')
        self.sub_datstring = self.sub_data.to_string(index=False, col_space=13).split('\n')
        self.title_string  = self.sub_datstring[0]
        
        #save the format of the lines, so we can update them without re-running df.to_string()
        self._get_line_format(self.title_string)
        self.scrollbarV = Scrollbar(f1, orient=VERTICAL)
        self.scrollbarV.grid(row=6, column=29, rowspan=3, sticky=N+S)
        self.scrollbarH = Scrollbar(f1, orient=HORIZONTAL)
        self.scrollbarH.grid(row=9, column=0, columnspan=28, sticky=E+W)
        
        #Listboxes: title+table
        self.title_lb = Listbox(f1,height=1, 
                                    font=tkFont.Font(f1, family="Courier", size=14),
                                    yscrollcommand=self.scrollbarV.set, 
                                    xscrollcommand=self.scrollbarH.set,
                                    exportselection=False)
        
        self.lbSites = Listbox(f1, 
                               font=tkFont.Font(f1, family="Courier", size=14),
                               yscrollcommand=self.scrollbarV.set,
                               xscrollcommand=self.scrollbarH.set,
                               exportselection=False,
                               selectmode=EXTENDED)
        self.title_lb.grid(row=1, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E)#4
        self.lbSites.grid(row=6, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E)#4
        self.scrollbarV.config(command=self.lbSites.yview)
        self.scrollbarH.config(command=self._xview)
        
        #Fill listboxes
        self._fill(f1)
        
        # order by column label + option menu
        #------------------------------------
        self.col_sel_lab = Label(f1, text='Order by column:',**self.lab_opt)
        self.col_sel_lab.grid(row=10, columnspan=1,sticky=W+E)
        self.var = StringVar()
        self.var.trace("w", self.order_by_)
        option = OptionMenu(f1, self.var, *['Choose', 'Name', '# Lesson', '# Absences', 'Group', 'Track'])
        option.grid(row=11, rowspan=2, columnspan=1,sticky=W+E)
        self.var.set('Choose') # initial value
        
        # Save-excel button
        #------------------
        self.button_xl = Button(f1, compound=TOP, image=self.xl_photo,
                                text="Save", command=self.xlsx_summary_conf_mssg)
        self.button_xl.grid(row=10, column = 2, rowspan=2, columnspan=1, sticky=NW)
        self.button_xl.image = self.xl_photo
        # Exit-door button
        #------------------
        self.exit_f1 = Button(f1, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f1.grid(row=10, column = 27, rowspan=3, columnspan=5, sticky=NW)
        
        ####################################################
        ####################################################
        self.tree_view_f2(f2)
        self.absence_selection(f3)
        self.tree_view_progress(f8)
        self.complete_lessons(f4)
        self.delete_restore_member(f5)
        self.send_weekly_email(f6)
        if self.me=='Dar Lador':
            self.print_notes(f7)

    #=================================================================
    #                     (f2) Trees - Attendances
    #=================================================================
    def tree_view_f2(self, parent):
        self.ttree = Treeview(parent)
        for file_ in os.listdir(os.path.join(self.path, self.folder_track)):
            self.taskmon = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,file_))
            sh = re.search(r'([\D]*)([\d]+)',file_).group(1)
            eet = re.search(r'([\D]*)([\d]+)',file_).group(2)
            sheet_n = sh+', '+eet
            id1 = self.ttree.insert('', 'end', text = sheet_n)
            for member in self.taskmon.index:
                s = self.taskmon.loc[member].replace('V', '').fillna('Missed')
                id2 = self.ttree.insert(id1, 'end', text = member)
                str_taskmonl = s.to_string().split('\n')
                for i in str_taskmonl:
                    self.ttree.insert(id2, 'end', text = i)
        self.ttree.grid(row=6, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E) 
        self.scrollbarV2 = Scrollbar(parent, orient=VERTICAL)
        self.scrollbarV2.grid(row=6, column=29, rowspan=3, sticky=N+S)

        # Save-excel button
        #------------------
        self.button_x2 = Button(parent, compound=TOP, image=self.xl_photo,
                                text="Save", command=self.xlsx_attendance_conf_mssg)
        self.button_x2.grid(row=12, column = 0, rowspan=2, columnspan=1, sticky=SE)
        self.button_x2.image = self.xl_photo
        self.exit_f2 = Button(parent, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f2.grid(row=12, column = 25, rowspan=3, columnspan=5, sticky=NW)
    
    #==================================================================
    #                    (f8) Progress
    #==================================================================
    def tree_view_progress(self, f8):
        self.ttree_ = Treeview(f8)
        self.sum_name_index = self.summery.set_index('Name')
        self.prog = self.LessonState.loc[self.LessonState.index.isin(self.sum_name_index.index)]
        for file_ in os.listdir(os.path.join(self.path, self.folder_track)):
            self.taskmon = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,file_))
            sh = re.search(r'([\D]*)([\d]+)',file_).group(1)
            eet = re.search(r'([\D]*)([\d]+)',file_).group(2)
            sheet_n = sh+', '+eet
            id1 = self.ttree_.insert('', 'end', text = sheet_n)
            for member in self.taskmon.index:
                try:
                    s = self.prog.loc[member].dropna()
                except:
                    print('debug')
                id2 = self.ttree_.insert(id1, 'end', text = member)
                str_taskmonl = s.to_string().split('\n')
                for i in str_taskmonl:
                    self.ttree_.insert(id2, 'end', text = i)
        self.ttree_.grid(row=6, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E) 
        
        # Save-excel button
        #------------------
        self.button_x8 = Button(f8, compound=TOP, image=self.xl_photo,
                                text="Save", command=self.xlsx_prog_conf_mssg)
        self.button_x8.grid(row=12, column = 0, rowspan=2, columnspan=1, sticky=SE)
        self.button_x8.image = self.xl_photo
        self.exit_f8 = Button(f8, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f8.grid(row=12, column = 25, rowspan=3, columnspan=5, sticky=NW)
    
    def xlsx_prog_conf_mssg(self):
        tx = 'Progress_'+self.specialty+'.xlsx'
        writer = ExcelWriter(os.path.join(os.path.join(self.path, 'SavedFiles'),tx), engine='xlsxwriter')
        self.prog.to_excel(writer, sheet_name = 'Sheet1')
        writer.save()
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.lab1 = Label(self.top1, text="File saved as "+tx)
        self.lab1.grid(row=0)
        self.lab2 = Label(self.top1, text="in SavedFiles in your she codes folder.")
        self.lab2.grid(row=1)
        self.b2 = Button(self.top1, text= 'OK', command = self.destroy_top1, style='Fun.TButton')
        self.b2.grid(row=2)
    #=================================================================
    #                     (f3) Absence Selection
    #=================================================================
    def absence_selection(self, f3):
        df = DataFrame()
        for file_ in os.listdir(os.path.join(self.path, self.folder_track)):
            taskmon = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,file_))
            df = taskmon if df.empty else df.append(taskmon)
        df = df.rename(columns = lambda col: to_datetime(col,format="%d/%m/%Y"))
        df = df.T.sort_index(ascending='False').T #sort_index
        df = df.loc[df.index.isin(self.w.FullName.values)]
        for name in df.index:
            start_date = self.w[self.w.FullName==name].Start.values[0]
            for col in df.columns:
                if col<start_date:
                    df.loc[name, col] = "Didn't start yet"
        df=df.rename(columns = lambda x: x.strftime('%d/%m/%Y'))
        df.index.name = 'Name'
        self.df = df.reset_index()
        dat_cols = list(self.df)
        dat_rows = range(len(self.df))
        sub_data = self.df.ix[dat_rows, dat_cols].sort_values(by='Name')
        sub_datstring = sub_data.to_string(index=False, col_space=13).split('\n')
        title_string  = sub_datstring[0]
        
        #save the format of the lines, so we can update them without re-running df.to_string()
        self._get_line_format(title_string)
        scrollbarV = Scrollbar(f3, orient=VERTICAL)
        scrollbarV.grid(row=6, column=29, rowspan=3, sticky=N+S)
        scrollbarH = Scrollbar(f3, orient=HORIZONTAL)
        scrollbarH.grid(row=9, column=0, columnspan=28, sticky=E+W)
        
        self.title_lb2 = Listbox(f3,height=1, 
                                    font=tkFont.Font(f3, family="Courier", size=14),
                                    yscrollcommand=scrollbarV.set, 
                                    xscrollcommand=scrollbarH.set,
                                    exportselection=False)
        
        self.lbSites2 = Listbox(f3, 
                               font=tkFont.Font(f3, family="Courier", size=14),
                               yscrollcommand=scrollbarV.set,
                               xscrollcommand=scrollbarH.set,
                               exportselection=False,
                               selectmode=EXTENDED)
        self.title_lb2.grid(row=1, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E)#4
        self.lbSites2.grid(row=6, column=0, columnspan=29, rowspan=3, sticky=N+W+S+E)#4
        scrollbarV.config(command=self.lbSites2.yview)
        scrollbarH.config(command=self._xview2)
        self._fill(f3)
        self.title_lb2.insert(END, title_string)
        for line in sub_datstring[1:]:
            self.lbSites2.insert(END, line) 
            self.lbSites2.bind('<ButtonRelease-1>',self._listbox_callback)
        self.lbSites2.select_set(0)
        self.dict = {'Filter by': ['','',''],
                     'Was absence in the recent': ['1 week', '2 weeks', '3 weeks', '4 weeks', 'One month and more'],
                     'Number of absences': [x for x in range(32)]}
        
        self.variable_a = StringVar()
        self.variable_b = StringVar()
        self.variable_c = StringVar()
        
        self.variable_a.trace("w", self.update_options)
        self.lstbox = Listbox(f3, 
                              listvariable=self.variable_c, 
                              selectmode=MULTIPLE, height=4, width=25)
        self.lstbox.grid(column=9, row=10, rowspan =3, columnspan=10, sticky=W+E)
        self.lstbox.bind('<<ListboxSelect>>', self.load_models)
        self.lstbox.insert(END, '')
        sclstbox = Scrollbar(f3, orient=VERTICAL)
        sclstbox.grid(row=10, column=19, rowspan =3, sticky=N+S)
        sclstbox.config(command=self.lstbox.yview)
        
        self.optionmenu_b = OptionMenu(f3, self.variable_b, '')
        self.optionmenu_a = OptionMenu(f3, self.variable_a, *['Filter by', 'Was absence in the recent', 'Number of absences'])
        
        self.variable_a.set('Filter by')
        
        
        self.optionmenu_a.grid(row=10, rowspan = 1, columnspan=1,sticky=W+E)
        self.optionmenu_a.configure(width=25)
        self.optionmenu_b.grid(row=11, rowspan = 1, columnspan=1,sticky=W+E)
        
        #-----------------------
        # Send an email button
        #-----------------------
        self.send_email_button = Button(f3, compound = TOP, text = 'email', image = self.email_icon, command=self.send_email)
        self.send_email_button.grid(row=10, column=3, rowspan=3, columnspan=6, sticky=W+E)
        #------------------
        # Exit-door button
        #------------------
        self.exit_f1 = Button(f3, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f1.grid(row=10, column = 27, rowspan=3, columnspan=5, sticky=NW)
        
    def load_models(self, *args):
        try:
            self.selected_names = self.lstbox.selection_get()
        except TclError:
            pass
        #print(self.selected_names)
        
    def update_options(self, *args):
        countries = self.dict[self.variable_a.get()]
        self.variable_b.set(countries[0])
 
        menu = self.optionmenu_b['menu']
        menu.delete(0, 'end')
 
        for country in countries:
            menu.add_command(label=country, command=lambda nation=country: self.filtered_table(nation))
    
    def filtered_table(self, nation):
        self.variable_b.set(nation)
        filter_ = self.variable_b.get()
        if filter_ == '':
            pass
        else:
            self.names_to_email = self.edit_decision(nation)
            self.lstbox.delete(0, END)
            for n in self.names_to_email:
                self.lstbox.insert(END, n)
            
            
    def summary_table(self):
        self.w_a = self.w.loc[self.w['track'].isin(self.track)]
        self.attlist_a = self.attlist.loc[self.attlist['Studying'].isin(self.track)]
        update_table(self.w_a, self.attlist_a, self.just_dates, self.former_att, os.path.join(self.path, self.folder_track))
        self.summery = DataFrame(columns = ['# Lesson', '# Absences', 'Group', 'Track'])
        for file_ in os.listdir(os.path.join(self.path, self.folder_track)):
            self.taskmon = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,file_))
            sh = re.search(r'([\D]*)([\d]+)',file_).group(1)
            eet = re.search(r'([\D]*)([\d]+)',file_).group(2)
            sheet_n = sh+', '+eet
            for member in self.taskmon.index:
                try:
                    for i in self.LessonState.loc[member].index:
                        if isnull(self.LessonState.loc[member][i]):
                            meetup_ = i-1
                            break
                except:
                    # Seems that a member was removed from everywhere except the 'month' file.
                    # Here we remove that member from the 'month' file.
                    self.taskmon = self.taskmon.drop(member)
                    self.taskmon.to_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,file_))
                try:
                    s = Series([meetup_, 
                                len(self.taskmon.columns)-self.taskmon.loc[member].count(),
                                sheet_n,
                                self.t_name_old if self.w_a[self.w_a.FullName==member]['track'].values[0]== self.track[0] else self.t_name_new], 
                               index=['# Lesson', '# Absences', 'Group', 'Track'])
                    s.name = member
                    self.summery = self.summery.append(s)
                except:
                    pass
        self.summery.index.name = 'Name'
        self.summery = self.summery.reset_index()
        self.dat_cols = list(self.summery)
        self.dat_rows = range(len(self.summery))
        self.rowmap   =  {i:row for i,row in enumerate(self.dat_rows)}
    
    def order_by_(self, *args):
        wanted = self.var.get()
        #self.summery = self.summery.sort_values(wanted)
        if wanted == 'Choose':
            pass
        else:
            if self.lbSites.cget('state') == DISABLED:
                self.lbSites.config(state=NORMAL)
                self.lbSites.delete(0, END)
                self.summery = self.summery.sort_values(wanted)
                self.dat_cols = list(self.summery)
                self.dat_rows = range(len(self.summery))
                self.sub_data = self.summery.ix[self.dat_rows, self.dat_cols].sort_values(by='Name')
                self.sub_datstring = self.sub_data.to_string(index=False, col_space=13).split('\n')
                for line in self.sub_datstring[1:]:
                    self.lbSites.insert(END, line) 
                    self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
                self.lb.config(state=DISABLED)     
            else:
                self.lbSites.delete(0, END)
                self.summery = self.summery.sort_values(wanted)
                self.dat_cols = list(self.summery)
                self.dat_rows = range(len(self.summery))
                self.sub_data = self.summery.ix[self.dat_rows, self.dat_cols].sort_values(by=wanted)
                self.sub_datstring = self.sub_data.to_string(index=False, col_space=13).split('\n')
                for line in self.sub_datstring[1:]:
                    self.lbSites.insert(END, line) 
                    self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
            self.lbSites.select_set(0)
        
        
        
    def edit_decision(self, nation):
        d = self.df.set_index('Name')
        if type(nation)==str:
            abse = {'1 week': -1, '2 weeks': -2, '3 weeks': -3, '4 weeks': -4, 'One month and more':-5}
            req = d[isnull(d.ix[:,abse[nation]:]).all(1)]
            req = req.reset_index()
            req_cols = list(req)
            req_rows = range(len(req))
            req_sub_data = req.ix[req_rows, req_cols].sort_values(by='Name')
            req_datastring = req_sub_data.to_string(index=False, col_space=14).split('\n')
            if self.lbSites2.cget('state') == DISABLED:
                self.lbSites2.config(state=NORMAL)
                self.lbSites2.delete(0, END)
                if not req.empty:
                    for line in req_datastring[1:]:
                        self.lbSites2.insert(END, line)
                        self.lbSites2.bind('<ButtonRelease-1>',self._listbox_callback)
                    self.lbSites2.config(state=DISABLED) 
            else:
                self.lbSites2.delete(0, END)
                if not req.empty:
                    for line in req_datastring[1:]:
                        self.lbSites2.insert(END, line)
                        self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
                self.lbSites2.config(state=DISABLED)
            self.lbSites2.select_set(0)
        else:
            wanted_names = self.summery[self.summery['# Absences']==nation]['Name'].values
            req = d.loc[d.index.isin(wanted_names)]
            #req = req.rename(columns = lambda col: col.strftime('%d/%m/%Y'))
            req = req.reset_index()
            req_cols = list(req)
            req_rows = range(len(req))
            req_sub_data = req.ix[req_rows, req_cols].sort_values(by='Name')
            req_datastring = req_sub_data.to_string(index=False, col_space=14).split('\n')
            if self.lbSites2.cget('state') == DISABLED:
                self.lbSites2.config(state=NORMAL)
                self.lbSites2.delete(0, END)
                for line in req_datastring[1:]:
                    self.lbSites2.insert(END, line)
                    self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
                self.lbSites2.config(state=DISABLED) 
            else:
                for line in req_datastring[1:]:
                    self.lbSites2.insert(END, line)
                    self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
                self.lbSites2.config(state=DISABLED)
        return req.Name.values
            #print('debug')
    
    def xlsx_summary_conf_mssg(self):
        tx = 'SummaryTable.xlsx'
        writer = ExcelWriter(os.path.join(os.path.join(self.path, 'SavedFiles'),tx), engine='xlsxwriter')
        self.summery.to_excel(writer, sheet_name='Sheet1')
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.lab1 = Label(self.top1, text="File saved as "+tx)
        self.lab1.grid(row=0)
        self.lab2 = Label(self.top1, text="in SavedFiles in your she codes folder.")
        self.lab2.grid(row=1)
        self.b2 = Button(self.top1, text= 'OK', command = self.destroy_top1, style='Fun.TButton')
        self.b2.grid(row=2)
    
    def xlsx_attendance_conf_mssg(self):
        tx = xlsx_attendance(os.path.join(self.path, 'SavedFiles'), 
                             os.path.join(self.path, self.folder_track), 
                             self.specialty)
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.lab1 = Label(self.top1, text="File saved as "+tx)
        self.lab1.grid(row=0)
        self.lab2 = Label(self.top1, text="in SavedFiles in your she codes folder.")
        self.lab2.grid(row=1)
        self.b2 = Button(self.top1, text= 'OK', command = self.destroy_top1, style='Fun.TButton')
        self.b2.grid(row=2)
    
    def send_email(self):
        try:
            to_ = ', '.join(self.selected_names)
            self.top1 = Toplevel()
            self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
            self.top1.geometry("%dx%d%+d%+d" % (500, 500, 300, 125))
            self.lab00 = Label(self.top1, text = 'To:')
            self.lab00.grid(row=0, column=0, sticky=W)
            self.to_names = tkst.ScrolledText(self.top1, wrap=WORD, width=50, height=3)
            to_ = self.selected_names.replace('\n',', ')
            self.to_names.insert(END, to_)
            self.to_names.grid(row=0, column=1, columnspan=5, sticky=NE)
            self.to_names.config(width=55)
            self.cc_ = Label(self.top1, text="CC: ")
            self.cc_.grid(row=2, column=0, sticky=W)
            picks = ['WIS managers', 'Android managers'] if self.specialty == 'android' else \
                    ['WIS managers', 'Web managers']
            self.cc_s = []
            self.cc_who = StringVar()
            self.cc_who.set('3')
            self.chk1 = Radiobutton(self.top1, text='WIS managers', variable=self.cc_who, value = '0', command = self.send_dar_toot)
            self.chk1.grid(row=2, column=1, sticky=W)
            
            self.lab1 = Label(self.top1, text="From: ")
            self.lab1.grid(row=3, column=0, sticky=W)
            self.my_var = StringVar()
            self.my_var.set('No one')
            self.from_ = Radiobutton(self.top1, text=self.me, variable=self.my_var, value=self.me, command = self.insert_pass)
            self.from_D = Radiobutton(self.top1, text='Dar SheCodes', variable=self.my_var, value='Dar SheCodes', command = self.insert_pass)
            self.from_.grid(row=3, column=1, sticky=W)
            self.from_D.grid(row=3, column=2, sticky=W)
            
            self.lab33 = Label(self.top1, text = 'Email body:')
            self.lab33.grid(row=5, column=0, columnspan=2, sticky=W)
            self.body_email1 = tkst.ScrolledText(self.top1, wrap=WORD, width=60, height=20)
            self.body_email1.grid(row=6, column=0, columnspan=5, sticky=N+E+S+W)
            
            self.subject1 = Label(self.top1, text='Subject: ')
            self.subject1.grid(row=4,column=0)
            self.subject11 = Entry(self.top1)
            self.subject11.grid(row=4,column=1, columnspan=4, sticky=N+E+S+W)
            self.b22 = Button(self.top1, text= 'Cancel', command = self.destroy_top1, style='Fun.TButton')
            self.b22.grid(row=7, column=1)
            self.b33 = Button(self.top1, text= 'Send', state=DISABLED, style='Fun.TButton', command = self.email_sender)
            self.b33.grid(row=7, column=3)
        except AttributeError:
            pass
    
    def send_dar_toot(self):
        #ans = self.chk1.get()
        if self.cc_who == '0':
            self.cc_s.append('dar.lador18@gmail.com')
            self.cc_s.append('toot.moran@gmail.com')
        if self.cc_who == '1':
            mamagers = list(self.manager_inf[self.manager_inf['Specialty']==self.specialty].email.values)
            for m in mamagers:
                if m !=self.email_:
                    self.cc_s.append(m)
        if self.cc_who == '3':
            pass
        if len(self.cc_s)>0:
            self.cc_s = list(self.set(self.cc_s))
    def add_sp_manager(self):
        mamagers = list(self.manager_inf[self.manager_inf['Specialty']==self.specialty].email.values)
        for m in mamagers:
            if m !=self.email_:
                self.cc_s.append(m)
        self.cc_s = list(self.set(self.cc_s))
    def insert_pass(self):
        if self.my_var.get() != 'No one':
            self.top2 = Toplevel()
            self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
            self.email_ = self.manager_inf.loc[self.me, 'email'] if self.my_var.get()!='Dar SheCodes' else 'dar.shecodes@gmail.com' 
            self.lab111 = Label(self.top2, text="Email: "+self.email_)
            self.lab111.grid(row=0, column=0)
            self.lab222 = Label(self.top2, text="Email password: ")
            self.lab222.grid(row=1, column=0, sticky=W+E)
            self.entry222 = Entry(self.top2, show="*")
            self.entry222.grid(row=1, column=1, sticky=W+E)
            self.b2 = Button(self.top2, text= 'OK', command = self.pass_destroy_top2, style='Fun.TButton')
            self.b2.grid(row=2,column=0)
            self.b2 = Button(self.top2, text= 'Cancel', command = self.destroy_top2, style='Fun.TButton')
            self.b2.grid(row=2,column=1)
        
    def email_sender(self):
        body = self.body_email1.get('1.0', END)
        msg = (MIMEText(body))
        msg['Subject'] = self.subject11.get()
        if len(self.cc_s)>0:
            msg['Cc'] = self.cc_s
        server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)
        server.ehlo()
        server.starttls()
        username_ = re.search(r'([\w.]+)@([\w.]+)', self.email_).group(1)
        try:
            server.login(username_, self.my_password)
            recipients = []
            to_ = self.selected_names.split('\n')
            for her_name in to_:
                print(her_name)
                recipients.append(self.w[self.w.FullName==her_name].email[self.w[self.w.FullName == her_name].index[0]])
            server.sendmail(self.email_, 
                            recipients, 
                            msg.as_string())
            self.top1.destroy()
            self.top2 = Toplevel()
            self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
            self.conf = Label(self.top2, text='Your message was sent successfully to:')
            self.conf.grid(row=0,column=0)
            self.conf1 = tkst.ScrolledText(self.top2, wrap=WORD, width=10, height=3)
            self.conf1.grid(row=1,column=0,columnspan=3, sticky=N+E+S+W)
            self.conf1.insert(INSERT, self.selected_names)
            self.b2 = Button(self.top2, text= 'OK', command = self.destroy_top2, style='Fun.TButton')
            self.b2.grid(row=4,column=0)
        except smtplib.SMTPAuthenticationError:
            self.err_password_email()
    def err_password_email(self):
        self.top2 = Toplevel()
        self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.err_ = Label(self.top2, text="ERROR", font = "Verdana 10 bold")
        self.err_.grid(row=0, column=0)
        self.lab_wrong_pass = Label(self.top2, text="The user name or password is incorrect.")
        self.lab_wrong_pass2 = Label(self.top2, text="Please try again or select another account.")
        self.lab_wrong_pass.grid(row=1,column=0,columnspan=3,sticky=N+E+S+W)
        self.lab_wrong_pass2.grid(row=2,column=0,columnspan=3,sticky=N+E+S+W)
        self.email_ = self.manager_inf.loc[self.me, 'email'] if self.my_var.get()!='Dar SheCodes' else 'dar.shecodes@gmail.com' 
        self.lab111 = Label(self.top2, text="Email: "+self.email_)
        self.lab111.grid(row=3, column=0, sticky=W+E)
        self.lab222 = Label(self.top2, text="Email password: ")
        self.lab222.grid(row=4, column=0, sticky=W+E)
        self.entry222 = Entry(self.top2, show="*")
        self.entry222.grid(row=4, column=1, sticky=W+E)
        self.b2 = Button(self.top2, text= 'OK', command = self.pass_destroy_top2, style='Fun.TButton')
        self.b2.grid(row=5,column=0, sticky=W+E)
        self.b2 = Button(self.top2, text= 'Select another account', command = self.destroy_top2_deselect, style='Fun.TButton')
        self.b2.grid(row=5,column=1,columnspan=2, sticky=W+E)
    #=================================================================
    #                     (f4) complete lessons (Missing)
    #=================================================================
    
    def complete_lessons(self, f4):
        #self.df  includes a column Name and columns of every Wed from
        #the very first week of she codes academy. If a member has not
        #started on date X, the attendance will be written as "Didn't start yet".
        self.empty_erea0 = Label(f4, image=self.mini_shecodes)
        self.empty_erea0.grid(row=0, column=0, columnspan=20, sticky=S)
        self.empty_erea0.image = self.mini_shecodes 
        self.empty_erea0.configure(width=30)
        
        self.step1 = Label(f4, text='Step 1: choose a member', **self.lab_opt)
        self.step1.config(font=('times', 12, 'bold')) 
        self.step1.grid(row=1, column=0, columnspan=9, sticky=N+S)
        
        self.scroll_y = Scrollbar(f4, orient=VERTICAL)
        self.scroll_y.grid(row=2, rowspan =16, column=9, sticky=N+S)
        self.lst_b = Listbox(f4,
                             font=tkFont.Font(f4, family='Courier', size=14),
                             yscrollcommand=self.scroll_y.set,
                              selectmode=SINGLE, width=17)
        self.lst_b.grid(row=2, column=0, rowspan =16, columnspan=9, sticky=N+W+S+E)
        self.lst_b.bind('<<ListboxSelect>>', self.her_name)
        self.scroll_y.config(command=self.lst_b.yview)
        # adding members that missed lessons by using self.summery
        for name in self.df.Name:
            if name not in ['Dar']: 
                if self.summery[self.summery.Name==name]['# Absences'].values[0]!=0:
                    self.lst_b.insert(END, name)
            else:
                if self.me!='Dar Lador':
                    pass
                else:
                    if self.summery[self.summery.Name==name]['# Absences'].values[0]!=0:
                        self.lst_b.insert(END, name)
        
        self.empty_erea = Label(f4, text='                ').grid(row=2, column=10, columnspan=3, sticky=N+W+S+E)
        
        self.step2 = Label(f4, text='Step 2: Choose a meetup', **self.lab_opt)
        self.step2.config(font=('times', 12, 'bold')) 
        self.step2.grid(row=1, column=14, columnspan=1, sticky=N+S)
        self.meetups = {'Rehovot - WIS':[2,14],
                        'TLV - Check Point Building (Web)':[3,14], 
                        'TLV - Google Campus (Android)':[4,14],
                        'TLV - Quixey':[5,14], 
                        'TLV - Klarna':[6,14], 
                        'TLV - TAU':[7,14], 
                        'Jerusalem - American Center':[8,14],
                        'Jerusalem - Cisco':[9,14], 
                        'Jerusalem - HUJI':[10,14], 
                        'Beer Sheva - SCE':[11,14], 
                        'Beer Sheva - BGU':[12,14],
                        'Haifa - Technion':[13,14], 
                        'Haifa - Haifa University':[14,14], 
                        'Modiin - MESH':[15,14], 
                        'Herzelia - IDC':[16,14], 
                        'Herzelia - Microsoft Ventures':[17,14]}
        self.br = IntVar()
        self.br.set(len(self.meetups))
        for i,m in enumerate(self.meetups.keys()):
            self.choose_b = Radiobutton(f4, text=m, variable=self.br, value = i, command = self.get_meetup)
            self.choose_b.grid(row=self.meetups[m][0], column=self.meetups[m][1], sticky=W+E)
    
        self.step3 = Label(f4, text='Step 3: fill the date', **self.lab_opt)
        self.step3.config(font=('times', 12, 'bold')) 
        self.step3.grid(row=19, column=0, columnspan=9)
        
        self.empty_erea1 = Label(f4, text='                ').grid(row=18, column=9, columnspan=3, sticky=N+W+S+E)
        self.inst3 = Label(f4, text='Use the format: dd/mm/20yy')
        self.inst3.grid(row=20, column=0, columnspan=9, sticky=N+W+S+E)
        self.day = Entry(f4, width=3)
        self.day.grid(row=21, column=0, columnspan=1)
        self.sl_1 = Label(f4, text='/', width=1)
        self.sl_1.grid(row=21, column=1, columnspan=1)
        self.sl_1.config(font=('Courier', 14, 'bold')) 
        self.month = Entry(f4, width=3)
        self.month.grid(row=21, column=2, columnspan=1)
        self.sl_2 = Label(f4, text='/ 20', width=4)
        self.sl_2.grid(row=21, column=3, columnspan=1, sticky=W+N)
        self.sl_2.config(font=('Courier', 14, 'bold')) 
        self.year = Entry(f4, width=3)
        self.year.grid(row=21, column=4, columnspan=1, sticky=W)
        
        self.empty_erea2 = Label(f4, text='                ').grid(row=21, column=10, columnspan=3, sticky=N+W+S+E)
        
        self.step4 = Label(f4, text='Step 4: Submit', **self.lab_opt)
        self.step4.config(font=('times', 12, 'bold')) 
        self.step4.grid(row=19, column=14, columnspan=1)
        style = Style()
        style.configure('Fun1.TButton', foreground='#228b22', font=('Helvetica', 12, 'bold'))        

        self.submit_b = Button(f4, style='Fun1.TButton', text='Submit', command = self.check_completation)
        self.empty_erea2 = Label(f4, text='                ').grid(row=20, column=10, columnspan=3, sticky=N+W+S+E)
        self.got_new_tasks = IntVar()
        self.got_new_tasks.set(0)
        self.check1 = Radiobutton(f4, text='She got new tasks', variable=self.got_new_tasks, value=1,
                                  style ='TCheckbutton')
        self.check1.grid(row=20, column=14)
        self.submit_b.grid(row=21, column=14, columnspan=1)
        
        #self.empty_erea2 = Label(f4, text='                ').grid(row=20, column=15, columnspan=3, sticky=N+W+S+E)
        self.clear_all_button = Button(f4, text='Clear all', command=self.clear_all_area)
        self.clear_all_button.grid(row=18, column=18, columnspan=5, sticky=W+N)
        self.exit_f4 = Button(f4, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f4.grid(row=19, column = 18, rowspan=3, columnspan=5, sticky=W+N)
        
        #button1 = Button(self, text = "Send", command = self.response1, height = 100, width = 100) 
    def check_completation(self):
        her_name, meetup, day_, month = 'shecodes','shecodes','shecodes','shecodes'
        t_day, t_month, t_year = False, False, False
        try:
            her_name = self.lst_b.selection_get()
        except (AttributeError,TclError):
            self.checker_error_win("You must select a name!")
        if her_name !='shecodes':
            try:
                meetup = self.meetup
            except (AttributeError,TclError):
                self.checker_error_win("You must select 'she codes' branch!")
        if meetup !='shecodes':
            try:
                day_ = self.day.get()
                t_day = True
                try: 
                    int(day_)
                    if int(day_)>31:
                        t_day = False
                        self.checker_error_win("Invalid day: day is not exist!")
                    if len(day_)<2:
                        self.checker_error_win("You must insert a two-digit day!")
                        t_day = False
                except ValueError: 
                    t_day = False
                    self.checker_error_win("You must insert a two-digit day!")
                    t_day = False
            except AttributeError:
                self.checker_error_win("You must insert a two-digit day!")
                t_day = False
        if day_ !='shecodes' and t_day is True:
            try:
                month = self.month.get()
                t_month = True
                try: 
                    int(month)
                    if int(month)>12:
                        t_month = False
                        self.checker_error_win("Invalid month: month is not exist!")
                    if len(month)<2:
                        t_month = False
                        self.checker_error_win("You must insert a two-digit month!")
                except ValueError: 
                    t_month = False
                    self.checker_error_win("You must insert a two-digit number as a month!")
            except AttributeError:
                t_month = False
                self.checker_error_win("You must insert a two-digit month!")
        if month !='shecodes' and t_month is True:
            try:
                year = self.year.get()
                t_year = True
                try: 
                    int(year)
                    if len(year)>2 or len(year)<2:
                        t_year = False
                        self.checker_error_win("You must insert a two-digit year!")
                except ValueError: 
                    t_year = False
                    self.checker_error_win("You must insert a two-digit year!")
            except AttributeError:
                t_year = False
                self.checker_error_win("You must insert a two-digit year!")
        if her_name !='shecodes' and meetup !='shecodes' and \
            t_day is True and t_month is True and t_year is True:
            self.confirmit(name= her_name, dd = day_+'/'+month+'/20'+year, b = meetup)
            if self.to_do == 1:
                self.update_table_(name=her_name, branch=meetup, 
                                   full_date = day_+'/'+month+'/20'+year,
                                   tasks=self.got_new_tasks.get())
            
    def confirmit(self, name,dd,b):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.err_ = Label(self.top1, text="Note!", font = "Verdana 14 bold")
        self.err_.grid(row=0, column=0)
        self.lab_name = Label(self.top1, text='You are about to update the attendance of '+name)
        self.lab_name.grid(row=1,column=0)
        self.lab_name1 = Label(self.top1, text='on '+dd)
        self.lab_name1.grid(row=2,column=0)
        self.lab_name2 = Label(self.top1, text='in '+b)
        self.lab_name2.grid(row=3,column=0)
        if self.got_new_tasks.get()==1:
            self.lab_name3 = Label(self.top1, text='and she GOT new tasks.')
            self.lab_name3.grid(row=4,column=0)
        if self.got_new_tasks.get()==0:
            self.lab_name3 = Label(self.top1, text='and she DID NOT get new tasks.')
            self.lab_name3.grid(row=4,column=0)
        self.to_do=0
        self.ok_b = Button(self.top1, text= 'OK', command = lambda: self.update_table_(name=name, branch=b, 
                                   full_date = dd, tasks=self.got_new_tasks.get()), style='Fun.TButton')
        self.ok_b.grid(row=5,column=0, stick=W)
        self.cancel_b = Button(self.top1, text= 'Cancel', command = self.destroy_topl_reset_, style='Fun.TButton')
        self.cancel_b.grid(row=5, stick=E)
        
    def checker_error_win(self, mssg):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.err_ = Label(self.top1, text="ERROR", font = "Verdana 14 bold")
        self.err_.grid(row=0, column=0)
        self.lab_name = Label(self.top1, text=mssg, font = "Verdana 12")
        self.lab_name.grid(row=1,column=0)
        self.ok_b = Button(self.top1, text= 'OK', command = self.destroy_topl_reset_, style='Fun.TButton')
        self.ok_b.grid(row=2,column=0)
    
    def update_table_(self, name, branch, full_date, tasks):
        her_group = self.summery[self.summery.Name==name]['Group'].values[0]
        mo = re.search(r'([\w]+), ([1].)', her_group).group(1)
        yy = re.search(r'([\w]+), ([1].)', her_group).group(2)
        self.group_ = mo+yy
        full_data = branch+' '+full_date
        group_file = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        for i in list(group_file.loc[name].index):
            if isnull(group_file.loc[name,i]): 
                group_file.loc[name,i] = full_data
                break
        group_file.to_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        if tasks==1:
            for i in self.LessonState.loc[name].index:
                    if type(self.LessonState.loc[name].loc[i])!=str:
                        self.lesson = i
                        break
            self.LessonState.loc[name, self.lesson]=full_data
            self.LessonState.to_pickle(os.path.join(self.path,'LessonState.dat'))
            excel_saver(self.LessonState, 'LessonState')
        self.top1.destroy()
        
        self.top2 = Toplevel()
        self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        which_files=Label(self.top2, text='Update succeeded. Put this/these file/s in Google Drive:')
        which_files.grid(row=0, column=0)
        filenames = Label(self.top2, text=self.group_+'.dat' if tasks==0 else self.group_+'.dat + LessonState.dat + LessonState.xlsx')
        filenames.grid(row=1, column=0)
        confirm_mm = Label(self.top2, text='By clicking OK an email will be send to WIS managers')
        confirm_mm.grid(row=2, column=0)
        confirm_mm1 = Label(self.top2, text='in order to notify them to download those files.')
        confirm_mm1.grid(row=3,column=0)
        self.ok_b1 = Button(self.top2, text= 'OK', command = lambda: self.notify_all_(name), style='Fun.TButton')
        self.ok_b1.grid(row=4,column=0)
    
    def notify_all_(self, name):
        self.top2.destroy()
        #p_body = 'IGNORE THIS MESSAGE!!!!\n\n'
        p_body = 'Dear All, \n\nI updated the presence of '+name+' from '+self.specialty+' track.\n' 
        body='Please download the following file/s:\n '
        files = self.group_+'.dat' if self.got_new_tasks.get()==0 else self.group_+'.dat + LessonState.dat + LessonState.xlsx\n'
        end = '\nBest,\n'+self.me
        msg = (MIMEText(p_body+body+files+end))
        msg['Subject'] = '***She codes WIS*** Files were updated'
        self.email_ = 'dar.shecodes@gmail.com' if self.me == 'Dar Lador' else self.manager_inf.loc[self.me, 'email'] 
        server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)
        server.ehlo()
        server.starttls()
        username_ = re.search(r'([\w.]+)@([\w.]+)', self.email_).group(1)
        if self.got_new_tasks.get()==0:
            recipients = list(self.manager_inf[self.manager_inf.Specialty.isin([self.specialty, 'manager'])].email.values)
        elif self.got_new_tasks.get()==1:
            recipients = list(self.manager_inf.email.values)
        
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.lab222 = Label(self.top1, text="Email password: ")
        self.lab222.grid(row=1, column=0, sticky=W+E)
        self.entry222 = Entry(self.top1, show="*")
        self.entry222.grid(row=1, column=1, sticky=W+E)
        self.b2 = Button(self.top1, text= 'OK', command = lambda: self.get_passs_send(s = server, u=username_,
                                                                                      email = self.email_, 
                                                                                      to=recipients, 
                                                                                      msg=msg), style='Fun.TButton')
        self.b2.grid(row=2,column=0, sticky=W+E)
        
    #=================================================================
    #                     (f5) Delete/ restore a member
    #=================================================================
    def delete_restore_member(self, f5):
        self.note_pr_delete = Label(f5, text='Note: Contact branch admin in order to permanently delete a member.')
        self.note_pr_delete.grid(row=11, column=0, columnspan=30, sticky=N+S)
        self.note_pr_delete.config(font=('times', 14, 'bold')) 
        
        #------------------------
        # Delete members buttons
        #------------------------
        self.delete_ins = Label(f5, text='Choose a member to delete', **self.lab_opt)
        self.delete_ins.config(font=('times', 12, 'bold')) 
        self.delete_ins.grid(row=12, column=0, columnspan=9, sticky=N+S)
        
        self.sc_y_del = Scrollbar(f5, orient=VERTICAL)
        self.sc_y_del.grid(row=13, rowspan =16, column=9, sticky=N+S)
        self.lst_del = Listbox(f5,
                             font=tkFont.Font(f5, family='Courier', size=14),
                             yscrollcommand=self.sc_y_del.set,
                              selectmode=SINGLE, width=17)
        self.lst_del.grid(row=13, column=0, rowspan =16, columnspan=9, sticky=N+W+S+E)
        self.lst_del.bind('<<ListboxSelect>>', self.delete_her)
        self.sc_y_del.config(command=self.lst_del.yview)
        # adding members by using self.summery
        for name in self.df.Name:
            self.lst_del.insert(END, name)
                
        self.empty_erea_1 = Label(f5, text='                ').grid(row=1, column=10, columnspan=3, sticky=N+W+S+E)
        style = Style()
        style.configure('Fun2.TButton', foreground='#ff0000', font=('Helvetica', 12, 'bold'))
        self.delete_button = Button(f5, text='Delete', command=self.sure_delete, style='Fun2.TButton')
        self.delete_button.grid(row=30, column=0, columnspan=9, sticky=N+W+S+E)
        self.delete_button.config(width=17)
        
        #-------------------------
        # Restore a member button
        #-------------------------
        self.restore_ins = Label(f5, text='Choose a member to restore', **self.lab_opt)
        self.restore_ins.config(font=('times', 12, 'bold')) 
        self.restore_ins.grid(row=12, column=11, columnspan=11, sticky=N+S)
        
        self.sc_y_res = Scrollbar(f5, orient=VERTICAL)
        self.sc_y_res.grid(row=13, rowspan =16, column=22, sticky=N+S)
        self.lst_res = Listbox(f5,
                             font=tkFont.Font(f5, family='Courier', size=14),
                             yscrollcommand=self.sc_y_res.set,
                              selectmode=SINGLE, width=20)
        self.lst_res.grid(row=13, column=11, rowspan =16, columnspan=11, sticky=N+W+S+E)
        self.lst_res.bind('<<ListboxSelect>>', self.restore_her)
        self.sc_y_res.config(command=self.lst_res.yview)
        
        
        # adding members by using self.summery#KeyError: 'the label [0] is not in the [index]'
        for name in self.former_att.index:
            if self.former_inf.loc[name, 'Track'] in self.track:
                self.lst_res.insert(END, name)
                
        self.empty_erea_ = Label(f5, text='                ').grid(row=1, column=10, columnspan=3, sticky=N+W+S+E)
        style = Style()
        style.configure('Fun3.TButton', foreground='#228b22', font=('Helvetica', 12, 'bold'))
        style.configure('Fun4.TButton', font=('Helvetica', 12, 'bold'))
        self.delete_button = Button(f5, text='Restore', command=self.sure_restore, style='Fun3.TButton')
        self.delete_button.grid(row=30, column=11, columnspan=11, sticky=N+W+S+E)
        self.delete_button.config(width=17)
        
        #------------
        # exit button
        #------------
        self.empty_erea_ = Label(f5, text='                ').grid(row=1, column=22, columnspan=3, sticky=N+W+S+E)
        self.exit_f5 = Button(f5, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f5.grid(row=27, column = 25, rowspan=5, columnspan=5, sticky=W+N)
    
    #-----------------
    # Delete a member
    #-----------------
    def sure_delete(self):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        try:
            self.ask_d = Label(self.top1, text='Are you sure you want to delete '+self.delete_name+'?', font = "Verdana 12")
            self.ask_d.pack(side=TOP, anchor=W, fill=X, expand=YES)
            self.ask_b_y = Button(self.top1, text='Yes', command=self.delete_from_files, style='Fun4.TButton')
            self.ask_b_y.pack(side=LEFT, anchor=E, expand=YES)
            self.ask_b_y = Button(self.top1, text='No', command=self.destroy_top1, style='Fun4.TButton')
            self.ask_b_y.pack(side=RIGHT, anchor=W, expand=YES)
        except AttributeError:
            self.ask_d1 = Label(self.top1, text='ERROR', font = "Verdana 14 bold").pack()
            self.ask_d = Label(self.top1, text='No member name has been given!', font = "Verdana 12").pack()
            self.close_b1 = Button(self.top1, text='Close', command=self.destroy_top1).pack()
    
    def delete_from_files(self):
        # adding her name to FormerMembersInf + FormerMembersAtt
        summ = self.summery.set_index('Name')
        dff = self.df.set_index('Name')
        self.former_inf = self.former_inf.append(summ.loc[self.delete_name][['Track', 'Group', '# Lesson']].to_frame().T)
        self.former_att = self.former_att.append(dff.loc[self.delete_name].to_frame().T)
        
        # search her group file and delete her from it but save her att
        her_group = self.summery[self.summery.Name==self.delete_name]['Group'].values[0]
        mo = re.search(r'([\w]+), ([1].)', her_group).group(1)
        yy = re.search(r'([\w]+), ([1].)', her_group).group(2)
        self.group_ = mo+yy
        group_file = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        
        #save att and other files
        self.former_attlist = self.former_attlist.append({'Name': self.delete_name, 
                                                          'LessonState': self.LessonState.loc[self.delete_name].values[0:int(summ.loc[self.delete_name, '# Lesson'])]},
                                                         ignore_index='True')
        self.former_attlist.to_pickle(os.path.join(self.path, 'FormerLessonState.dat'))
        group_file = group_file.drop(self.delete_name)
        group_file.to_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        self.former_inf.to_pickle(os.path.join(self.path, 'FormerMembersInf.dat'))
        self.former_att.to_pickle(os.path.join(self.path, 'FormerMembersAtt.dat'))
        # delete from LessonState
        self.LessonState = self.LessonState.drop(self.delete_name)
        self.LessonState.to_pickle(os.path.join(self.path, 'LessonState.dat'))
        excel_saver(self.LessonState, 'LessonState')
        self.top1.destroy()
        self.notify_files_changed(what_i_did='remove ', who=self.delete_name, subj='A member was deleted')
        
    #------------------
    # Restore a member
    #------------------
    def sure_restore(self):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        try:
            self.ask_d = Label(self.top1, text='Are you sure you want to bring '+self.restore_name+' back?', font = "Verdana 12")
            self.ask_d.pack(side=TOP, anchor=W, fill=X, expand=YES)
            self.ask_b_y = Button(self.top1, text='Yes', command=self.restore_to_files, style='Fun4.TButton')
            self.ask_b_y.pack(side=LEFT, anchor=E, expand=YES)
            self.ask_b_y = Button(self.top1, text='No', command=self.destroy_top1, style='Fun4.TButton')
            self.ask_b_y.pack(side=RIGHT, anchor=W, expand=YES)
        except AttributeError:
            self.ask_d1 = Label(self.top1, text='ERROR', font = "Verdana 14 bold").pack()
            self.ask_d = Label(self.top1, text='No member name has been given!', font = "Verdana 12").pack()
            self.close_b1 = Button(self.top1, text='Close', command=self.destroy_top1).pack()
          
    def restore_to_files(self):
        # add her name to LessonState + group_file
        self.LessonState = self.LessonState.append(self.former_att.loc[self.restore_name].to_frame().T)
        self.LessonState.to_pickle(os.path.join(self.path, 'LessonState.dat'))
        excel_saver(self.LessonState, 'LessonState')
        her_group = self.former_inf[self.former_inf.loc[self.restore_name, 'Starting month']]
        mo = re.search(r'([\w]+), ([1].)', her_group).group(1)
        yy = re.search(r'([\w]+), ([1].)', her_group).group(2)
        self.group_ = mo+yy
        group_file = read_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        group_file.loc[self.restore_name] = np.nan
        for col in group_file.columns:
            if col in list(self.former_attlist[self.former_attlist.Name==self.restore_name].atts.values)[0]:
                group_file[self.restore_name, col] = 'V'
        group_file.to_pickle(os.path.join(os.path.join(self.path, self.folder_track) ,self.group_+'.dat'))
        former_att = self.former_att.set_index('Name')
        
        # delete her name from FormerMembersInf + FormerMembersAtt + FormerAttList
        former_att = former_att.drop(self.restore_name)
        former_att = former_att.reset_index()
        former_att.to_pickle(os.path.join(self.path, 'FormerAttList.dat'))
        self.former_inf = self.former_inf.drop(self.restore_name)
        self.former_inf.to_pickle(os.path.join(self.path, 'FormerMembersInf.dat'))
        self.former_att = self.former_att.drop(self.restore_name)
        self.former_att.to_pickle(os.path.join(self.path, 'FormerMembersAtt.dat'))
       
        self.top1.destroy()
        self.notify_files_changed(what_i_did='brought back ', who=self.restore_name, subj='A member was returned')
    
    #-------------------------------------
    # Notify a member was deleted/restored
    #-------------------------------------
    def notify_files_changed (self, what_i_did, who, subj):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.finish_mssg = Label(self.top1, text="Move the following files into Google Drive > 'she codes WIS':",
                                 font = "Verdana 12").pack()
        self.file1 = Label(self.top1, text='FormerMembersInf.dat', font = "Verdana 12 bold").pack()
        self.file2 = Label(self.top1, text = 'FormerMembersAtt.dat', font = "Verdana 12 bold").pack()
        self.file3 = Label(self.top1, text = 'LessonState.dat', font = "Verdana 12 bold").pack()
        self.file4 = Label(self.top1, text = 'FormerAttList.dat', font = "Verdana 12 bold").pack()
        self.finish_mssg = Label(self.top1, text='Move the following file into '+self.folder_track+' folder:',
                                 font = "Verdana 12").pack()
        self.file3 = Label(self.top1, text=self.group_+'.dat', font = "Verdana 12 bold")
        self.finish_mssg2 = Label(self.top1, text='By clicking OK a message will be sent to the other admins',
                                 font = "Verdana 12").pack()
        self.finish_mssg3 = Label(self.top1, text='in order to notify them to download those files.',
                                 font = "Verdana 12").pack()
        self.OK_del_res = Button(self.top1, text='OK', style='Fun4.TButton', 
                                 command=lambda: self.email_notify_changed_members(what_i_did, who, subj)).pack()
    
    def email_notify_changed_members(self, what_i_did, who, subj):
        self.top1.destroy()
        to_or_from = ' to ' if what_i_did=='brought back ' else ' from '
        p_body = "Dear All, \n\nI've just "+ what_i_did + who + to_or_from + self.specialty +' track.\n'
        body='Please update the following files:\n'
        files = 'FormerMembersInf.dat & FormerMembersAtt.dat & LessonState.dat & FormerAttList.dat\n'
        file = 'Also move '+self.group_+'.dat file to '+self.folder_track+' folder.'
        end = '\nBest,\n'+self.me
        msg = (MIMEText(p_body+body+files+file+end))
        msg['Subject'] = '***She codes WIS*** '+subj
        self.email_ = 'dar.shecodes@gmail.com' if self.me == 'Dar Lador' else self.manager_inf.loc[self.me, 'email'] 
        server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)
        server.ehlo()
        server.starttls()
        username_ = re.search(r'([\w.]+)@([\w.]+)', self.email_).group(1)
        if self.got_new_tasks.get()==0:
            recipients = list(self.manager_inf[self.manager_inf.Specialty==self.specialty].email.values)
        elif self.got_new_tasks.get()==1:
            recipients = list(self.manager_inf.email.values)
        
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.lab222 = Label(self.top1, text="Email password: ")
        self.lab222.grid(row=1, column=0, sticky=W+E)
        self.entry222 = Entry(self.top1, show="*")
        self.entry222.grid(row=1, column=1, sticky=W+E)
        self.b2 = Button(self.top1, text= 'OK', command = lambda: self.get_passs_send(s = server, u=username_,
                                                                                      email = self.email_, 
                                                                                      to=recipients, 
                                                                                      msg=msg), style='Fun.TButton')
        self.b2.grid(row=2,column=0, sticky=W+E)
        
    
    #=================================================================
    #                     (f6) Send weekly email
    #=================================================================
    def send_weekly_email(self, f6):
        self.empty_erea6 = Label(f6, image=self.mini_shecodes)
        self.empty_erea6.grid(row=0, column=0, columnspan=20, sticky=S)
        self.empty_erea6.image = self.mini_shecodes 
        self.empty_erea6.configure(width=30)
        self.lesson_my_track = self.LessonState.loc[self.LessonState.index.isin(self.w_a.FullName)]
        self.lesson_my_track = self.lesson_my_track.loc[~self.lesson_my_track.index.isin(self.debug_dar)]
        
        date_w = self.first_wed_of_month(date.today())
        self.date_w = str(date_w.day)+'/'+str(date_w.month)+'/'+str(date_w.year)
        self.date_w1 = date_w.strftime('%d/%m/%Y')
        self.to_whom = StringVar()
        
        self.to_whom.trace("w", self.get_emails_members)
        
        self.label_to_mem = Label(f6, text='Send to:', font=('times', 12, 'bold'))
        self.label_to_mem.grid(row=1, column=0, rowspan = 1, columnspan=1,sticky=W)
        self.add_lessons = 'No member has been chosen'
        self.num_lessons = Text(f6, width=70, height=1)
        self.num_lessons.insert(END, self.add_lessons)
        self.num_lessons.config(state=DISABLED)
        self.select_members = OptionMenu(f6, self.to_whom, *['Send to...',
                                                             'Whoever was attended our last meetup',
                                                             'Whoever was attended during the last 2 weeks',
                                                             'Whoever was attended during the last 3 weeks',
                                                             'Whoever was attended during the last month'])
        self.select_members.grid(row=1, column=1, rowspan = 1, columnspan=4,sticky=W)
        style = Style()
        style.configure("BW.TMenubutton", foreground="white", background="green", font=('times', 11, 'bold'))
        self.select_members.configure(width=40)
        self.select_members.configure(style = "BW.TMenubutton")
        self.lesson_label = Label(f6, text='They are currently work on lessons: ', font=('times', 12, 'bold'))
        self.lesson_label.grid(row=3, column=0, rowspan = 1, columnspan=3, sticky=W)
        self.lesson_label.configure(width=35)
        self.who_details = Button(f6, text='Recipients details', style='Fun4.TButton', command = self.show_mdetails)
        self.who_details.grid(row=3, column=4, rowspan = 1, columnspan=2, sticky=W)
        self.who_details.config(state=DISABLED)
        self.num_lessons.grid(row=4, column=0, rowspan = 1, columnspan=12,sticky=W)
        self.to_whom.set('Send to...')
        
        self.to_names_ = tkst.ScrolledText(f6, wrap=WORD, width=50, height=3)
        
        self.lab33 = Label(f6, text = 'Email body:', font=('times', 12, 'bold'), foreground="green")
        self.lab33.grid(row=7, column=0, columnspan=1, sticky=W)
        self.body_email1 = tkst.ScrolledText(f6, wrap=WORD, width=70, height=17)
        self.body_email1.grid(row=8, column=0, columnspan=12, sticky=W)
        
        self.last_wed = date.today()+dtime.timedelta(2-date.today().weekday())
        self.last_wed = self.last_wed.strftime('%d/%m/%Y')
        self.sub = '>>>She codes<<< '
        self.sub_edited_part = Entry(f6, width=44, font=('times', 12))
        self.sub_edited_part.insert(END, 'Tasks remainder ('+self.last_wed+')')
        self.sub_edited_part.grid(row=5, column=2, rowspan = 1, columnspan=4,sticky=W)
        self.subject1 = Label(f6, text='Subject:     ', font=('times', 12, 'bold'))
        self.subject1.grid(row=5,column=0, columnspan=1, sticky=W)
        self.sub1 = Label(f6, text=self.sub, font=('times', 12))
        self.sub1.grid(row=5,column=1, columnspan=1, sticky=W)
        
        self.b33 = Button(f6, text= 'Send', style='Fun3.TButton', command = self.week_insert_pass)
        self.b33.grid(row=10, column=1)
        self.exit_f6 = Button(f6, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f6.grid(row=10, column = 5, rowspan=5, columnspan=5, sticky=W+N)
    
    def get_emails_members(self, *args):
        wanted = self.to_whom.get()
        if wanted == 'Send to...':
            try:
                self.num_lessons.config(state=NORMAL)
                self.num_lessons.delete(1.0, END)
            except TclError:
                pass
            self.num_lessons.insert(END, self.add_lessons)
            self.num_lessons.config(state=DISABLED)
        else:
            d = self.df.set_index('Name')
            d = d.loc[~d.index.isin(self.debug_dar)]
            for n in self.lesson_my_track.index:
                if n not in d.index:
                    if n not in list(self.former_inf.index):
                        d.loc[n, self.date_w1]='V'
            attn = {'Whoever was attended our last meetup':-1, 'Whoever was attended during the last 2 weeks':-2, 
                    'Whoever was attended during the last 3 weeks': -3, 'Whoever was attended during the last month':-5}
            self.req = d[~isnull(d.ix[:,attn[wanted]:]).all(1)]
            track_sp_ = self.lesson_my_track.loc[self.lesson_my_track.index.isin(self.req.index)]
            summ = self.summery.set_index('Name')
            self.state_emailed_members = summ.loc[summ.index.isin(self.req.index)]
            self.state_emailed_members = self.state_emailed_members[['# Lesson', 'Group', 'Track']]
            self.new_members = [new_m for new_m in track_sp_.index if new_m not in self.state_emailed_members.index]
            for new in self.new_members:
                self.state_emailed_members.loc[new, '# Lesson']=1
                self.state_emailed_members.loc[new, 'Group']=date.today().strftime('%b, %y')
                self.state_emailed_members.loc[new, 'Track']=self.track[3]
            learning_state = [str(int(i)) for i in list(set(self.state_emailed_members['# Lesson'].values))]
            self.state_emailed_members = self.state_emailed_members.sort_values(by='# Lesson')
            self.sub_state_emailed_members = self.state_emailed_members.reset_index()
            cols_s = list(self.sub_state_emailed_members)
            row_s = range(len(self.sub_state_emailed_members))
            self.sub_state_emailed_members = self.sub_state_emailed_members.ix[row_s, cols_s]
            self.sub_show_s = self.sub_state_emailed_members.to_string(index=False, col_space=10)
            
#             for col in list(track_sp_.columns)[1:]:
#                 if track_sp_.dropna(subset=[col]).shape[0]<track_sp_.dropna(subset=[col-1]).shape[0]:
#                     learning_state.append(str(col-1))
            self.add_lessons = ', '.join(learning_state)
            self.num_lessons.config(state=NORMAL)
            self.num_lessons.delete(1.0, END)
            self.num_lessons.insert(END, self.add_lessons)
            self.num_lessons.config(state=DISABLED)
            self.who_details.config(state=NORMAL)
    
    def show_mdetails(self):
        self.top1 = Toplevel()
        self.top1.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.label_mdetails = Label(self.top1, text='Recipients details', font = "Verdana 10 bold")
        self.label_mdetails.grid(row=0,column=2)
        self.scroll_details = tkst.ScrolledText(self.top1, wrap=WORD, width=65, height=15)
        self.scroll_details.grid(row=1,column=0, rowspan=10, columnspan=5, sticky=N+E+S+W)
        self.scroll_details.insert(INSERT, self.sub_show_s)
        self.clost_details = Button(self.top1, text='Close', command=self.destroy_top1)
        self.clost_details.grid(row=11,column=2)
        
    def send_weekly_e(self):
        self.my_password = self.entry222.get()
        self.top2.destroy()
        body = self.body_email1.get('1.0', END)
        msg = MIMEText(body, 'plain')
        msg['Subject'] = self.sub+self.sub_edited_part.get()
        cc_ = list(self.manager_inf[self.manager_inf.Specialty==self.specialty].email.values) + list(self.manager_inf[self.manager_inf.Specialty=='manager'].email.values)
        cc_l = ', '.join(cc_)
        msg['Cc'] = cc_l
        server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)
        server.ehlo()
        server.starttls()
        username_ = re.search(r'([\w.]+)@([\w.]+)', self.email_).group(1)
        try:
            server.login(username_, self.my_password)
            recipients = []
            for her_name in self.req.index:
                recipients.append(self.w[self.w.FullName==her_name].email[self.w[self.w.FullName == her_name].index[0]])
            server.sendmail(self.email_, 
                            recipients + cc_, 
                            msg.as_string())
            self.top2 = Toplevel()
            self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
            self.conf = Label(self.top2, text='Your message was sent successfully to:')
            self.conf.grid(row=0,column=0)
            self.conf1 = tkst.ScrolledText(self.top2, wrap=WORD, width=10, height=3)
            self.conf1.grid(row=1,column=0,columnspan=3, sticky=N+E+S+W)
            self.conf1.insert(INSERT, self.req.index)

            self.b2 = Button(self.top2, text= 'OK', command = self.destroy_top2, style='Fun.TButton')
            self.b2.grid(row=4,column=0)
        except smtplib.SMTPAuthenticationError:
            self.send_weekly_error()
            
    def send_weekly_error(self):
        self.top2 = Toplevel()
        self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.err_ = Label(self.top2, text="ERROR", font = "Verdana 10 bold")
        self.err_.grid(row=0, column=0)
        self.lab_wrong_pass = Label(self.top2, text="The user name or password is incorrect.")
        self.lab_wrong_pass2 = Label(self.top2, text="Please try again or select another account.")
        self.lab_wrong_pass.grid(row=1,column=0,columnspan=3,sticky=N+E+S+W)
        self.lab_wrong_pass2.grid(row=2,column=0,columnspan=3,sticky=N+E+S+W)
        self.email_ = self.manager_inf.loc[self.me, 'email'] if self.my_var.get()!='Dar SheCodes' else 'dar.shecodes@gmail.com' 
        self.lab111 = Label(self.top2, text="Email: "+self.email_)
        self.lab111.grid(row=3, column=0, sticky=W+E)
        self.lab222 = Label(self.top2, text="Email password: ")
        self.lab222.grid(row=4, column=0, sticky=W+E)
        self.entry222 = Entry(self.top2, show="*")
        self.entry222.grid(row=4, column=1, sticky=W+E)
        self.b2 = Button(self.top2, text= 'OK', command = self.send_weekly_e, style='Fun.TButton')
        self.b2.grid(row=5,column=0, sticky=W+E)
        
    def week_insert_pass(self):
        self.top2 = Toplevel()
        self.top2.iconbitmap(os.path.join(self.path, 'CookieMonster.ico'))
        self.email_ = self.manager_inf.loc[self.me, 'email']
        self.lab111 = Label(self.top2, text="Email: "+self.email_)
        self.lab111.grid(row=0, column=0)
        self.lab222 = Label(self.top2, text="Email password: ")
        self.lab222.grid(row=1, column=0, sticky=W+E)
        self.entry222 = Entry(self.top2, show="*")
        self.entry222.grid(row=1, column=1, sticky=W+E)
        self.b2 = Button(self.top2, text= 'OK', command = self.send_weekly_e, style='Fun.TButton')
        self.b2.grid(row=2,column=0)
    
    #=================================================================
    #                     (f7) Dar's Notes
    #=================================================================
    def print_notes(self, f7):
        self.notes_win = tkst.ScrolledText(f7, wrap=WORD, width=70, height=16)
        self.notes_win.grid(row=8, column=0, columnspan=12, sticky=W)
        self.notes_win.insert(INSERT, """\
        
        Please move the following files to the Drive:
        1) attlist.dat - if it was changed manually
        2) LessonState.dat - if it was changed manually 
                             OR members complete lessons and got tasks.
        3) Lessonstate.excel - if a member was deleted/restored
                                OR members complete lessons *and* got tasks.
        3) track_contacts.xlsx - if it was changed manually
        4) FormerMembersInf.xlsx - every week
        """)
        self.exit_f7 = Button(f7, compound=TOP, text='Quit', image=self.exitdoor_photo, command=self.quit)
        self.exit_f7.grid(row=10, column = 11, rowspan=3, columnspan=5, sticky=NW)
        
    def quit(self):
        tx = 'FormerMembersInf.xlsx'
        writer = ExcelWriter(os.path.join(os.path.join(self.path, 'SavedFiles'),tx), engine='xlsxwriter')
        self.former_inf.to_excel(writer, sheet_name='Sheet1')
        self.root.quit()
    def destroy_top1(self):
        self.top1.destroy()
    def pass_destroy_top2(self):
        if len(self.entry222.get()) >1:
            self.my_password = self.entry222.get()
            self.b33.config(state="normal")
        self.top2.destroy()
    def destroy_top2(self):
        self.top2.destroy()
    def destroy_top2_deselect(self):
        self.b33.config(state=DISABLED)
        try:
            self.from_.config(state='normal')
        except AttributeError:
            self.from_D.config(state='normal')
        self.top2.destroy()
    def get_meetup(self):
        self.meetup = list(self.meetups.keys())[self.br.get()]
    def her_name(self, *args):
        self.her_name_ = self.lst_b.selection_get()
    def destroy_topl_reset_(self):
        try: self.day.delete(0, 'end')
        except AttributeError: pass
        try: self.month.delete(0, 'end')
        except AttributeError: pass
        try: self.year.delete(0, 'end')
        except AttributeError: pass
        self.top1.destroy()
    def clear_all_area(self):
        try: self.day.delete(0, 'end')
        except AttributeError: pass
        try: self.month.delete(0, 'end')
        except AttributeError: pass
        try: self.year.delete(0, 'end')
        except AttributeError: pass
        self.br.set(len(self.meetups))
        self.got_new_tasks.set(0)
        self.lst_b.selection_clear(0, END)
    def change_to_do(self):
        self.to_do=1
        self.top1.destroy()
    def get_passs_send(self, s, u, email, to, msg):
        self.my_password = self.entry222.get()
        try:
            s.login(u, self.my_password)
            s.sendmail(self.email_, to, msg.as_string())
            try:
                self.top1.destroy()
            except:
                pass
            self.root.destroy()
            main(self.path, self.specialty,self.me,restart=True)
        except smtplib.SMTPAuthenticationError:
            self.err_password_email()   
    def delete_her(self, *args):
        self.delete_name = self.lst_del.selection_get()
    def restore_her(self, *args):
        self.restore_name = self.lst_res.selection_get()
    def first_wed_of_month(self, d):
        for i in range(1, 8):
            if date(d.year, d.month, i).weekday()==2:
                return date(d.year, d.month, i)
##################
# ADDING WIDGETS #
##################
    def _fill(self, f):
        self._pack_bind_lb()
        self._fill_listbox()
        
    
##############
# SCROLLBARS #
##############
    def _init_scroll(self, f):
        self.scrollbar  = Scrollbar(f, orient="vertical")
        self.xscrollbar = Scrollbar(f, orient="horizontal")


    def _onMouseWheel(self, event):
        self.title_lb.yview("scroll", event.delta,"units")
        self.lbSites.yview("scroll", event.delta,"units")
        return "break"

    def _xview(self, *args):
        """connect the yview action together"""
        self.lbSites.xview(*args)
        self.title_lb.xview(*args)
    
    def _xview2(self, *args):
        """connect the yview action together"""
        self.lbSites2.xview(*args)
        self.title_lb2.xview(*args)
#     
#     def _yview(self, *args):
#         self.lbSites.yview(*args)
#         self.title_lb.yview(*args)


    def _pack_bind_lb(self):
        self.title_lb.bind("<MouseWheel>", self._onMouseWheel)
        self.lbSites.bind("<MouseWheel>", self._onMouseWheel)

    def _fill_listbox(self):
        """ fill the listbox with rows from the dataframe"""
        self.title_lb.insert(END, self.title_string)
        for line in self.sub_datstring[1:]:
            self.lbSites.insert(END, line) 
            self.lbSites.bind('<ButtonRelease-1>',self._listbox_callback)
        self.lbSites.select_set(0)

    def _listbox_callback(self, event):
        """ when a listbox item is selected"""
        pass

##################
# UPDATING LINES #
##################
    def _rewrite(self): 
        """ re-writing the dataframe string in the listbox"""
        new_col_vals = self.summery.ix[ self.row , self.dat_cols ].astype(str).tolist() 
        new_line     = self._make_line( new_col_vals ) 
        if self.lbSites.cget('state') == DISABLED:
            self.lbSites.config(state=NORMAL)
            self.lbSites.delete(self.idx)
            self.lbSites.insert(self.idx,new_line)
            self.lbSites.config(state=DISABLED)
        else:
            self.lbSites.delete(self.idx)
            self.lbSites.insert(self.idx,new_line)

    def _get_line_format(self, line) :
        """ save the format of the title string, stores positions
            of the column breaks"""
        pos = [1+line.find(' '+n)+len(n) for n in self.dat_cols]
        self.entry_length = [pos[0]] + [ p2-p1 for p1,p2 in zip(  pos[:-1], pos[1:] ) ]

    def _make_line( self , col_entries):
        """ add a new line to the database in the correct format"""
        new_line_entries = [ ('{0: >%d}'%self.entry_length[i]).format(entry)  
                            for  i,entry in enumerate(col_entries) ] 
        new_line = "".join(new_line_entries)
        return new_line

def main(my_path, track, me, restart=True):
    my_path = r'C:/Users/Dar1/Documents/she_codes'
    #start
    root = Tk()
    root.iconbitmap(os.path.join(my_path, 'CookieMonster.ico'))
    if restart:
        ManagerView(root, my_path, track, me)
        root.mainloop()
    #re-assign dataframe
    
    
    
if __name__ == '__main__':
    main(my_path = r'C:/Users/me',
         track='web', me='Dar Lador')
    #my_path = r'C:/Users/me'
    #root = Tk()
#     root.iconbitmap(os.path.join(my_path, 'CookieMonster.ico'))
#     b = ManagerView(root, my_path, 'android')
#     root.mainloop()