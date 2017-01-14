'''
Created on Feb 7, 2016

@author: Dar1
'''
import pandas as pd
import os
import re
import time
from time import mktime
from datetime import datetime, timedelta
import numpy as np
def create_lesson_state_file(path):
    df = pd.read_excel(os.path.join(path, 'WISAcademy.xlsx'), sheetname = 'Sheet1', header = 0)
    df.to_pickle(os.path.join(path, 'LessonState.dat'))

def create_WISAcademy(path):
    try:
        att = pd.read_pickle(os.path.join(path, 'attendance.dat'))
    except IOError:
        att = pd.read_pickle(os.path.join(path, 'LessonState.dat'))
    writer = pd.ExcelWriter(os.path.join(path, 'WISAcademy.xlsx'), engine = 'xlsxwriter')
    att.to_excel(writer, sheet_name = 'Sheet1')
    writer.save()
    
def create_attfiles(path, track, d_start = '2015-08-05', d_end = '2016-02-03', 
                    filename = 'Aug15.dat', sheetname_ = 'Aug, 2015'):
    attwis = pd.read_excel(os.path.join(path, 'AttWis.xlsx'), sheetname=sheetname_, header = 0)
    w = pd.read_excel(os.path.join(path, 'track_contacts.xlsx'), sheetname = 'tracks', header = 0)
    w['Start']=w.apply(lambda row: row['Start'].strftime('%b, %Y'),1)
    datelist = pd.date_range(start=d_start, end=d_end, freq='7D', closed='left').tolist()
    datelist = [x.strftime('%d/%m/%Y') for x in datelist]
    df = pd.DataFrame(columns=datelist)
    android = w.loc[w['track'].isin(['android', 'android_new'])]
    web = w.loc[w['track'].isin(['web', 'web_new'])]
    saving_path = path+'/'+'android_' if track == 'android' else path+'/'+'web_'
    if not os.path.exists(saving_path):
        os.makedirs(saving_path)
    attwis = attwis.loc[attwis.index.isin(android.FullName)] if track == 'android' else \
            attwis.loc[attwis.index.isin(web.FullName)]
    for i in attwis.index:
        df = df.append(attwis.loc[i])
    df.to_pickle(os.path.join(saving_path, filename))

def managers_editor(path):
    df = pd.DataFrame({'email': ['1@gmail.com', '2@gmail.com', '3.shecodes@gmail.com',
                                 '3@gmail.com', '4@gmail.com'], 
                       'Password': ['shecodes', 'shecodes', 'shecodes', 'shecodes', 'shecodes'],
                       'Specialty':['web', 'android', 'android', 'manager', 'manager']}, 
                      index = ['Manager1', 'Manager2', 'Manager3', 'Manager4', 'Dar Lador'])
    df.index.name = 'Name'
    df.to_pickle(os.path.join(path, 'ManagersInfo.dat'))

def update_table(w, attlist, just_dates, former_member_file, my_path):
    w['Start']=w.apply(lambda row: row['Start'].strftime('%b, %Y'),1)
    formert_members = list(former_member_file.index)
    for name in attlist.index:
        if name not in formert_members:
            fi = re.search(r'([\D.]*), ([\d]+)', attlist.loc[name, 'Starting month']).group(1)
            le = re.search(r'([\D.]*), ([\d]+)', attlist.loc[name, 'Starting month']).group(2)
            if len(le)==4:
                le = le[-2]+le[-1] 
            file_ = fi+le+'.dat'
            if not os.path.exists(os.path.join(my_path, file_)):
                new_file = just_dates.loc[name].dropna().to_frame().T
                for col in new_file.columns:
                    if ( datetime.fromtimestamp(mktime(time.strptime(col, "%d/%m/%Y"))) - datetime.fromtimestamp(mktime(time.strptime(fi+le, "%b%y"))) ).days <0:
                        new_file = new_file.drop(col,1)
                new_file.to_pickle(os.path.join(my_path, file_))
            else:
                df = pd.read_pickle(os.path.join(my_path, file_))
                if name not in df.index:
                    '''
                    If a new file was created for the previous member (as a new group was opened
                    at the beginning of the month), the next member will not shown in the index of
                    the new file so we need to add this member.
                    '''
                    for col in just_dates.columns:
                        if ( datetime.fromtimestamp(mktime(time.strptime(col, "%d/%m/%Y"))) - datetime.fromtimestamp(mktime(time.strptime(fi+le, "%b%y"))) ).days <0:
                            pass
                        else:
                            df.loc[name, col] = just_dates.loc[name, col]
                else:
                    for col in just_dates.columns:
                        if ( datetime.fromtimestamp(mktime(time.strptime(col, "%d/%m/%Y"))) - datetime.fromtimestamp(mktime(time.strptime(fi+le, "%b%y"))) ).days <0:
                            pass
                        else:
                            if col in df:
                                if type(df.loc[name, col])!=str:
                                    df.loc[name, col] = just_dates.loc[name, col]
                            elif col not in df:
                                df.loc[name, col] = just_dates.loc[name, col]
                df.to_pickle(os.path.join(my_path, file_))

def xlsx_attendance(xl_path, my_path, name):
    n_file = name+'_attendance.xlsx'
    if not os.path.exists(xl_path): 
        os.makedirs(xl_path)
    writer = pd.ExcelWriter(os.path.join(xl_path, n_file), engine='xlsxwriter')
    for file_ in os.listdir(my_path):
        df = pd.read_pickle(os.path.join(my_path, file_))
        sh = re.search(r'([\D]*)([\d]+)',file_).group(1)
        eet = re.search(r'([\D]*)([\d]+)',file_).group(2)
        sheet_n = sh+', '+eet
        df.to_excel(writer, sheet_name=sheet_n)
    return n_file

def excel_saver(df, filename):
    path = r'C:/Users/me'
    writer = pd.ExcelWriter(os.path.join(path, filename+'.xlsx'), engine = 'xlsxwriter')
    df.to_excel(writer, sheet_name = 'Sheet1')
    writer.save()
    
if __name__ == "__main__":
    path = r'C:/Users/me'
    print('debug')
    managers_editor(path)
#     create_lesson_state_file(path)
#     track_='web'
#     create_attfiles(path, track=track_, d_start = '2015-08-05',filename = 'Aug15.dat', sheetname_ = 'Aug, 2015')
    