# V3
# you can decide how many cut point (1 or 2) for each signal in this version.

import pandas as pd
import matplotlib.pyplot as plt
import os
from tkinter import Tk, filedialog

def get_file_path():
    print('Please select an Excel file: ')
    print('(Please make sure the first column of the first row of your Excel file is "Frequency" first!)')
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    root.destroy()
    return file_path

def signal_name(i):
    if i == 0:
        column_name = input('Please enter the 1st signal name: ')
    if i == 1:
        column_name = input('Please enter the 2nd signal name: ')
    if i == 2:
        column_name = input('Please enter the 3rd signal name: ')
    if i > 2:
        column_name = input(f'Please enter the {i}th signal name: ')
    return column_name
    
def Multiple_in(df, x1, mask, column_name, scale):
    x = df[mask]
    y = scale * x[[column_name]]
    x_before = df[df['Frequency']<x1]
    y_before = x_before[[column_name]]
    y_sub = y.iloc[0][column_name]-y_before.iloc[-1][column_name]
    df.loc[mask, column_name] = y[column_name].values
    return df[column_name], y_sub, y # this y is from mask1

def Multiple_after(df, mask1, mask2, column_name):
    x_in = df[mask1]
    y_in = x_in[[column_name]]
    x = df[mask2]
    y = x[[column_name]]
    y_sub = y.iloc[0][column_name] - y_in.iloc[-1][column_name]
    df.loc[mask2, column_name] = y[column_name].values
    return df[column_name], y_sub, y # this y is from mask2

def Correction(y_sub, y):
    if y_sub > 0:
        y = y + y_sub
    else:
        y = y - y_sub
    return y
def Correction_inverse(y_sub, y):
    if y_sub > 0:
        y = y - y_sub
    else:
        y = y + y_sub
    return y

filepath = get_file_path()
df = pd.read_excel(filepath)
signal_number = int(input('How many signal do you want to process? '))
end_decide = int(input('How many cut point for each signal? (1 or 2): '))
x_start = float(input('Please enter the start frequency (ex.5): '))
x_final = float(input('Please enter the final frequency (ex.500): '))
cn_list = list()
y_sub1_list = list()
y_sub2_list = list()
y1_list = list()
y2_list = list()
mask1_list = list()
mask2_list = list()
if end_decide == 1:
    for i in range(signal_number):
        column_name = signal_name(i)
        x1 = float(input('Please enter the cut point for this signal: '))
        x2 = x_final
        scale = float(input('Please enter scale (ex.enter 0.5 if you want half of the peak): '))
        mask1 = (df['Frequency']>=x1)&(df['Frequency']<=x2)
        df[column_name], y_sub1, y1 = Multiple_in(df, x1, mask1, column_name, scale)
        mask1_list.append(mask1)
        cn_list.append(column_name)
        y_sub1_list.append(y_sub1)
        y1_list.append(y1)
        plt.plot(df['Frequency'], df[column_name])
        plt.xlim(x_start, x_final)
if end_decide == 2:
    for i in range(signal_number):
        column_name = signal_name(i)
        x1 = float(input('Please enter the first cut point for this signal: '))
        x2 = float(input('Please enter the second cut point for this signal: '))
        scale = float(input('Please enter scale (ex.enter 0.5 if you want half of the peak): '))
        mask1 = (df['Frequency']>=x1)&(df['Frequency']<=x2)
        mask2 = (df['Frequency']>x2)&(df['Frequency']<=x_final)
        df[column_name], y_sub1, y1 = Multiple_in(df, x1, mask1, column_name, scale)
        df[column_name], y_sub2, y2 = Multiple_after(df, mask1, mask2, column_name)
        mask1_list.append(mask1)
        mask2_list.append(mask2)
        cn_list.append(column_name)
        y_sub1_list.append(y_sub1)
        y_sub2_list.append(y_sub2)
        y1_list.append(y1)
        y2_list.append(y2)
        plt.plot(df['Frequency'], df[column_name])
        plt.xlim(x_start, x_final)
plt.legend(cn_list)
plt.grid()
plt.show()


if end_decide == 1:
    for i in range(signal_number):
        print(f'signal {cn_list[i]} may need a correction of {y_sub1_list[i]}')
if end_decide == 2:
    for i in range(signal_number):
        print(f'signal {cn_list[i]} may need a correction of {y_sub1_list[i]} at first cut point')
        print(f'signal {cn_list[i]} may need a correction of {y_sub2_list[i]} at second cut point')
    
ask_correction = input('Do you want a correction for your signal(s)? (y/n)')
if ask_correction == 'n':
    output_path = os.path.join(os.path.dirname(filepath), 'output.xlsx')
    df.to_excel(output_path, index=False)
    print('Check out the file "output.xlsx" in directory of file you just selected.')
if ask_correction == 'y':
    if end_decide == 1:
        for i in range(signal_number):
            column_name = cn_list[i]
            y_sub1 = y_sub1_list[i]
            y1 = y1_list[i]
            df.loc[mask1_list[i], column_name] = Correction(y_sub1, y1)[column_name].values
            plt.plot(df['Frequency'], df[column_name])
            plt.xlim(x_start, x_final)
        plt.legend(cn_list)
        plt.grid()
        plt.show()
    if end_decide == 2:
        for i in range(signal_number):
            column_name = cn_list[i]
            y_sub1 = y_sub1_list[i]
            y_sub2 = y_sub2_list[i]
            y1 = y1_list[i]
            y2 = y2_list[i]
            df.loc[mask1_list[i], column_name] = Correction(y_sub1, y1)[column_name].values
            df.loc[mask2_list[i], column_name] = Correction_inverse(y_sub2, y2)[column_name].values
            plt.plot(df['Frequency'], df[column_name])
            plt.xlim(x_start, x_final)
        plt.legend(cn_list)
        plt.grid()
        plt.show()
    print('Correction complete!')
    output_path = os.path.join(os.path.dirname(filepath), 'output.xlsx')
    df.to_excel(output_path, index=False)
    print('Check out the file "output.xlsx" in directory of file you just selected.')
    
#    plt.plot(df['Frequency'], df[column_name])
#    plt.xlim(x_start, x_final)

