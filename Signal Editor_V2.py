# Signal Editor V2
# V2版提供只修改特定範圍的功能
# 例如測試範圍為5~500Hz,但只需要修改200Hz~250Hz這一段
# 

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

filepath = get_file_path()
df = pd.read_excel(filepath)


signal_number = int(input('How many signal do you want to process? '))
x_start = float(input('Please enter the start frequency (ex.5): '))
x_final = float(input('Please enter the final frequency (ex.500): '))
l = list()
for i in range(1, signal_number+1):
    if i == 1:
        column_name = input('Please enter the 1st signal name: ')
    if i == 2:
        column_name = input('Please enter the 2nd signal name: ')
    if i == 3:
        column_name = input('Please enter the 3rd signal name: ')
    if i > 3:
        column_name = input(f'Please enter the {i}th signal name: ')
    l.append(column_name)
    x_1st_cut = float(input('Please enter the first cut point for this signal: '))
    x_2nd_cut = float(input('Please enter the second cut point for this signal: '))
    scale = float(input('Please enter scale (ex.enter 0.5 if you want half of the peak): '))
    if x_2nd_cut == x_final:
        x = df[(df['Frequency']>=x_1st_cut)&(df['Frequency']<=x_final)]
        y = scale * x[[column_name]]
        x_before_1st_cut = df[df['Frequency']<x_1st_cut]
        y_before_1st_cut = x_before_1st_cut[[column_name]]
        y_sub = y.iloc[0][column_name]-y_before_1st_cut.iloc[-1][column_name]
        if y_sub > 0:
            y_final = y + y_sub
        if y_sub <= 0:
            y_final = y - y_sub
        mask = (df['Frequency']>=x_1st_cut)&(df['Frequency']<=x_final)
        df.loc[mask, column_name] = y_final[column_name].values
        plt.plot(df['Frequency'], df[column_name])
        plt.xlim(x_start, x_final)
    else:
        x_in_range = df[(df['Frequency']>=x_1st_cut)&(df['Frequency']<=x_2nd_cut)]
        y_in_range = scale * x_in_range[[column_name]]
        x_before_1st_cut = df[df['Frequency']<x_1st_cut]
        y_before_1st_cut = x_before_1st_cut[[column_name]]
        y_sub = y_in_range.iloc[0][column_name] - y_before_1st_cut.iloc[-1][column_name]
        if y_sub > 0:
            y_in_range = y_in_range + y_sub
        if y_sub < 0:
            y_in_range = y_in_range - y_sub
        mask = (df['Frequency']>=x_1st_cut)&(df['Frequency']<=x_2nd_cut)
        df.loc[mask, column_name] = y_in_range[column_name].values
        
        x_after_range = df[(df['Frequency']>x_2nd_cut)&(df['Frequency']<=x_final)]
        y_after_range = x_after_range[[column_name]]
        y_sub2 = y_after_range.iloc[0][column_name] - y_in_range.iloc[-1][column_name]
        if y_sub2 > 0:
            y_after_range = y_after_range - y_sub2
        if y_sub2 < 0:
            y_after_range = y_after_range + y_sub2
        mask2 = (df['Frequency']>x_2nd_cut)&(df['Frequency']<=x_final)
        df.loc[mask2, column_name] = y_after_range[column_name].values
        plt.plot(df['Frequency'], df[column_name])
        plt.xlim(x_start, x_final)
        
plt.legend(l)
plt.grid()
plt.show()
output_path = os.path.join(os.path.dirname(filepath), 'output.xlsx')
df.to_excel(output_path, index=False)
print('Check out the file "output.xlsx" in directory of file you just selected.')