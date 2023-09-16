# 修圖大師V1
# V1版僅提供某頻率之後的修改(例如測試範圍為5~500Hz,修改250~500Hz這一段)

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
x_start = float(input('Please enter the start frequency: '))
x_final = float(input('Please enter the final frequency: '))
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
    x_break = float(input('Please enter the break point for this signal: '))
    scale = float(input('Please enter scale (ex.enter 0.5 if you want half of the peak): '))
    x = df[(df['Frequency']>=x_break)&(df['Frequency']<=x_final)]
    y = scale * x[[column_name]]
    x_before_break = df[df['Frequency']<x_break]
    y_before_break = x_before_break[[column_name]]
    y_sub = y.iloc[0][column_name]-y_before_break.iloc[-1][column_name]
    if y_sub > 0:
        y_final = y + y_sub
    if y_sub <= 0:
        y_final = y - y_sub
    mask = (df['Frequency']>=x_break)&(df['Frequency']<=x_final)
    df.loc[mask, column_name] = y_final[column_name].values
    plt.plot(df['Frequency'], df[column_name])
    plt.xlim(x_start, x_final)

plt.legend(l)
plt.grid()
plt.show()
output_path = os.path.join(os.path.dirname(filepath), 'output.xlsx')
df.to_excel(output_path, index=False)
print('Check out the file "output.xlsx" in directory of file you just selected.')