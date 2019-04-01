# -*- coding:utf-8 -*-

from pyautocad import Autocad, APoint
import pandas as pd
import numpy as np
import os

# open AutoCAD
file = [i for i in os.listdir() if i.endswith('.xlsx') or i.endswith('xls')][0]
acad = Autocad(create_if_not_exists=True)
acad.prompt("Hello, Autocad from Python\n")

# read excel
table = pd.read_excel(file).fillna(0).values
num_row = len(table)
num_column = len(table[0])
# print(table)

# pre setting
x0 = 0.
y0 = 0.
size = 500
width = 1
margin = size/2.5
ratio_alpha = 1.5/2.5
ratio_c = 3.5/2.5

# start point
start_point = APoint(x0, y0)


# word size function
def word_length(word):
    if word:
        word = str(word)
        num_c = sum(['\u4e00' < c < '\u9fff' for c in word])
        return (len(word) - num_c) * ratio_alpha * size + num_c * ratio_c * size  
    else:
        return 0


word_size = []
for line in table:
    temp = []
    for word in line:
        temp.append(word_length(word))
    word_size.append(temp)

word_size = np.array(word_size)
# print(word_size)

# table size
table_height = [size + margin * 2 for i in range(num_row)]
table_length = np.max(word_size, axis=0) + margin * 2

# table position
table_x_length = np.cumsum(table_length)
table_x = start_point.x + np.concatenate(([0], table_x_length))
table_y_height = np.cumsum(table_height)
table_y = start_point.y - np.concatenate(([0], table_y_height))

end_point = APoint(table_x[-1], table_y[-1])

# print(table_length, table_x)
# print(table_height, table_y)

# word position
mid = True
if mid:
    word_x = table_x[:-1] + (table_length - word_size) / 2
else:
    word_x = table_x[:-1] + margin

word_y = table_y[1:] + margin

# draw table
# color = 'red'
for i in table_x[1:-1]:
    acad.model.AddLine(APoint(i, start_point.y), APoint(i, end_point.y))
for i in table_y[1:-1]:
    acad.model.AddLine(APoint(start_point.x, i), APoint(end_point.x, i))

# draw frame
for i in table_x[[0, -1]]:
    acad.model.AddLine(APoint(i, start_point.y), APoint(i, end_point.y))
for i in table_y[[0, -1]]:
    acad.model.AddLine(APoint(start_point.x, i), APoint(end_point.x, i))

# print(word_x, word_y)
# tape word
for i in range(len(table)):
    for j in range(len(table[0])):
        word = table[i][j]
        if word:
            acad.model.AddText(word, APoint(word_x[i][j], word_y[i]), size)
