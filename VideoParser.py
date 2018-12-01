#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  VideoParser.py
#
#  Copyright 2018 赵国涛 <guotao.zhao@vivo.com>
#
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#
#



import os
import sys
import xlwt
from moviepy.editor import VideoFileClip

file_dir = "F:/video/android&arduino/" #定义文件目录

class VideoCheck():

    def __init__(self):
        self.file_dir = file_dir

    def get_filesize(self,filename):
        """
        Get file size（M: MB）
        """
        file_byte = os.path.getsize(filename)
        return self.sizeConvert(file_byte)

    def get_file_times(self,filename):
        """
        Get video time（s）
        """
        clip = VideoFileClip(filename)
        file_time = self.timeConvert(clip.duration)
        return file_time

    def sizeConvert(self,size):# 单位换算
        K, M, G = 1024, 1024**2, 1024**3
        if size >= G:
            return str(round(size/G, 2))+'GB'
        elif size >= M:
            return str(round(size/M,2))+'MB'
        elif size >= K:
            return str(round(size/K,2))+'KB'
        else:
            return str(size)+'B'

    def timeConvert(self,size):# 单位换算
        M, H = 60, 60**2
        if size < M:
            return str(size)+'s'
        if size < H:
            return '%sm%ss'%(int(size/M),int(size%M))
        else:
            hour = int(size/H)
            mine = int(size%H/M)
            second = int(size%H%M)
            tim_srt = '%sh%sm%ss'%(hour,mine,second)
            return tim_srt

    def get_all_file(self):
        """
        Get all video files
        """
        for root, dirs, files in os.walk(file_dir):
            return files


print(">>>>>>>>>>>>>>>>>start...")
fc = VideoCheck()
files = fc.get_all_file()
datas = [['文件名称', '文件大小', '视频时长']]#二维数组
for f in files:
    cell = []
    file_path = os.path.join(file_dir,f)
    #~ print("file path:",file_path)
    file_size = fc.get_filesize(file_path)
    #~ print("file_size:",file_size)
    file_times = fc.get_file_times(file_path)
    #~ print("file times:",file_times)
    print("文件名字：{filename},大小：{filesize},时长：{filetimes}".format(filename=f,filesize=file_size,filetimes=file_times))
#~ '''
    cell.append(f)
    cell.append(file_size)
    cell.append(file_times)
    datas.append(cell)

wb = xlwt.Workbook() #Create xlsx file
sheet = wb.add_sheet('test')#sheet的名称为test

#单元格的格式
style = 'pattern: pattern solid, fore_colour yellow; '#背景颜色为黄色
style += 'font: bold on; '#粗体字
style += 'align: horz centre, vert center; '#居中
header_style = xlwt.easyxf(style)

row_count = len(datas)
col_count = len(datas[0])
for row in range(0, row_count):
    col_count = len(datas[row])
    for col in range(0, col_count):
        if row == 0:#设置表头单元格的格式
            sheet.write(row, col, datas[row][col], header_style)
        else:
            sheet.write(row, col, datas[row][col])
wb.save(file_dir+"videoParserResults.xls")
print("============Finished================")
#~ '''