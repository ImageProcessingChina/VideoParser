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
import time
import sys
import xlwt
from moviepy.editor import VideoFileClip
#~ from moviepy.video.tools.segmenting import findObjects

#~ file_dir = "F:/video/demo/" #定义文件目录
file_dir = "." #定义文件目录为当前py文件所在目录

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
        clip.close()
        return file_time

    def get_file_Resulotion(self,filename):
        """
        Get video width x height
        """
        clip = VideoFileClip(filename)
        file_Resolution=clip.size
        f_res=(str(file_Resolution[0])+"x"+str(file_Resolution[1]))
        clip.close()
        return f_res

    def get_file_width(self,filename):
        """
        Get video width
        """
        clip = VideoFileClip(filename)
        file_width =clip.w
        clip.close()
        return file_width

    def get_file_height(self,filename):
        """
        Get video height
        """
        clip = VideoFileClip(filename)
        file_height = clip.h
        clip.close()
        return file_height

    def get_file_FrameRate(self,filename):
        """
        Get video frame rate
        """
        clip = VideoFileClip(filename)
        file_fps = clip.fps
        clip.close()
        return file_fps

    def get_file_DataRate(self,filename):
        """
        Get video data rate
        """
        clip = VideoFileClip(filename)
        file_datarate = clip.aspect_ratio
        clip.close()
        return file_datarate

    def get_file_BitRate(self,filename):
        """
        Get video bit rate
        """
        #~ clip = VideoFileClip(filename)
        #~ clip.close()
        return 0

    def get_file_Date(self,filename):
        """
        Get video date
        """
        #~ clip = VideoFileClip(filename)
        #~ clip.close()
        return 0

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
        Get all files
        """
        for root, dirs, files in os.walk(file_dir):
            return files

#demo.mp4 => demo_400x120@30fps_5.95s_40.87KB.mp4
print(">>>>>>>>>>>>>>>>>start...")
fc = VideoCheck()
#~ print(dir(fc))
files = fc.get_all_file()
#~ print(files)
datas = [['文件名称', '文件大小', '视频时长',"分辨率",'帧宽',"帧高","帧率","数据速率","总比特率","修改日期","新文件名"]]#二维数组
for f in files:
    if f.endswith('.mp4') or f.endswith('.wmv'):#os.path.splitext(f)[1]
        cell = []
        #~ print("old name:",f)
        file_path     = os.path.join(file_dir,f)
        #~ print(file_path)
        file_size     = fc.get_filesize(file_path)
        file_times    = fc.get_file_times(file_path)
        file_Res      = fc.get_file_Resulotion(file_path)
        file_width    = fc.get_file_width(file_path)
        file_height   = fc.get_file_height(file_path)
        file_Frate    = fc.get_file_FrameRate(file_path)
        file_DataRate = "-"#fc.get_file_DataRate(file_path)#not
        file_BitRate  = "-"#fc.get_file_BitRate(file_path)#not
        filemt=time.localtime(os.stat(file_path).st_mtime)
        fmt=time.strftime("%Y.%m.%d_%H:%M:%S",filemt)
        file_date     = fmt
        '''
        #修改时间
        filemt=time.localtime(os.stat(file_path).st_mtime)
        fmt=time.strftime("%Y%m%d_%H%M%S",filemt)
        print(fmt)
        #创建时间
        filect=time.localtime(os.stat(file_path).st_ctime)
        fct=time.strftime("%Y%m%d_%H%M%S",filect)
        print(fct)
        '''
        #~ print("文件名字：{filename},大小：{filesize},时长：{filetimes}".format(filename=f,filesize=file_size,filetimes=file_times))
        new_name = os.path.splitext(f)[0].split("#")[0]+"#"+file_Res+"@"+str(int(file_Frate))+"fps_"+file_times+"_"+file_size+os.path.splitext(f)[1]
        tot = 1
        while os.path.exists(new_name):
            new_name = os.path.splitext(f)[0].split("#")[0]+"#"+file_Res+"@"+str(int(file_Frate))+"fps_"+file_times+"_"+file_size+"_("+str(tot)+")"+os.path.splitext(f)[1]
            tot += 1
        print('新文件名：', new_name)
        #~ os.rename(f, new_name)
        print("----------------------------")

        cell.append(f)
        cell.append(file_size)
        cell.append(file_times)
        cell.append(file_Res)
        cell.append(file_width)
        cell.append(file_height)
        cell.append(file_Frate)
        cell.append(file_DataRate)
        cell.append(file_BitRate)
        cell.append(file_date)
        cell.append(new_name)
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
        #~ if col_count > 10:
            #~ sheet.col(col).width = 899*col_count
        if row == 0:#设置表头单元格的格式
            sheet.write(row, col, datas[row][col], header_style)
        else:
            sheet.write(row, col, datas[row][col])
wb.save(file_dir+"videoParserResults.xls")
print("============Finished================")
