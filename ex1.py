from openpyxl.drawing.image import Image
from openpyxl import Workbook
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from ex import *
from function import *
def final_files():
	name_files =['outfile0.xlsx','outfile1.xlsx','outfile2.xlsx','outfile3.xlsx','outfile4.xlsx','outfile5.xlsx','outfile6.xlsx','outfile7.xlsx','outfile8.xlsx','outfile9.xlsx','outfile10.xlsx','outfile11.xlsx','outfile12.xlsx','outfile13.xlsx','outfile14.xlsx','outfile15.xlsx','outfile16.xlsx','outfile17.xlsx','outfile18.xlsx','outfile19.xlsx']

	for i in range(0,len(name_files)):
		wb = openpyxl.load_workbook(name_files[i])
		ws = wb.active
		gpicture=getpicture(i)
		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\pics_2'+'\\'+gpicture)
		img.height=70
		img.width=100
		img.anchor = 'J3'
		ws.add_image(img)

		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\bar_2\\'+str(i)+'.png')
		img.height=150
		img.width=170
		img.anchor='D27'
		ws.add_image(img)
		
		
		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie1_2\\'+str(i)+'.png')
		img.height=150
		img.width=170
		img.anchor = 'I27'
		ws.add_image(img)
		
		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie2_2\\'+str(i)+'.png')
		img.height=150
		img.width=170
		img.anchor = 'D37'
		ws.add_image(img)
		
		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie3_2\\'+str(i)+'.png')
		img.height=150
		img.width=172.20000000000002
		img.anchor = 'I37'
		ws.add_image(img)
		
		img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie4_2\\'+str(i)+'.png')
		img.height=150
		img.width=170.79999999999998
		img.anchor = 'D48'
		ws.add_image(img)
		
		wb.save("new_outfile"+str(i)+'.xlsx')
		os.remove("outfile"+str(i)+'.xlsx')     #cleanup code to remove excess files

