import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import sys
import pathlib
import glob  
import cv2
from openpyxl.drawing.image import Image
from openpyxl import Workbook
import openpyxl

'''
def bar()
question=['Q1','Q2','Q3','Q4','Q5']
time=[20,40,60,45,15]
y_pos = np.arange(len(question))
plt.bar(y_pos,time)
plt.xticks(y_pos, question)
plt.title('Time(Sec)')
plt.show()
'''
'''
labels = 'Q1', 'Q2', 'Q3', 'Q4','Q5'
sizes = [60, 82, 49, 51,56]
colors = ['gold', 'yellow', 'lightcoral', 'lightskyblue','green']
wedges,patches, texts = plt.pie(sizes,autopct='%1.1f%%', colors=colors,labels = labels,shadow=True)
plt.legend(patches, wedges, loc="best")
plt.axis('equal')
plt.show()
'''
'''
labels=['Attempted','Not Attempted']
question=['Attempted','Unattempted','Attempted','Unattempted']
ro=[question.count('Attempted'),question.count('Unattempted')]
colors=['gold','red']
plt.pie(ro,labels=labels, colors=colors,autopct='%1.1f%%')
plt.axis('equal')
plt.show()
'''
def getpicture(sno):                                              #Pictures of individual candidates
	file_list =os.listdir('C:\Projects\ReportCard\pics_2') 
	pick_index=file_list.index(str(sno+1)+'.jpg') 
	return file_list[pick_index]
'''
wb = openpyxl.load_workbook('Assignment.xlsx')
ws = wb.active
path=glob.glob("C:/Projects/ReportCard/Pics/*.jpg")
for file in path:
	img=cv2.imread(file)
	cv2.imshow("Image",img)
	c3=cv2.imwrite(str(file)+'.jpg',img)
	cv2.waitKey(0)
	cv2.destroyAllWindows()
	img1 = openpyxl.drawing.image.Image(c3)
	img1.anchor = 'I27'
	ws.add_image(img1)
	wb.save('outfile'+str(file)+'.xlsx')
'''
'''
Plotting of graphs and Saving the images in Separate files
def graphs():
	plt1=bar(ctime)
	img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\bar\\'+str(i)+'.png')
	img.heigh=50
	img.width=100
	img.anchor='D27'
	ws.add_image(img)
		
	plt2=pie1(ctime)
	img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie1\\'+str(i)+'.png')
	img.height=100
	img.width=100
	img.anchor = 'I27'
	ws.add_image(img)
	
	plt3=pie2(cattempt)
	img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie2\\'+str(i)+'.png')
	img.height=100
	img.width=100
	img.anchor = 'D37'
	ws.add_image(img)
	plt4=pie3(coutcome)
	img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie3\\'+str(i)+'.png')
	img.height=100
	img.width=100
	img.anchor = 'I37'
	ws.add_image(img)
	plt5=pie4(coutcome)
	img = openpyxl.drawing.image.Image('C:\Projects\ReportCard\\pie4\\'+str(i)+'.png')
	img.height=100
	img.width=100
	img.anchor = 'D48'
	ws.add_image(img)
'''
