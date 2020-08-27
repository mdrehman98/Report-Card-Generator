from win32com import client
import win32api
from openpyxl import Workbook
import openpyxl
import pandas as pd
import numpy as np
from openpyxl.drawing.image import Image
from report import *
import os
import sys
import pathlib
import glob  
import cv2
from ex import *
from ex1 import *
from function import *
'''
name_files =['new_outfile0.xlsx','new_outfile1.xlsx','new_outfile2.xlsx','new_outfile3.xlsx','new_outfile4.xlsx','new_outfile5.xlsx','new_outfile6.xlsx','new_outfile7.xlsx','new_outfile8.xlsx','new_outfile9.xlsx']
for i in range(0,len(name_files)):
	input_file = r'C:\Projects\ReportCard'+'\new_outfile'+str(i)+'.xlsx'
	#give your file name with valid path 
	output_file = r'C:\Projects\ReportCard\dummy_pdfs'+'\dummy'+str(i)+'.pdf'
	#give valid output file name and path
	app = client.DispatchEx("Excel.Application")
	app.Interactive = False
	app.Visible = False
	Workbook = app.Workbooks.Open(input_file)
	try:
		Workbook.ActiveSheet.ExportAsFixedFormat(0,output_file)
	finally:
    	Workbook.Close()
    	app.Exit()	
''' 
name_files =['new_outfile0.xlsx','new_outfile1.xlsx','new_outfile2.xlsx','new_outfile3.xlsx','new_outfile4.xlsx','new_outfile5.xlsx','new_outfile6.xlsx','new_outfile7.xlsx','new_outfile8.xlsx','new_outfile9.xlsx','new_outfile10.xlsx','new_outfile11.xlsx','new_outfile12.xlsx','new_outfile13.xlsx','new_outfile14.xlsx','new_outfile15.xlsx','new_outfile16.xlsx','new_outfile17.xlsx','new_outfile18.xlsx','new_outfile19.xlsx']
for i in range(0,len(name_files)):   	
	input_file = r'C:\Projects\ReportCard'+'\\new_outfile'+str(i)+'.xlsx'
	#give your file name with valid path 
	output_file = r'C:\Projects\ReportCard\dummy_pdfs'+'\\dummy'+str(i)+'.pdf'
	#give valid output file name and path
	app = client.DispatchEx("Excel.Application")
	app.Interactive = False
	app.Visible = False
	Workbook = app.Workbooks.Open(input_file)
	try:
	    Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
	except Exception as e:
	    print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
	    print(str(e))
	finally:
	    Workbook.Close()
	        		