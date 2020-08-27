from openpyxl import Workbook
import openpyxl
import pandas as pd
import numpy as np
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
df=pd.read_csv("example_2.csv")
df=df.values.tolist()
'''
wb = openpyxl.load_workbook('Assignment.xlsx')

ws = wb.active
ws['E7']='Ataur'

wb.save("Assignment.xlsx")
'''
                                     #Percentile formula= no. of students who have lesser score *100/total students
def percentile(total_score,sno):
	list1_percentile=[]
	for i in range(0,len(total_score)):
		count=0
		for j in range(0,len(total_score)):
			if(total_score[i]>total_score[j]):
				count+=1
				j+=1
			else:
				j+=1	
		list1_percentile+=[(count+1)*100/len(total_score)]		
			
	#print(list1)
	return list1_percentile[sno]


def second():
	data=[]
	for i in range(0,len(df),5):
		data+=df[i:i+5]  # data is 2D list with entire raw data where 'i' is every candidate with its 5 values
	time=[]
	score_c=[]
	score_ic=[]
	attempt=[]
	marked=[]
	correct=[]
	outcome=[]
	your_score=[]
	for i in range(0,len(df)):            # All the list values below are 1D lists
		time+=[data[i][12]]               # list of values of cloumn 'Time spent on question'
		score_c+=[data[i][13]]            # list of values of cloumn 'Score if correct' 
		score_ic+=[data[i][14]]           # list of values of cloumn 'Score if incorrect'
		attempt+=[data[i][15]]            # list of values of cloumn 'Attempt status '
		marked+=[data[i][16]]			  # list of values of cloumn 'What you marked'	
		correct+=[data[i][17]]            # list of values of cloumn 'Correct Answer'
		outcome+=[data[i][18]]            # list of values of cloumn 'Outcome (Correct/Incorrect/Not Attempted)'
		your_score+=[data[i][19]]	      # list of values of cloumn 'Your score'
	#totalScore(your_score)
	n=5	   # Can be defined as (len(df[j])/len(data))
	cscore= [score_c[i * n:(i + 1) * n] for i in range((len(score_c) + n - 1) // n )]
	n=5	
	icscore= [score_ic[i * n:(i + 1) * n] for i in range((len(score_ic) + n - 1) // n )]         
	n=5	                                                                                         #These functions convert a 1D list above
	cattempt= [attempt[i * n:(i + 1) * n] for i in range((len(attempt) + n - 1) // n )]          #into 2D list where 'i' is first Student
	n=5	                                                                                         #details and so on
	cmarked= [marked[i * n:(i + 1) * n] for i in range((len(marked) + n - 1) // n )]             
	n=5	
	ctime= [time[i * n:(i + 1) * n] for i in range((len(time) + n - 1) // n )]
	n=5	
	ccorrect= [correct[i * n:(i + 1) * n] for i in range((len(correct) + n - 1) // n )]
	n=5	
	coutcome= [outcome[i * n:(i + 1) * n] for i in range((len(outcome) + n - 1) // n )]
	n=5	
	cyour_score= [your_score[i * n:(i + 1) * n] for i in range((len(your_score) + n - 1) // n )]
	#for i in range(0,len(ctime)):
		#ctime[i]
	#return ctime[sno],cscore[sno],your_score,icscore[sno],cattempt[sno],cmarked[sno],ccorrect[sno],coutcome[sno],cyour_score[sno]
	for i in range(0,len(ctime)):
		bar(ctime[i])


def bar(time):
	question=['Q1','Q2','Q3','Q4','Q5']
	y_pos = np.arange(len(question))
	plt.bar(y_pos,time)
	plt.xticks(y_pos, question)
	plt.title('Time(Sec)')
	my_path = os.path.abspath('C:\Projects\ReportCard\bar_2\\')
	my_file= +str(i)+'.png'
	plt.savefig(os.path.join(my_path,my_file))

def pie1(time):
	labels = 'Q1', 'Q2', 'Q3', 'Q4','Q5'
	colors = ['gold', 'red', 'lightcoral', 'lightskyblue','green']
	wedges,patches, texts = plt.pie(time,autopct='%1.1f%%', colors=colors,labels = labels,shadow=True)
	plt.legend(patches, wedges, loc="best")
	plt.axis('equal')
	plt.title('Time Spent as a function of total time')
	my_path = os.path.abspath('C:\Projects\ReportCard\pie1\\') 
    my_file = +str(i)+'.png'
	plt.savefig(os.path.join(my_path,my_file))

def pie2(attempt):
	labels=['Attempted','Not Attempted']
	ro=[attempt.count('Attempted'),attempt.count('Unattempted')]
	colors=['gold','red']
	plt.pie(ro,labels=labels, colors=colors,autopct='%1.1f%%')
	plt.axis('equal')
	plt.title('Attempts')
	my_path = os.path.abspath('C:\Projects\ReportCard\pie2_2\\') 
    my_file = +str(i)+'.png'
	plt.savefig(os.path.join(my_path,my_file))

def pie4(outcome):
	labels=['Correct','Incorrect','Not Attempted']
	ro=[outcome.count('Correct'),outcome.count('Incorrect'),outcome.count('Unattempted')]
	colors=['gold','red','green']
	plt.pie(ro,labels=labels, colors=colors,autopct='%1.1f%%')
	plt.axis('equal')
	plt.title('Overall Performance against the Test')
	my_path = os.path.abspath('C:\Projects\ReportCard\pie4_2\\') 
    my_file = +str(i)+'.png'
	plt.savefig(os.path.join(my_path,my_file))

def pie3(outcome):
	labels=['Correct','Incorrect']
	ro=[outcome.count('Correct'),outcome.count('Incorrect')]
	colors=['gold','red']
	plt.pie(ro,labels=labels, colors=colors,autopct='%1.1f%%')
	plt.axis('equal')
	plt.title('Accuracy from Attempted Questions')
	my_path = os.path.abspath('C:\Projects\ReportCard\pie3\\') 
    my_file = +str(i)+'.png'
	plt.savefig(os.path.join(my_path,my_file))