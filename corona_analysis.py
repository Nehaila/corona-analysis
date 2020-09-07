
import os 
import xlsxwriter
import openpyxl
from ggplot import *
import pandas as pd 
from plotnine import * 
from adjustText import adjust_text


df=pd.read_excel('C:/Users/32466/Desktop/python2/Corona_Combined.xlsx')
#print(os.getcwd()) 	#checking what is my directory here, it was /python but I want to change it
os.chdir('C:/Users/32466/Desktop/python2')	#chdir is for change direction to python2, a new folder
#print(os.getcwd()) 	#checking my new directory here
directory=os.getcwd()
#All of the following 3 steps are important to create an excel file using python xlsxwriter
workbook = xlsxwriter.Workbook('C:/Users/32466/Desktop/python2/Corona_Combined.xlsx') 
worksheet = workbook.add_worksheet() 
workbook.close()

#first put all the files into a dataframe and then append to one final large
for filename in os.listdir(directory):
	if filename.endswith('2020.csv'):
		df_or= pd.read_csv(filename)
		for col in df_or.columns:
			if col == 'Province_State':
				df_or= df_or[['Province_State', 'Country_Region' , 'Last_Update', 'Confirmed' , 'Deaths', 'Recovered' , 'FIPS', 'Admin2', 'Lat' , 'Long_', 'Active', 'Combined_Key']]
		
#The goal here is to have one homogenous file, the column order was different for the files 
		if len(df_or.columns) > 6:
			df_or= df_or.drop(columns= list(df_or.columns[6:len(df_or.columns)]), axis=1)
		df_or= df_or.rename(columns = {df_or.columns[0]:'Province' ,df_or.columns[1]:'Country', df_or.columns[2]:'Last Update', df_or.columns[3]:'Confirmed', df_or.columns[4]:'Deaths', df_or.columns[5]:'Recovered' }) 	
		df = df.append(df_or, ignore_index= True)
		if len(df.columns) > 6:
			df= df.drop(columns= list(df.columns[6:len(df.columns)]), axis=1)

df.loc[df['Deaths'].isnull() , 'Deaths'] == 0  

#then create a file for it, now we can analyse the data in one file
df.to_csv('combined.csv')
#print(df.info())
#print(df)
#for row in df.iloc[row]:
df=df.dropna(axis=0, subset=['Confirmed'])
pd.set_option('display.max_columns',6)
pd.set_option('display.max_rows',200)
#print(df)
# print(df.columns)
Nulls_in_rows = df.isnull().sum(axis=1)
Nulls_in_columns = df.isnull().sum(axis=0)
df=df.sort_values(by=['Last Update'],axis=0,ascending=True)


#print(df)
# #axis 1 for columns !
# print(Nulls_in_rows)
#print(Nulls_in_columns)  #we can see a lot of missing data in Province, we'll just stick to country
# print(df)
# print(df['Confirmed'])
# print(type(df['Confirmed']))
#I see that we have very few nulls that we can take care of easily without affecting data
#print(df.columns)
#print(df_or.columns)
#both have the same columns!! 
#print (df)
df['Last Update'] = pd.to_datetime(df['Last Update'])


df= df.groupby(by=[pd.Grouper('Country'),pd.Grouper(key='Last Update', freq='D')]).sum().reset_index()
df.to_excel('combined_excel.xlsx')


sum_deathss= df.groupby('Country')['Deaths'].max().reset_index()
sum_deathss=sum_deathss.sort_values(by=['Deaths'], ascending=False)
sum_deaths= sum_deathss[0:7]
sum_deaths['Last Update']= pd.to_datetime('2020-9-01')
print(sum_deaths)
sum_deaths.to_excel('Nehaila.xlsx')

#df=df.groupby(['Country']).sum()
#See the evolution of confirmed cases:
#ggplot(df, aes(x='Last Update' , y= 'Deaths'))


p = ggplot(aes(x='Last Update' , y='Deaths', group='Country', color='Country',label='Country'), data= df) +\
geom_line() + geom_text(aes(x='Last Update', y='Deaths',label='Country', size=5),data=sum_deaths)

print(p)