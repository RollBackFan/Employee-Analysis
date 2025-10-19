#importing pandas for dataframe and openpyxl to read excel
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

#Read excel sheet from downloads
employees=pd.read_excel('C:/Users/dawod/Downloads/Employee Sample Data - A.xlsx')

#assuming empty bonus cells mean no bonus fill all of them with 0
employees.fillna({'Bonus %':0},inplace=True)

#create a dataframe of all employees with lacking information except exit date since it does not need to be filled
employeesNoN=employees.dropna(subset=['EEID','Full Name','Job Title','Department','Business Unit','Gender','Ethnicity','Age','Hire Date','Annual Salary','Bonus %','Country','City']) 
lacksInfo=pd.concat([employees,employeesNoN])
lacksInfo.drop_duplicates(keep=False,inplace=True)

#saving a spreadsheet of employees with missing info
lacksInfo.to_excel('C:/Users/dawod/Downloads/Lacksinfo.xlsx')

#Changing name of first employee
employeesNoN.loc[0,'Full Name']="Emily David"   
#Raise bonus of second employee and decrease third employee's bonus
employeesNoN.loc[1,'Bonus %']=0.1
employeesNoN.loc[2,'Bonus %']=0.1
#increase pay of employee 4 and decrease pay of empoyee 5
employeesNoN.loc[3,'Annual Salary']=85000
employeesNoN.loc[4,'Annual Salary']=95000

#printing row of largest salary
print(employeesNoN.loc[employeesNoN['Annual Salary'].idxmax()])

#Finding the average age and salary by department
print(employeesNoN.groupby('Department').agg({'Age':'mean','Annual Salary':'mean'}))

#max/min salary and median salary by ethnicity and department
print(employeesNoN.groupby(['Department','Ethnicity']).agg({'Age':['max','min'],'Annual Salary':'median'}))

#saving a spreadsheet of the final version of the excel sheet
employeesNoN.to_excel('C:/Users/dawod/Downloads/task.xlsx')

#checking employment and pay difference between men and women
print("\n",employeesNoN.groupby(['Gender']).agg({'EEID':'count','Annual Salary':'mean'}))

#checking employment and pay difference between men and women
print("\n",employeesNoN.groupby(['Department','Gender']).agg({'EEID':'count','Annual Salary':'mean'}))

#Checking the number of employees of each ethnicity
print(employeesNoN.groupby('Ethnicity')['EEID'].count())
#results showed black people as the least employed in the company, to check if the situation is not more sepcific we check where black people hired at the company mostly live and where most employees in the company in general live
print("\n",employeesNoN.groupby(['Ethnicity','Country'] )['EEID'].count())
#This also explains the majority Asian ethnicity as all employees in China are asian and the majority of Asian employees are employed in China
print("\n",employeesNoN.groupby(['Country'])['EEID'].count())
employeesNoN.groupby(['Country'])['EEID'].count().plot(x='Country',y='count',kind='hist',title='plswork ',xlim=[0,10],ylim=[0,500])
plt.show()
#Found out that all black employees live in the US and that most of the company's employees live in the US which means there is an issue of some kind, check the details of ages by ethnicity of employees
print("\n",employeesNoN.groupby('Ethnicity').agg({'Age':['mean','min','max']}))
#Found out that black employees have the same min/max age as other employees but have the lowest mean age, the issue could be by department so we look into the details by department
print("\n",employeesNoN.groupby(['Department','Ethnicity']).agg({'EEID':'count','Age':['mean','min','max']}))
#While black employees are less in count than other ethnicities accounting, engineering, finance, and marketing are especially lacking, in accounting, engineering and HR black employees have a significantly smaller max age, in IT and accounting the youngest age for black employees was higher than other employees by a margin which could be alleviated with hiring programs focused on black youths

#Checking average annual salary by country
print("\n",employeesNoN.groupby(['Country']).agg({'Annual Salary':['mean','min','max']}))
#Annual Salaries are close in all aspects in all 3 countries where employees are hired even though avergae living cost in the USA is 5 times the average of China and 10 times that of Brazil, to check if that's because of the employment position we check the average abbual salary by department then check the average annual salary by department by country
print("\n",employeesNoN.groupby('Department')['Annual Salary'].mean())
print("\n",employeesNoN.groupby(['Country','Department']).agg({'EEID':'count','Annual Salary':'mean'}))
#Accountants and Sales have annual salaries higher than Brazil and China but every other department is paid less than at least 1 of the other countries