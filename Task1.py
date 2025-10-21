#importing pandas for dataframe and openpyxl to read excel
import pandas as pd
import openpyxl
import numpy as np
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
deparments=employeesNoN['Department'].unique()
avgAgeByDept=employeesNoN.groupby('Department')['Age'].mean()
avgSalaryByDept=employeesNoN.groupby('Department')['Annual Salary'].mean()

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

#Plotting average age and salary by department
x=range(len(deparments))
plt.subplot(121)
plt.bar(x,avgAgeByDept,width=0.4,label='Average Age',align='center')
plt.xlabel('Departments')
plt.xticks([i for i in x],deparments,rotation=90)
plt.ylabel('Average Age')
plt.title('Average Age by Department')
plt.ylim(42,48)
plt.legend()
plt.tight_layout()
plt.subplot(122)
plt.bar(x,avgSalaryByDept,width=0.4,label='Average Salary',align='center',color='orange')
plt.xlabel('Departments')
plt.xticks([i for i in x],deparments,rotation=90)
plt.ylabel('Average Salary')
plt.title('Average Salary by Department')
plt.ylim(90000,130000)
plt.legend()
plt.tight_layout()
plt.show()

#Plotting average age and salary by busness unit
businessUnits=employeesNoN['Business Unit'].unique()
avgAgeByDept=employeesNoN.groupby('Business Unit')['Age'].mean()
avgSalaryByDept=employeesNoN.groupby('Business Unit')['Annual Salary'].mean()
x=range(len(businessUnits))
plt.subplot(121)
plt.bar(x,avgAgeByDept,width=0.4,label='Average Age',align='center')
plt.xlabel('Business Units')
plt.xticks([i for i in x],businessUnits,rotation=90)
plt.ylabel('Average Age')
plt.title('Average Age by Business Unit')
plt.ylim(42,48)
plt.legend()
plt.tight_layout()
plt.subplot(122)
plt.bar(x,avgSalaryByDept,width=0.4,label='Average Salary',align='center',color='orange')
plt.xlabel('Business Units')
plt.xticks([i for i in x],businessUnits,rotation=90)
plt.ylabel('Average Salary')
plt.title('Average Salary by Business Unit')
plt.ylim(90000,130000)
plt.legend()
plt.tight_layout()
plt.show()

#plotting average salary by age group for each department
minAge=employeesNoN['Age'].min()
maxAge=employeesNoN['Age'].max()
ageGroups=range(int(minAge),int(maxAge+1))
count=1
for dept in deparments:
    plt.subplot(3,3,count,label=dept, xlabel='Age',ylabel='Average Annual Salary')
    plt.title(dept)
    plt.subplots_adjust(wspace=0.4,hspace=0.6)
    count+=1
    deptData={}
    for x in ageGroups:
        deptData[x]=employeesNoN[(employeesNoN['Department']==dept) & (employeesNoN['Age']==x)]['Annual Salary'].mean()
    plt.scatter(ageGroups,deptData.values())
plt.legend()
plt.show()

#plotting maximum and minimum age by ethnicity per department
count=1
ethinicities=employeesNoN['Ethnicity'].unique()
ethAgeMin=employeesNoN.groupby('Ethnicity')['Age'].min()
ethAgeMax=employeesNoN.groupby('Ethnicity')['Age'].max()
for dept in deparments:
    plt.subplot(3,3,count,label=dept,xlabel='Ethinicity',ylabel='Age')
    plt.title(dept)
    plt.subplots_adjust(wspace=0.4,hspace=0.6)
    count+=1
    for eth in ethinicities:
        ethDeptData=employeesNoN[(employeesNoN['Department']==dept) & (employeesNoN['Ethnicity']==eth)]
        plt.scatter(eth,ethDeptData['Age'].min(),label=eth)
        plt.scatter(eth,ethDeptData['Age'].max(),label=eth)
    plt.ylim(20,70)
plt.show()

#plotting median salaries by ethnicity in each department
for eth in ethinicities:
    ethDeptData=employeesNoN[(employeesNoN['Ethnicity']==eth)].groupby('Department')['Annual Salary'].median()
    plt.scatter(deparments,ethDeptData,label=eth)
plt.legend()
plt.xlabel('Department')
plt.ylabel('Median Annual Salary')
plt.title('Median Annual Salary by Ethinicity in Each Department')
plt.show()


#plotting employee count info by ethnicity
empEth={}
for eth in ethinicities:
    empEth[eth]=employeesNoN[employeesNoN['Ethnicity']==eth]['EEID'].count()
empEth=pd.Series(empEth)
plt.subplot(131)
plt.bar(ethinicities,empEth)
plt.xlabel('Ethinicity')
plt.ylabel('Employee Count')
plt.title('Employee Count by Ethinicity')

#plotting employee count info by country
empCnt={}
for country in employeesNoN['Country'].unique():
    empCnt[country]=employeesNoN[employeesNoN['Country']==country]['EEID'].count()
empCnt=pd.Series(empCnt)
plt.subplot(132)
countries=employeesNoN['Country'].unique()
plt.bar(countries,empCnt)
plt.xlabel('Country')
plt.ylabel('Employee Count')
plt.title('Employee Count by Country')

#plotting employee  count info by ethincity per country
blackEmpByCountry={}
asianEmpByCountry={}
latinoEmpByCountry={}
caucasianEmpByCountry={}
for country in countries:
    blackEmpByCountry[country]=employeesNoN[(employeesNoN['Ethnicity']=='Black') & (employeesNoN['Country']==country)]['EEID'].count()
    asianEmpByCountry[country]=employeesNoN[(employeesNoN['Ethnicity']=='Asian') & (employeesNoN['Country']==country)]['EEID'].count()
    latinoEmpByCountry[country]=employeesNoN[(employeesNoN['Ethnicity']=='Latino') & (employeesNoN['Country']==country)]['EEID'].count()
    caucasianEmpByCountry[country]=employeesNoN[(employeesNoN['Ethnicity']=='Caucasian') & (employeesNoN['Country']==country)]['EEID'].count()
asianEmpByCountry=pd.Series(asianEmpByCountry)
latinoEmpByCountry=pd.Series(latinoEmpByCountry)
caucasianEmpByCountry=pd.Series(caucasianEmpByCountry)
blackEmpByCountry=pd.Series(blackEmpByCountry)

plt.subplot(133)
x=range(len(countries))
plt.scatter(x,blackEmpByCountry.values,label='Black')
plt.scatter(x,asianEmpByCountry.values,label='Asian')
plt.scatter(x,latinoEmpByCountry.values,label='Latino')
plt.scatter(x,caucasianEmpByCountry.values,label='Caucasian')
plt.xlabel('Country')
plt.xticks([i for i in x],countries)
plt.ylabel('Employee Count')
plt.title('Employee Count by Ethinicity per Country')
plt.legend()
plt.tight_layout()
plt.show()

#plotting employee count and salary by gender
plt.subplot(121)
plt.bar(employeesNoN['Gender'].unique(),employeesNoN.groupby('Gender')['EEID'].count())
plt.ylim(400,550)
plt.title('Employee Count by Gender')
plt.subplot(122)
plt.bar(employeesNoN['Gender'].unique(),employeesNoN.groupby('Gender')['Annual Salary'].mean())
plt.ylim(110000,115000)
plt.title('Average Annual Salary by Gender')
plt.show()  

#Annual salary average by gender per department
count=1
for dept in deparments:
    plt.subplot(3,3,count,label=dept, xlabel='Gender',ylabel='Average Annual Salary')
    plt.title(dept)
    plt.subplots_adjust(wspace=0.4,hspace=0.6)
    count+=1
    deptData={}
    for x in employeesNoN['Gender'].unique():
        deptData[x]=employeesNoN[(employeesNoN['Department']==dept) & (employeesNoN['Gender']==x)]['Annual Salary'].mean()
    plt.bar(employeesNoN['Gender'].unique(),deptData.values())
    print(deptData.values())
    plt.ylim(min(deptData.values())-5000,max(deptData.values())+5000)
plt.legend()
plt.show()

#Annual salary by gender by deparment by age for accounting marketing human resources and engineering
deparmentsSubset=['Accounting','Marketing','Human Resources','Engineering']
count=1
for dept in deparmentsSubset:
    plt.subplot(2,2,count,label=dept, xlabel='Age',ylabel='Annual Salary')
    plt.title(dept)
    plt.subplots_adjust(wspace=0.4,hspace=0.6)
    count+=1
    for gender in employeesNoN['Gender'].unique():
        deptData={}
        for x in ageGroups:
            deptData[x]=employeesNoN[(employeesNoN['Department']==dept) & (employeesNoN['Age']==x) & (employeesNoN['Gender']==gender)]['Annual Salary'].mean()
        plt.scatter(ageGroups,deptData.values(),label=gender)
    plt.ylim(min(deptData.values())-5000,max(deptData.values())+5000)
plt.legend()
plt.show()
