import pandas as pd
import numpy as np


#import all tabs from excel
Maths_df = pd.read_excel('Python_Assignment.xlsx',sheet_name='Maths')
Physics_df = pd.read_excel('Python_Assignment.xlsx',sheet_name='Physics')
Hindi_df = pd.read_excel('Python_Assignment.xlsx',sheet_name='Hindi')
Economics_df = pd.read_excel('Python_Assignment.xlsx',sheet_name='Economics')
Music_df = pd.read_excel('Python_Assignment.xlsx',sheet_name='Music')

#aggregation of % for each subject
Maths_df['Maths_marks']=Maths_df['Theory Marks']+Maths_df['Numerical Marks']+Maths_df['Practical Marks']
Maths_df['Maths_marks_%']=(Maths_df['Maths_marks']/300)*100
Maths=Maths_df[['Roll No','Class','Maths_marks_%']]

Physics_df['Physics_marks']=Physics_df['Theory Marks']+Physics_df['Numerical Marks']+Physics_df['Practical Marks']
Physics_df['Physics_marks_%']=(Physics_df['Physics_marks']/300)*100
Physics=Physics_df[['Roll No','Class','Physics_marks_%']]

Hindi_df['Hindi_marks']=Hindi_df['Marks']
Hindi_df['Hindi_marks_%']=(Hindi_df['Hindi_marks']/100)*100
Hindi=Hindi_df[['Roll No','Class','Hindi_marks_%']]

Economics_df['Economics_marks']=Economics_df['Theory Marks']+Economics_df['Numerical Marks']
Economics_df['Economics_marks_%']=(Economics_df['Economics_marks']/200)*100
Economics=Economics_df[['Roll No','Class','Economics_marks_%']]

Music_df['Music_marks']=Music_df['Theory Marks']+Music_df['Practical Marks']
Music_df['Music_marks_%']=(Music_df['Music_marks']/200)*100
Music=Music_df[['Roll No','Class','Music_marks_%']]

#final df of all subjects % Marks
merge_df=pd.merge(Maths,Physics,on=['Roll No','Class'],how='outer')
merge_df=pd.merge(merge_df,Hindi,on=['Roll No','Class'],how='outer')
merge_df=pd.merge(merge_df,Economics,on=['Roll No','Class'],how='outer')
merge_df=pd.merge(merge_df,Music,on=['Roll No','Class'],how='outer')

# Q1
df=merge_df.fillna('NA')
#df.to_excel("output.xlsx",sheet_name='Q_1')

# Q2
# How many students in total are enrolled with the tuition provider?
ans1=len(merge_df)

# How many students have taken all the five subjects?
ans2=len(df[(df['Maths_marks_%']!='NA')&(df['Physics_marks_%']!='NA')&(df['Hindi_marks_%']!='NA')
       &(df['Economics_marks_%']!='NA')&(df['Music_marks_%']!='NA')])

#- Which class has the most number of students?
students=merge_df.groupby(['Class']).count()
r=students[(students['Roll No']==max(students['Roll No']))]
ans3=list(r.index.values)

# Which class has the highest average percentage of marks across all subjects?
c_df=merge_df
c_df=c_df.drop(['Roll No'], axis = 1)
avg_df=c_df.groupby(['Class']).mean()
s=avg_df.mean(axis=0)
ans4=s.idxmax()

# Which subject has the highest average percentage of marks across all classes?
c=avg_df.mean(axis=1)
ans5=c.idxmax()

#final Q2 df
df1=pd.DataFrame([['How many students in total are enrolled with the tuition provider?',ans1],
                  ['How many students have taken all the five subjects?',ans2],
                  ['Which class has the most number of students?',ans3],
                  ['Which class has the highest average percentage of marks across all subjects?',ans5],
                  ['Which subject has the highest average percentage of marks across all classes?',ans4]
                 ])
df1.rename(columns = {0: 'Questions', 1: 'Answers'}, inplace = True)

with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, sheet_name='Q_1')
    df1.to_excel(writer, sheet_name='Q_2')