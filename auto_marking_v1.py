import pandas as pd
import openpyxl
import difflib

#Set Path for AnswerKey and StudentAnswer files
answer_key='C:\\Auto-Marking-System-v2\\answer_key.xlsx'
answer_file='C:\\Auto-Marking-System-v2\\student1_ans_file.xlsx'

#Create a data frame and get input from AnswerKey
df=pd.read_excel(answer_key) 
df=df.set_index('qno')

#Prepare StudentAnswer file for marking
writer = pd.ExcelWriter(answer_file,engine='openpyxl')
writer.book = openpyxl.load_workbook(answer_file)
current_file=writer.book

#Auto marking starts here

for i in df.index:
    comment=" "
    marks=0
    current_sheet=current_file[df.loc[i,'worksheet']]
    
    #Fetching student's answer
    v=str(current_sheet[df.loc[i,'cell_address']].value)   
         
    if v=="None":
        std_ans=""
    elif v[0]=='=':
        std_ans=v[1:]
    else:
        std_ans=v[0:]  
        
        
    #Finding matching ratio    
    df.loc[i,'student_answer']=std_ans
    ans=df.loc[i,'answer']
    sequence = difflib.SequenceMatcher(isjunk=None,a=std_ans.upper(),b=ans.upper())
    match_ratio =sequence.ratio()*100
    df.loc[i,'match_ratio']=match_ratio
    #Giving Marks based on matching ratio  
    mks=df.loc[i,'marks']
    if df.loc[i,'answer_functions']=="NIL":
        marks=0
        comment="None"
    elif match_ratio== 100:
        marks=mks
        comment="Matching Ratio 100"
    else:
        if match_ratio < 100 and match_ratio >= 80:
            marks=marks+(mks * 0.50)
            comment="Matching Ratio between 80-99"
        else:
            comment="Matching Ratio < 80"
            
        strfn=df.loc[i,'answer_functions'].split(",")
       
        t=std_ans.upper()
        if len(strfn)==1 and t.find(strfn[0].upper())!= -1:
            marks=marks+(mks * 0.25)
            
        elif len(strfn)==2:
             if(strfn[0]== strfn[1] and t.count(strfn[0].upper())== 2):
                 marks=marks+(mks * 0.25)
                 
             if(strfn[0]== strfn[1] and t.count(strfn[0].upper())== 1):
                 marks=marks+(mks * 0.125)
                 
             if(strfn[0] != strfn[1] and t.find(strfn[0].upper())!= -1):
                 marks=marks+(mks * 0.125)
                
             if(strfn[0] != strfn[1] and t.find(strfn[1].upper())!= -1):
                 marks=marks+(mks * 0.125)
                
        elif len(strfn)> 2:
            strfn1=list(set(strfn))
            if len(strfn1)==1:
                if t.count(strfn1[0].upper()) > 2:
                    marks=marks+(mks * 0.25)
                   
                elif t.count(strfn1[0].upper()) > 0 and t.count(strfn1[0].upper()) <= 2:
                    marks=marks+(mks * 0.125)
                    
            elif len(strfn1)==2:   
                if(t.find(strfn1[0].upper())!= -1):
                    marks=marks+(mks * 0.125) 
                    
                if(t.find(strfn1[1].upper())!= -1):
                    marks=marks+(mks * 0.125) 
                    

    df.loc[i,'student_marks']=marks
    df.loc[i,'comment']=comment
  
df.at['Total','student_marks'] = df['student_marks'].sum()    
#Remove mark sheet if already exists in student answer excel file
if 'marks_sheet' in current_file.sheetnames:
    current_file.remove(current_file['marks_sheet'])
    
#Moving data from data frame to student answer excel file     
df.to_excel(writer, 'marks_sheet')

#Format the student answer excel file
current_file['marks_sheet'].column_dimensions['D'].width =45
current_file['marks_sheet'].column_dimensions['G'].width =45  
current_file['marks_sheet'].column_dimensions['E'].width =20    
current_file['marks_sheet'].column_dimensions['B'].width =20    
current_file['marks_sheet'].column_dimensions['C'].width =10  
for i in range(1,35):
    current_file['marks_sheet'].row_dimensions[i].height = 25
    
#Finally save the student answer excel file
writer.save()



