# -*- coding: utf-8 -*-
"""
Created on Mon Jul 11 10:08:06 2016

Explore Nikil's data with Pandas

@author: jrr
"""

import pandas as pd

CENSUS1_PATH = '/Users/jrr/Documents/Stevens/StudentDB/Python/census 1 data file.xlsx'
CENSUS1_WRKSHT = 'census1'
CENSUS3_PATH = '/Users/jrr/Documents/Stevens/StudentDB/Python/census 3 data file.xlsx'
CENSUS3_WRKSHT = 'census3'

REQUIRED_COURSES = {
    'SFEN' : ['SSW 540', 'SSW 564', 'SSw 565', 'SSW 555', 'SWW 567', 'SSW533', 'SSW 690', 'SSW 695']
}

REQUIRED_CREDITS  = {
    'SFEN' : 30
}

def load_wksht(path,sheetName):
    ''' load census1 data as provided by Nikhil 
        path is the file path
        sheetName is the name of the worksheet with the data
        return a pandas dataframe with the data or raise an exception '''
    try:
        return pd.ExcelFile(path).parse(sheetName)
        
    except FileNotFoundError:
        print("Can't load census1 data from: {}".format(path))
 #   except xlrd.biffh.XLRDError:
        print("invalid worksheet name: {}".format(sheetName))
        
        
def map_semester(row):
    ''' map_semester maps semester of form 'YYYY[FSABY] to a sortable value.
        row is a row that includes 'Semester'
    '''
    smap ={ 'F': '1', 'S':'2', 'A':'3','B':'4', 'Y':'0'}
    if len(row['Semester']) == 5:
        return row['Semester'][0:4] + smap.get(row['Semester'][4],'0')
    else:
        return row['Semester']
        
def map_semester_back(row):
    ''' map_semester maps semester of form 'YYYY[FSABY] to a sortable value.
        row is a row that includes 'Semester'
    '''
    smap ={ '0':'Y', '1':'F', '2':'S', '3':'A','4':'B'}
    if len(row['Semester_sort']) == 5:
        return row['Semester_sort'][0:4] + smap.get(row['Semester_sort'][4],'?')
    else:
        return row['Semester_sort']
        
def write_to_excel_init(fileName):
    ''' get ready to write to an excel spreadsheet '''
    return pd.ExcelWriter(fileName, engine='xlsxwriter')
    
def write_to_excel_worksheet(writer, df, sheetName):
    df.to_excel(writer, sheet_name=sheetName, index=False)
    
def write_to_excel_close(writer):
    writer.save()
    
        

# load the census spreadsheets
census1 = load_wksht(CENSUS1_PATH, CENSUS1_WRKSHT)
census3 = load_wksht(CENSUS3_PATH, CENSUS3_WRKSHT)

'''
# create a unique list of cwids from census1/students
cwids = census1.SID.drop_duplicates()
assert cwids.count() == census1.SID.nunique()
'''
# add a new Semester_sort column to census1 with a sortable version of the Semester
census1['Semester_sort'] = census1.apply(map_semester,axis=1)

# build up a students dataframe with CWID, Name, Email, Entry/firstSemester
students = census1.groupby(['SID','Deg'])['Name'].min().reset_index()  # get the student's SID, degree, and name

# add first semester and last semester to students
first_semester = census1.groupby(['SID','Deg'])['Semester_sort'].min().reset_index()  # get the first semester for the student and degree
first_semester['First_semester'] = first_semester.apply(map_semester_back, axis=1)
first_semester.drop(['Semester_sort'], axis=1, inplace=True)

last_semester = census1.groupby(['SID','Deg'])['Semester_sort'].max().reset_index()  # get the first semester for the student and degree
last_semester['Last_semester'] = last_semester.apply(map_semester_back, axis=1)
last_semester.drop(['Semester_sort'], axis=1, inplace=True)

# merge students, first and last semester
students2 = pd.merge(students, first_semester, left_on=['SID','Deg'], right_on=['SID','Deg'])
students3 = pd.merge(students2, last_semester, left_on=['SID','Deg'], right_on=['SID','Deg'])

# add the Col, Deg, Dept, and Maj1 to students from the Last semester
students4 = pd.merge(students3, census1[['SID', 'Col', 'Deg', 'Dept', 'Maj1', 'Semester']], how='left', left_on=['SID', 'Deg', 'Last_semester'], right_on=['SID', 'Deg', 'Semester'])
#assert students3.SID.nunique() == census1.SID.nunique()

# rename and drop columns in students3 to "final" state
students4.rename(columns={'SID':'CWID'}, inplace=True)
students4.drop(['Semester'], axis=1, inplace=True)

# Next, generate grades by merging students and grades
sem_deg = census1.groupby(['SID', 'Semester'])['Deg'].max().reset_index()  # get a mapping of SID,Sem
grades = pd.merge(census3, sem_deg, how='left', left_on=['CWID','Semester'], right_on=['SID','Semester'])
grades2 = pd.merge(grades, students4, how='left', left_on=['CWID', 'Deg'], right_on=['CWID', 'Deg'])
grades2.drop(['SID','date'], axis=1, inplace=True)

outputFile = 'studentdb.xlsx' #input('Enter .xlsx file name: ')
writer = write_to_excel_init(outputFile)
write_to_excel_worksheet(writer, students4, 'Students')
write_to_excel_worksheet(writer, grades2, 'Grades')
write_to_excel_close(writer)

