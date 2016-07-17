# -*- coding: utf-8 -*-
"""
Created on Sun Jul 17 10:49:47 2016

@author: jrr
"""

def parse_student_name(str):
    ''' given a student'name in the form last,first
        parse the string and return a tuple with (first last)
    '''
    names = str.split(',')
    if len(names) >= 2:
        return (names[0].strip(), names[1].strip())
    else:
        return names[0]
        
print(parse_student_name('Rowland, Jim'))
