# -*- coding: utf-8 -*-
"""
Created on Thu Sep 15 13:10:58 2016

@author: korinfox
"""
# function here

# (bmw_classes, novice_people_only) = import_registered_workers( full_filename )

def import_registered_workers( full_filename ):
    import xlrd
    
#    direct = 'C:/Dropbox (Personal)/Autocross/BMWCCA/workerassignment/old worker assignment lists/' # directory with files
#    filename = 'Event 7 - Worker Assignments - 9-25-16 - Final.xls'
    
#    full_filename = direct + filename # concatenating filepath with filename to import workbook
    
    wkbk = xlrd.open_workbook(full_filename) # open the work book
    
    wkreg = wkbk.sheet_by_index(1) # import the second sheet which has registered workers
    
    # import worker reg as list of lists, rows then columns
    
    num_cols = wkreg.ncols
    num_rows = wkreg.nrows
    
    wkreg_lol = [[] for i in range(num_rows)]
    for i in range(num_rows):
        wkreg_lol[i] = wkreg.row(i)
        
    for i in range(num_rows):
        for j in [1, 2, 10]: # only care to do this for columns 1, 2 and 10
            wkreg_lol[i][j] = wkreg_lol[i][j].value.encode('ascii','replace')
    
    wkreg_name_novice_class = [[] for i in range(num_rows-1)]
    
    # get only rows corresponding to name novice/instructor and class out of the worksheet
    for i in range(len(wkreg_name_novice_class)):
        wkreg_name_novice_class[i].append(wkreg_lol[i+1][1])
        wkreg_name_novice_class[i].append(wkreg_lol[i+1][2])
        wkreg_name_novice_class[i].append(wkreg_lol[i+1][10])
        
    bmw_classes = dict(A=[],B=[],C=[],D=[],E=[],F=[],G=[],I=[],J=[],L=[],M=[])
    
    novice_people_inst = dict(Yes=[],No=[],Inst=[])
    
    for i in range(len(wkreg_name_novice_class)):
        working_row = wkreg_name_novice_class[i]
        
        for class_key in bmw_classes:
            if working_row[2] == class_key:
                bmw_classes[class_key].append(working_row[0])
                
        for novice_key in novice_people_inst:
            if working_row[1] == novice_key:
                novice_people_inst[novice_key].append(working_row[0])
    
    novice_people_only = dict(Yes=[],No=[])
    novice_people_only['Yes'].append(novice_people_inst['Yes'])
    novice_people_only['No'].append(novice_people_inst['No'] + novice_people_inst['Inst'])
    
    return (bmw_classes, novice_people_only)
    
# importing registered workers based on sheet 2 of the worker assignments
# Name is in column B (1) and Vehicle class is in column K (10)
# Novice or not is in column C or (2)

#import xlrd
#
#direct = 'C:/Dropbox (Personal)/Autocross/BMWCCA/workerassignment/old worker assignment lists/' # directory with files
#filename = 'Event 7 - Worker Assignments - 9-25-16 - Final.xls'
#
#full_filename = direct + filename # concatenating filepath with filename to import workbook
#
#wkbk = xlrd.open_workbook(full_filename) # open the work book
#
#wkreg = wkbk.sheet_by_index(1) # import the second sheet which has registered workers
#
## import worker reg as list of lists, rows then columns
#
#num_cols = wkreg.ncols
#num_rows = wkreg.nrows
#
#wkreg_lol = [[] for i in range(num_rows)]
#for i in range(num_rows):
#    wkreg_lol[i] = wkreg.row(i)
#    
#for i in range(num_rows):
#    for j in [1, 2, 10]: # only care to do this for columns 1, 2 and 10
#        wkreg_lol[i][j] = wkreg_lol[i][j].value.encode('ascii','replace')
#
#wkreg_name_novice_class = [[] for i in range(num_rows-1)]
#
## get only rows corresponding to name novice/instructor and class out of the worksheet
#for i in range(len(wkreg_name_novice_class)):
#    wkreg_name_novice_class[i].append(wkreg_lol[i+1][1])
#    wkreg_name_novice_class[i].append(wkreg_lol[i+1][2])
#    wkreg_name_novice_class[i].append(wkreg_lol[i+1][10])
#    
#bmw_classes = dict(A=[],B=[],C=[],D=[],E=[],F=[],G=[],I=[],J=[],L=[],M=[])
#
#novice_people_inst = dict(Yes=[],No=[],Inst=[])
#
#for i in range(len(wkreg_name_novice_class)):
#    working_row = wkreg_name_novice_class[i]
#    
#    for class_key in bmw_classes:
#        if working_row[2] == class_key:
#            bmw_classes[class_key].append(working_row[0])
#            
#    for novice_key in novice_people_inst:
#        if working_row[1] == novice_key:
#            novice_people_inst[novice_key].append(working_row[0])
#
#novice_people_only = dict(Yes=[],No=[])
#novice_people_only['Yes'].append(novice_people_inst['Yes'])
#novice_people_only['No'].append(novice_people_inst['No'] + novice_people_inst['Inst'])

# output bmw_classes and novice_people_only
