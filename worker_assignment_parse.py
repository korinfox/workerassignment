# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 23:21:35 2016

@author: Justin
"""

import xlrd
import os
import numpy as np

direct = 'C:/Dropbox (Personal)/Autocross/BMWCCA/workerassignment/old worker assignment lists/' # directory with files
file_list = os.listdir(direct)

for k in range(0,len(file_list)):

  
    filename = file_list[k] # specific filename for the excel spreadsheet with worker assignments
    
    full_filename = direct + filename # concatenating filepath with filename to import workbook
    
    wkbk = xlrd.open_workbook(full_filename) # open the work book
    
    wkass = wkbk.sheet_by_index(0) # import the first sheet which has worker assignments
    
    # import whole sheet into a list of lists in python, columns then rows
    
    num_cols = wkass.ncols
    num_rows = wkass.nrows
    
    wkass_lol = [[] for i in range(num_cols)]
    for i in range(num_cols):
        wkass_lol[i] = wkass.col(i)
        
    for i in range(num_cols):
        for j in range(num_rows):
            wkass_lol[i][j] = wkass_lol[i][j].value.encode('ascii','replace')
    
    # make an empty list of lists for worker positions, so we can iterate over this faster, add special cases separately?
    # need special cases for finding 'Gate' and 'Lunch'
    
    position_name_list = ['Early Gate','Course Design','Tech','Early Setup','Worker Boss','Control','Computer','Timing','Announcer','Starter','Staging','Lunch','Gate','Lunch GATE','Sound','Corner Boss']
    
    position_out = [[] for i in range(len(position_name_list))]
    
    # start finding work assignment names and save into those empty lists by name/string
    
    for i in range(0,len(position_name_list)):
    #for i in range(0,4):
    
        matching = [n for (n,e) in enumerate(wkass_lol[2]) if position_name_list[i] in e]
    
        for mat_pos in matching:
            position_out[i].append(wkass_lol[1][mat_pos]) # looking to the left of the position name
            position_out[i].append(wkass_lol[3][mat_pos]) # looking to the right of the position name
        
        position_out[i] = filter(None, position_out[i]) # removing empty entries
          
    # Controlling for ambiguity between early gate and gate
    position_out[12] = list(set(position_out[12]) - set(position_out[0]))
    # Subtracting for ambiguity between Lunch and Lunch GATE
    position_out[11] = list(set(position_out[11]) - set(position_out[13]))
    
    # find corner workers
    
    position_name_list.append('Corner Peon')
    position_out.append([])
    
    # corner workers will be anything between the first instance of Corner and the second instance of Class, MINUS the Corner Bosses
    
    corner_matching = [n for (n,e) in enumerate(wkass_lol[2]) if 'Corner' in e]
    class_matching = [n for (n,e) in enumerate(wkass_lol[2]) if 'Class' in e]
    
    for i in range(int(corner_matching[0]),int(class_matching[1])):
        position_out[16].append(wkass_lol[1][i])
        position_out[16].append(wkass_lol[3][i])
        
    position_out[16] = filter(None, position_out[16])
    
    position_out[16] = list(set(position_out[16]) - set(position_out[15]))
    
    # find instructors, anything between Instructors and first instance of Class
    position_name_list.append('Instructors')
    position_out.append([])
    
    instructor_matching = [n for (n,e) in enumerate(wkass_lol[2]) if 'Instructors' in e]
    
    # fuck me, instructor is misspelled as "Instuctor" in someo f these worksheets
    if not instructor_matching:
        instructor_matching = [n for (n,e) in enumerate(wkass_lol[2]) if 'Instuctors' in e]
    
    for i in range(int(instructor_matching[0])+1,int(class_matching[0])):
        position_out[17].append(wkass_lol[1][i])
        position_out[17].append(wkass_lol[3][i])
    
    position_out[17] = filter(None, position_out[17])
    
    position_dict = dict(zip(position_name_list, position_out))
    
    save_name = 'worker_dict_' + str(k) + '.npy'
    
    np.save(save_name,position_dict)
    
    del position_name_list
    del position_out
    del position_dict
    del wkass_lol    

# extra bullshit snippets of code

#early_gate = []
#bus_setup = []
#tech_worker = []
#course_design = []
#worker_boss = []
#control = []
#computer = []
#timing = []
#announcer = []
#starter = []
#staging = []
#gate = []
#lunch = []
#sound = []
#corner_boss = []
#corner_peon = []
#
## tech workers
#
#matching = [s for s in wkass_lol[2] if "Tech" in s]
#
#for mat_str in matching:
#    tech_worker.append(wkass_lol[1][wkass_lol[2].index(mat_str)])
#    tech_worker.append(wkass_lol[3][wkass_lol[2].index(mat_str)])
#    
#tech_worker = filter(None, tech_worker)