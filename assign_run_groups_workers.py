# -*- coding: utf-8 -*-
"""
Created on Thu Sep 15 16:21:18 2016

@author: korinfox
"""
from import_registered_workers import import_registered_workers

direct = 'C:/Dropbox (Personal)/Autocross/BMWCCA/workerassignment/old worker assignment lists/' # directory with files
filename = 'Event 7 - Worker Assignments - 9-25-16 - Final.xls'

full_filename = direct + filename # concatenating filepath with filename to import workbook

# import registered workers from a spreadsheet
workers_tuple = import_registered_workers(full_filename)

workers_class = workers_tuple[0]
workers_novice = workers_tuple[1]

# get sizes of each class
class_sizes = {key: len(value) for key, value in workers_class.items()}
# trim any empty classes
class_sizes_nonzero = {key: value for key, value in class_sizes.items() if value != 0}

total_workers = sum(class_sizes_nonzero.values())
# for all possible combinations compute the class size splits

# given the length of class_sizes_nonzero, there will be 2^that many different combinations to try
# make this more elegant sometime

# use binary to use as indicator to assign groups methodically
# output should be four values? groups in 1, groups in 2, total workers in 1, total workers in 2

group_combos = [] # will be format of (list of classes in RGA, number of workers in RGA, list of classes in RGB, number of workers in RGB)

for i in range(2**len(class_sizes_nonzero)):
    group_assigner = bin(i)[2:].zfill(len(class_sizes_nonzero))
    run_group_a = []
    run_group_b = []
    run_group_a_size = 0
    run_group_b_size = 0
    for j in range(len(group_assigner)):

        if int(group_assigner[j]) == 1:
            run_group_a.append(class_sizes_nonzero.keys()[j])
            run_group_a_size = run_group_a_size + class_sizes_nonzero.values()[j]
        else:
            run_group_b.append(class_sizes_nonzero.keys()[j])
            run_group_b_size = run_group_b_size + class_sizes_nonzero.values()[j]
    individual_group_combo = (run_group_a,run_group_a_size,run_group_b,run_group_b_size)
    group_combos.append(individual_group_combo)
    print i


# now find combos of groups where the classes are almost half and half, say largest difference of 10%?

evenish_group_combos = [ind_groups for ind_groups in group_combos if abs(ind_groups[1] - ind_groups[3]) < 0.1*total_workers]

# Must have HARD worker constraints here

# At least one chair/co-chair in different run groups

chair_list = ['Audra Tella','William Brundige','Scott Baston']

# at least one computer qualified person in different run groups

computer_list = ['Matthew Cwieka','Matt Angle','Laura Rosen']