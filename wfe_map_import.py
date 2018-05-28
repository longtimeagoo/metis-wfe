# -*- coding: utf-8 -*-
"""
Created on Thu May 24 12:24:41 2018

@author: agocs
"""

import codecs
import matplotlib.pyplot as plt
import numpy as np
#from numpy import loadtxt


#read and plot the WFS map from Zemax

folder_input = 'D:/Users/agocs/Desktop/WORK FILES 20180525/metis wfe python/data/'
folder_output = 'D:/Users/agocs/Desktop/WORK FILES 20180525/metis wfe python/data_ok/'
#file = '01.dat'

for i in range(10):
    
    matrix = []

    if i<9:
        file = '0' + str(i+1) 
    else: 
        file = str(i+1)
    
    #open(folder_input + file, 'r')
    
    with codecs.open(folder_input + file + '.DAT', encoding='utf-16') as f:
        for line in f:
            try:            
                line = line.strip()
                if len(line) > 0:
                    matrix.append([float(n) for n in line.split()])
            except ValueError:
                print('ValueError for line:')
                print(line)
    
    print(matrix)
    
    plt.imshow(matrix)
    plt.colorbar()
            
    print()
    print(np.std(matrix))
    
    np.savetxt(folder_output + file + '.txt.gz', matrix)