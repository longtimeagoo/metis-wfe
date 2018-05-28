# -*- coding: utf-8 -*-
"""
Created on Mon May 28 10:21:05 2018

@author: agocs
"""

import matplotlib.pyplot as plt
from numpy import loadtxt

folder = 'D:/Users/agocs/Desktop/WORK FILES 20180525/metis wfe python/data_ok/'

for i in range(10): 
    
    if i<9:
        file = '0' + str(i+1) 
    else: 
        file = str(i+1)
        
    matrix = loadtxt(folder + file + '.txt.gz')
    plt.imshow(matrix)
