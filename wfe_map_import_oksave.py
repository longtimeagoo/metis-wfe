# -*- coding: utf-8 -*-
"""
Created on Thu May 24 12:24:41 2018

@author: agocs
"""

import codecs
import matplotlib.pyplot as plt
import numpy as np
#from numpy import loadtxt

#############################################
# read and plot the WFS map from Zemax
#############################################

folder = 'D:/Users/agocs/Documents/Python Tibor/metis wfe python/test/'
file = 'tezi01ansi'
# select whether unicode or ansi (uni_flag=1 means unicode, others ansi)
uni_flag = 0

matrix = []

if uni_flag==1:     
    with codecs.open(folder + file + '.txt', encoding='utf-16') as f:
        for line in f:
            try:            
                line = line.strip()
                if len(line) > 0:
                    matrix.append([float(n) for n in line.split()])
            except ValueError:
                print('ValueError for line:')
                print(line)
            except: 
                print('other unknwown error')
else:
    with open(folder + file + '.txt', 'r') as f: 
        for line in f:
            try:            
                line = line.strip()
                if len(line) > 0:
                    matrix.append([float(n) for n in line.split()])
            except ValueError:
                print('ValueError for line:')
                print(line)        
            except: 
                print('other unknwown error')
            
# print the matrix
print(matrix)

RMS = 1e3*np.std(matrix)
PV = 1e3*max(max(matrix))-min(min(matrix))

# plot the results 
fig = plt.figure(1)
ax1 = fig.add_subplot(111)
im1 = ax1.imshow(matrix, cmap='hot')
ax1.set_title('%.1f nm rms - %.1f nm ptv' %(RMS,PV))
ax1.set_xlabel('x (pixels)')
ax1.set_ylabel('y (pixels)')
#plt.pcolormesh(matrix, cmap=plt.cm.get_cmap('plasma'))
plt.colorbar(im1)
plt.show()

# save the matrix in compressed format
np.savetxt(folder + file + '.txt.gz', matrix)
