from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
from win32com.client import gencache
import matplotlib.pyplot as plt
import numpy as np

def PtV(P,mask):
        return max(P[np.where(mask==1)])-min(P[np.where(mask==1)])
                             
def RMS(P,mask):
        return np.std(P[np.where(mask==1)])

# IMPORTANT!!!
# Press 'Interactive Extension' Button under 'Programming' in Zemax before starting this !!!

# make sure the Python wrappers are available for the COM client and interfaces
gencache.EnsureModule('{EA433010-2BAC-43C4-857C-7AEAC4A8CCE0}', 0, 1, 0)
gencache.EnsureModule('{F66684D7-AAFE-4A62-9156-FF7A7853F764}', 0, 1, 0)
# Note - the above can also be accomplished using 'makepy.py' in the
# following directory:
#      {PythonEnv}\Lib\site-packages\wind32com\client\
# Also note that the generate wrappers do not get refreshed when the
# COM library changes.
# To refresh the wrappers, you can manually delete everything in the
# cache directory:
#	   {PythonEnv}\Lib\site-packages\win32com\gen_py\*.*

TheConnection = EnsureDispatch("ZOSAPI.ZOSAPI_Connection")
if TheConnection is None:
    raise Exception("Unable to intialize COM connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

################################################################################
# written by TVA 20180609 

#folder_input = 'D:/Users/agocs/Documents/Python Tibor/metis wfe python/zemax files/'
folder_output = 'D:/Users/agocs/Documents/Python Tibor/metis wfe python/_v12/'
      
#surfaces to import 
#surf = [150, 201, 242, 248, 289, 320, 322, 459, 502, 511]

#names of surf
name = ['CFO FP2', 'IMG-LM PP1', 'IMG-LM DET', 'IMG-NQ PP1', 'IMG-NQ DET', 
        'LMS PP1', 'LMS IFU', 'LMS DET', 'SCAO PYR', 'SCAO DET']

dataValuesAll = []

# calculate NCPA-s? then all pupils have to be circular (set up in cfg)
ncpaFlag = 1

TheAnalyses = TheSystem.Analyses

for i in range(10): 

    TheSystem.MCE.SetCurrentConfiguration(i+1)
    ThePrimarySystem = TheApplication.PrimarySystem
    TheSystemData = ThePrimarySystem.SystemData

    # use wave 2 (it should be in the cfg of the WFE map too):
    wave = 2
    w = TheSystemData.Wavelengths.GetWavelength(wave).Wavelength
    
    # Open WFE 
    newWFE = TheAnalyses.New_WavefrontMap()
    # if settings need to be changed
    # if settings are not changed, the cfg setting will be used
    # so cfg can be re-saved in Zemax if needed and it will be ok!!!
    newWFE_Settings =  newWFE.GetSettings()
    newWFE_SettingsCast = CastTo(newWFE_Settings,'IAS_WavefrontMap')
    #newWFE_SettingsCast.Surface.SetSurfaceNumber(surf[i])
    newWFE_SettingsCast.Wavelength.SetWavelengthNumber(wave)
    
    # Run Analysis
    newWFE.ApplyAndWaitForCompletion()
    
    # Get Analysis Results
    newWFE_Results = newWFE.GetResults()
    
    # get the data, print some 
    data = newWFE_Results.GetDataGrid(0)
    print(data.Nx)
    print(data.Ny)
    #print(data.MinX)
    #print(data.Dx)
    #print(data.Values)
    
    # convert from Tuple to array
    dataValues = np.array(data.Values)
    
    # calculate in nm
    dataValues = 1000*w*dataValues
    dataValuesAll.append(dataValues)
    
    # Calculate mask, where 'nan' values 0, otherwise 1 then invert 1 and 0
    mask = np.isnan(data.Values).astype(int)
    mask = abs(mask-1)
    
    # plot output and save 
    if i<9:
        file = '0' + str(i+1) 
    else: 
        file = str(i+1)
    
    file_txtgz = folder_output + file + '.txt.gz'
    file_jpg = folder_output + file + '.jpg'
    file_pdf = folder_output + file + '.pdf'
    
    fig = plt.figure()
    ax1 = fig.add_subplot(111)
    im1 = ax1.imshow(dataValues, cmap='jet')
    ax1.set_title('%s - %.1f nm rms - %.1f nm ptv' %(name[i],RMS(dataValues, mask),PtV(dataValues, mask)))
    ax1.set_xlabel('x (pixels)')
    ax1.set_ylabel('y (pixels)')
    plt.colorbar(im1)
    
    plt.show()

    #plt.savefig(file_jpg)
    #plt.close()

    #np.savetxt(file_txtgz, dataValues)
    

# Calculate NCPA-s - only work if the pupils are circular (set it up in the cfg file)
   
if ncpaFlag==1: 
    namencpa = ['NCPA(IMGLM-SCAO)', 'NCPA(IMGNQ-SCAO)', 'NCPA(IMGLM-LMSIFU)', 
        'NCPA(IMGLM-LMSDET)', 'NCPA(LMSIFU-SCAO)', 'NCPA(LMSDET-SCAO)']    
    ncpa = []
    ncpa.append(dataValuesAll[2]-dataValuesAll[9])
    ncpa.append(dataValuesAll[4]-dataValuesAll[9])
    ncpa.append(dataValuesAll[2]-dataValuesAll[6])
    ncpa.append(dataValuesAll[2]-dataValuesAll[7])
    ncpa.append(dataValuesAll[6]-dataValuesAll[9])
    ncpa.append(dataValuesAll[7]-dataValuesAll[9])
    
    for i in range(11,17):
        # plot output and save 
 
        file = str(i+1)
        
        file_txtgz = folder_output + file + '.txt.gz'
        file_jpg = folder_output + file + '.jpg'
        file_pdf = folder_output + file + '.pdf'
        
        # Calculate mask, where 'nan' values 0, otherwise 1 then invert 1 and 0
        ncpatemp = ncpa[i-11]
        mask = np.isnan(ncpatemp).astype(int)
        mask = abs(mask-1)  
                
        fig = plt.figure()
        ax1 = fig.add_subplot(111)
        im1 = ax1.imshow(ncpatemp, cmap='jet')
        ax1.set_title('%s - %.1f nm rms - %.1f nm ptv' %(namencpa[i-11],RMS(ncpatemp, mask),PtV(ncpatemp, mask)))
        ax1.set_xlabel('x (pixels)')
        ax1.set_ylabel('y (pixels)')
        plt.colorbar(im1)
        
        plt.show()
    
        #plt.savefig(file_jpg)
        #plt.close()
    
        #np.savetxt(file_txtgz, dataValues)        


