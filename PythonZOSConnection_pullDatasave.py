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
# Press 'Interactive Extension' Button under 'Programming' in Zemax before 
# starting this script !!!!!!!!!!!!!!
#
# Notes
#
# The python project and script was tested with the following tools:
#       Python 3.4.3 for Windows (32-bit) (https://www.python.org/downloads/) - Python interpreter
#       Python for Windows Extensions (32-bit, Python 3.4) (http://sourceforge.net/projects/pywin32/) - for COM support
#       Microsoft Visual Studio Express 2013 for Windows Desktop (https://www.visualstudio.com/en-us/products/visual-studio-express-vs.aspx) - easy-to-use IDE
#       Python Tools for Visual Studio (https://pytools.codeplex.com/) - integration into Visual Studio
#
# Note that Visual Studio and Python Tools make development easier, however this python script should should run without either installed.


# make sure the Python wrappers are available for the COM client and
# interfaces
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
# don't forhet to press 'Interactive Extension' button in Zemax before running this!!!!!!!!!!!!
# written by TVA 20180609 

TheAnalyses = TheSystem.Analyses

##############################################################################
## FFT analysis
#
#newWin = TheAnalyses.New_FftMtf()
## Settings
#newWin_Settings = newWin.GetSettings()
#newWin_SettingsCast = CastTo(newWin_Settings,'IAS_FftMtf')
#newWin_SettingsCast.MaximumFrequency = 10
#newWin_SettingsCast.SampleSize = constants.SampleSizes_S_256x256
## Run Analysis
#newWin.ApplyAndWaitForCompletion()
#
## Get Analysis Results
#newWin_Results = newWin.GetResults()
#newWin_ResultsCast = CastTo(newWin_Results,'IAR_')
#
## Read and plot data series
## NOTE: numpy functions are used to unpack and plot the 2D tuple for Sagittal & Tangential MTF
## You will need to import the numpy module to get this part of the code to work
#
##test data extraction
#data0 = newWin_ResultsCast.GetDataSeries(0)
#
#colors = ('b','g','r','c', 'm', 'y', 'k')
#for seriesNum in range(0,newWin_ResultsCast.NumberOfDataSeries,1):
#    data = newWin_ResultsCast.GetDataSeries(seriesNum)
#    x = np.array(data.XData.Data)
#    y = np.array(data.YData.Data)
#
#    plt.plot(x[:],y[:,0],color=colors[seriesNum])
#    plt.plot(x[:],y[:,1],linestyle='--',color=colors[seriesNum])
#
## format the plot
#plt.title('FFT MTF: ')
#plt.xlabel('Spatial Frequency in cycles per mm')
#plt.ylabel('Modulus of the OTF')
#plt.legend([r'$0^\circ$ tangential',r'$0^\circ$ sagittal',r'$14^\circ$ tangential',r'$14^\circ$ sagittal',r'$20^\circ$ tangential',r'$20^\circ$ sagittal'])
#plt.grid(True)
##plt.legend('0^\circ tangential', '0^\circ sagittal', '14^\circ tangential','14^\circ sagittal', '20^\circ tangential', '20^\circ sagittal')
#plt.show()
#
##############################################################################
## Spot diagram analysis
#
## Open Spot Diagram to See Result!
#newSpot = TheAnalyses.New_StandardSpot()
#print("Spot has analysis specific settings? ", newSpot.HasAnalysisSpecificSettings)  # True; no ModifySettings
#newSettings = newSpot.GetSettings()
#spotSet = CastTo(newSettings, "IAS_Spot")  # Cast to IAS_Spot interface; enables access to Spot Diagram properties
#spotSet.RayDensity = 15
#newSpot.ApplyAndWaitForCompletion()

##############################################################################
## PSF map
#
#newPSF = TheAnalyses.New_HuygensPsf()
#newPSF.ApplyAndWaitForCompletion()
#newPSF_Results = newPSF.GetResults()
#data = newPSF_Results.GetDataGrid(0)
#print(data.Nx)
#print(data.Ny)
#print(data.MinX)
#print(data.Dx)
#print(data.Values)
#
#RMS = 1e3*np.std(data.Values)
#PV = 1e3*max(max(data.Values))-min(min(data.Values))
#
## plot
#fig1 = plt.figure(1)
#ax1 = fig1.add_subplot(111)
#im1 = ax1.imshow(data.Values, cmap='hot')
#ax1.set_title('%.1f nm rms - %.1f nm ptv' %(RMS,PV))
#ax1.set_xlabel('x (pixels)')
#ax1.set_ylabel('y (pixels)')
##plt.pcolormesh(matrix, cmap=plt.cm.get_cmap('plasma'))
#plt.colorbar(im1)
#plt.show()


#############################################################################
# WFE map

#Open WFE 
newWFE = TheAnalyses.New_WavefrontMap()
#if settings need to be changed
analysisSettings =  newWFE.GetSettings()
newWFE_SettingsCast = CastTo(analysisSettings,'IAS_')

# Run Analysis
newWFE.ApplyAndWaitForCompletion()

# Get Analysis Results
newWFE_Results = newWFE.GetResults()

#newWFE_ResultsCast = CastTo(newWFE_Results,'IAR_')
data = newWFE_Results.GetDataGrid(0)
print(data.Nx)
print(data.Ny)
print(data.MinX)
print(data.Dx)
print(data.Values)

# convert from Tuple to array
dataValues = np.array(data.Values)

# Calculate mask, where 'nan' values 0, otherwise 1 then invert 1 and 0
mask = np.isnan(data.Values).astype(int)
mask = abs(mask-1)

# plot
fig2 = plt.figure(2)
ax2 = fig2.add_subplot(111)
im2 = ax2.imshow(1000*dataValues, cmap='hot')
ax2.set_title('%.1f nm rms - %.1f nm ptv' %(RMS(1000*dataValues, mask),PtV(1000*dataValues, mask)))
ax2.set_xlabel('x (pixels)')
ax2.set_ylabel('y (pixels)')
#plt.pcolormesh(matrix, cmap=plt.cm.get_cmap('plasma'))
plt.colorbar(im2)
plt.show()


## Modifysettings with cfg
## modify the cfg file 
#cfgFile = "C:\\Users\\Agocs\\Documents\\Python Tibor\\metis wfe python\\wfe_map_save.cfg"
#analysisSettings.SaveTo(cfgFile)  # Save current settings to a cfg file
## MODIFYSETTINGS are defined in ZPL help files: The Programming Tab > About the ZPL > Keywords
##Wavefront Map:
##WFM_SAMP: The sampling, use 1 for 32, 2 for 64, etc. 
##WFM_FIELD: The field number.
##WFM_WAVE: The wavelength number. 
# 
#analysisSettings.ModifySettings(cfgFile, "WFM_SAMP", "1")
#analysisSettings.ModifySettings(cfgFile, "WFM_FIELD", "1")
#analysisSettings.ModifySettings(cfgFile, "WFM_WAVE", "1")
#analysisSettings.LoadFrom(cfgFile)  # Load in the newly modified settings
#
## If you want to overwrite your default CFG, save over it with newly modified CFG file:
##  CFG_fullname = "C:\\Users\\zachary.Derocher\\Documents\\Zemax\\Configs\\POP.CFG"
##  import shutil
##  shutil.copy(cfgFile, CFG_fullname)
#
#newWFE.ApplyAndWaitForCompletion()
