
"""
Need to Install on the Anaconda Prompt:
    $ conda install openpyxl
    $ conda install seaborn
    $ conda install scipy
"""
import csv
import os
import openpyxl as xl
import matplotlib.pyplot as plt
from scipy.interpolate import CubicSpline
import numpy as np
import re
import math

# -------------------------- User Can Edit ----------------------------------#

use_All_CSV_Files = True # If False, Populate the CV_CSV_Data_List Yourself in the Next Section Below
data_Directory = "../NASA Project Cortisol/Tryptophan/tryptophan MIP/" # The Path to the Folder with the CSV Files
startDataRow = 22        # The First CSV/Excel Row with Data (otential, Current); The Same for All Files

# SUPER IMPORTANT PARAMETER THAT WILL CHANGE YOUR PEAK MAGNITUDE
errorVal = 0.02          # While Fitting Edges, Start Interpolating at the Following Error
edgeCollectionLeft = 20  # Number of Points on Left Hand Side to Fit
edgeCollectionRight = 20 # Number of points on the Right Hand Side to Fit

# Specify Figure Asthetics
numSubplotWidth = 4
figWidth = 25
figHeight = 13

# ---------------------------------------------------------------------------#
# --------------------- Specify/Find File Names -----------------------------#

if use_All_CSV_Files:
    CV_CSV_Data_List = []
    for file in os.listdir(data_Directory):
        if file.endswith(".csv"):
            CV_CSV_Data_List.append(file)
else:
    CV_CSV_Data_List = [
         'New PB Heat 100C (1).csv',
         'New PB Heat 100C + Sonicate (2).csv',
         'New PB Heat 100C + Sonicate + Pellet of Centrifuge (3).csv',
         'New PB Heat 100C + Sonicate + Supernatant of Centrifuge (3).csv',
         'New PB Heat 100C + Sonicate + Supernatant of Centrifuge + Filter (4).csv',
         ]
    
    # Check to see if the Inputed CSV Files Exist
    for CV_CSV_Data in CV_CSV_Data_List:    
        if not os.path.isfile(data_Directory + CV_CSV_Data):
            print("The File ", data_Directory + CV_CSV_Data," Mentioned Does NOT Exist")
            exit()

# Create Output Folder
outputData = data_Directory +  "Peak_Current_Plots/"
os.system("mkdir '" + outputData + "'")

# ---------------------- User Does NOT Have to Edit -------------------------#
# ---------------------------------------------------------------------------#
# ---------------------------------------------------------------------------#
# ----------------------------- Functions -----------------------------------#

def normalize(point, low, high):
        return (point-low)/(high-low)

def saveplot(fig1):
    # Plot and Save
    plt.title(base + " DPV Graph")
    plt.xlabel("Potential (V)")
    plt.ylabel("Current (Amps)")
    plt.ylim(axisLimits)
    plt.legend()
    fig1.savefig(outputData + base + ".png", dpi=300)
    
    
def saveSubplot(fig):
    # Plot and Save
    plt.title("Subplots of all DPV")
    #fig.legend(bbox_to_anchor=(.5, 1))
    #plt.subplots_adjust(hspace=0.5, wspace=0.5)
    fig.savefig(outputData + "subplots.png", dpi=300)

# ---------------------------------------------------------------------------#
# -------------------- Extract and Plot the Ip Data -------------------------#

# For Eaqch CSV File, Extract the Important Data and Plot
fig, ax = plt.subplots(math.ceil(len(CV_CSV_Data_List)/numSubplotWidth), numSubplotWidth, sharey=False, sharex = True, figsize=(figWidth,figHeight))
fig.tight_layout(pad=3.0)
data = {}
for i,CV_CSV_Data in enumerate(sorted(CV_CSV_Data_List)):    
    
    # ----------------- Convert Data to Excel Format ------------------------#
    
    # Rename the File with an Excel Extension
    base = os.path.splitext(CV_CSV_Data)[0]
    excel_file = data_Directory + base + ".xlsx"
    # If the File is Not Already Converted: Convert
    if not os.path.isfile(excel_file):
        # Make Excel WorkBook
        wb = xl.Workbook()
        ws = wb.active
        # Write to Excel WorkBook
        with open(data_Directory + CV_CSV_Data) as f:
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                ws.append(row)
        # Save as New Excel File
        wb.save(excel_file)
    else:
        print("You already renamed the '.csv' to '.xlsx'")
    
    # Load Data from New Excel File
    WB = xl.load_workbook(excel_file) 
    WB_worksheets = WB.worksheets
    Main = WB_worksheets[0]
        
    # -----------------------------------------------------------------------#
    # -------------------- Find Ip Data and Plot ----------------------------#
    
    # Get Potential, Current Data from Excel
    potential = []
    current = []
    for cell in Main['A'][startDataRow - 1:]:
        # Break out of Loop if no More Data (edge effect if someone edits excel)
        if cell.value == None:
            break
        # Find the Potential and Current Data points
        row = cell.row - 1
        potential.append(float(cell.value))
        current.append(float(Main['B'][row].value))
        
    # Plot the Data
    fig1 = plt.figure(2+i)
    plt.plot(potential, current, label="True Data: " + base)    
    
    # Fit Lines to Ends of Graph
    m0, b0 = np.polyfit(potential[0:edgeCollectionLeft], current[0:edgeCollectionLeft], 1)
    mf, bf = np.polyfit(potential[-edgeCollectionRight:-1], current[-edgeCollectionRight:-1], 1)
    potentialNumpy = np.array(potential)
    y0 = m0*potentialNumpy+b0
    yf = mf*potentialNumpy+bf
    
    plt.plot(potential, y0)
    plt.plot(potential, yf)
    
    # Find Where Data Begins to Deviate from the Lines
    stopInitial = np.argwhere(abs(((y0-current)/current)) < errorVal)[-1][0]
    stopFinal = np.argwhere(abs(((yf-current)/current)) < errorVal)[0][0]
    
    # Get the Points inside the Peak
    potentialEnds = potential[0:stopInitial] + potential[stopFinal:-1]
    currentEnds = current[0:stopInitial] + current[stopFinal:-1]
    
    # Fit the Peak with a Cubic Spline
    cs = CubicSpline(potentialEnds, currentEnds)
    xs = np.arange(potential[0], potential[-1], (potential[-1]-potential[0])/len(potential))
    
    # Get the Peak Current (Max Differenc between the Data and the Spline/Background)
    peakCurrents = current[stopInitial:stopFinal+1] - cs(potential[stopInitial:stopFinal+1])
    IpIndex = np.argmax(peakCurrents)
    Ip = peakCurrents[IpIndex]
    Vp = potential[IpIndex+stopInitial]
    
    # Plot Fit
    plt.plot(potential, cs(xs), label="Spline Interpolation")
    axisLimits = [min(current) - min(current)/10, max(current) + max(current)/10]
    
    # Plot the Peak Current (Verticle Line) for Visualization
    plt.axvline(x=Vp, ymin=normalize(float(cs(Vp)), axisLimits[0], axisLimits[1]), ymax=normalize(float(cs(Vp)+Ip), axisLimits[0], axisLimits[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
    saveplot(fig1)
    
    # Keep Running Subplots
    if numSubplotWidth == 1:
        currentAxes = ax[i//numSubplotWidth]
    elif numSubplotWidth == len(CV_CSV_Data_List):
        currentAxes = ax[i]
    elif numSubplotWidth > 1:
        currentAxes = ax[i//numSubplotWidth][i%numSubplotWidth]
    else:
        print("numSubplotWidth CANNOT be < 1. Currently it is: ", numSubplotWidth)
        exit
    currentAxes.plot(potential, current, label="True Data")  
    currentAxes.plot(potential, cs(xs), label="Spline Interpolation")
    currentAxes.axvline(x=Vp, ymin=normalize(float(cs(Vp)), currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), ymax=normalize(float(cs(Vp)+Ip), currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
    currentAxes.set_xlabel("Potential (V)")
    currentAxes.set_ylabel("Current (Amps)")
    currentAxes.set_title(base)
    currentAxes.legend()  
    
    # Save Data in a Dictionary for Plotting Later
    data[base] = Ip
    
            
# ---------------------------------------------------------------------------#
# --------------------- Plot and Save the Data ------------------------------#

saveSubplot(fig)
plt.title(base + " DPV Graph") # Need this Line as we Change the Title When we Save Subplots
plt.show() # Must be the Last Line

# -------------- Specific Plotting Method for This Data----------------------#

fig = plt.figure(0)
#fig.tight_layout(pad=3) #tight margins
fig.set_figwidth(6.5)
#ax = fig.add_axes([0.1, 0.1, 0.7, 0.9])
for i,filename in enumerate(sorted(data.keys())):
    # Extract Data from Name
    stringDigits = re.findall(r'\d+', filename) 
    digitsInName = list(map(int, stringDigits)) 
    concentration = digitsInName[0]
    timePoint = digitsInName[1]
    print(filename, timePoint, concentration)
    
    # Get Peak Current
    Ip = data[filename]
    
    if i%2 == 0:
        time = [timePoint]
        current  = [Ip]
    else:
        time.append(timePoint)
        current.append(Ip)
        
        # Plot Ip
        plt.plot(time, current, 'o-', label=filename.split("-7")[0])
    
    
# Plot Curves
plt.title("Time Dependant DPV Peak Current: NIP")
plt.xlabel("Time (minutes)")
plt.ylabel("DPV Peak Current (Amps)")
plt.legend(loc=9, bbox_to_anchor=(1.2, 1))
plt.savefig(outputData + "Time Dependant DPV NIP.png", dpi=300)
plt.show()
    
    




    