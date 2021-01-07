
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
import numpy as np
import re
import math
from scipy.signal import argrelextrema

# -------------------------- User Can Edit ----------------------------------#

use_All_CSV_Files = True # If False, Populate the CV_CSV_Data_List Yourself in the Next Section Below
data_Directory = "../NASA Project Cortisol/Prussian Blue/01-5-2021 Cortisol DPV/" # The Path to the Folder with the CSV Files

# Use CHI's Predicted Peak Values (Must be in CSV/Excel)
useCHIPeaks = False  # Only Performs Baseline Subtractiomn if CHI Didnt Label a Peak
# If useCHIPeaks is False: THESE ARE SUPER IMPORTANT PARAMETERS THAT WILL CHANGE YOUR PEAK
order = 3        # Order of the Polynomial Fit in Baseline Calculation
Iterations = 15  # The Number of Polynomial Fit and Subtractions in Baseline Calculation

# Specify Figure Asthetics
numSubplotWidth = 4
figWidth = 25
figHeight = 13

# Make Subplots of Only Final Current or Show All Steps
displayOnlyBaselineSubtraction = False

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

# Specify Which Files to Ignore
ignoreFiles = []

# Create Output Folder if None
try:
    outputData = data_Directory +  "Peak_Current_Plots/"
    os.mkdir(outputData)
# Else, Continue On
except:
    pass

# ---------------------- User Does NOT Have to Edit -------------------------#
# ---------------------------------------------------------------------------#
# ---------------------------------------------------------------------------#
# ----------------------------- Functions -----------------------------------#

def normalize(point, low, high):
        return (point-low)/(high-low)

def saveplot(fig1, axisLimits, base, outputData):
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
    

def getBase(potential, currentReal, Iterations, order):
    current = currentReal.copy()
    for _ in range(Iterations):
        fitI = np.polyfit(potential, current, order)
        baseline = np.polyval(fitI, potential)
        for i in range(len(current)):
            if current[i] > baseline[i]:
                current[i] = baseline[i]
    return baseline
        

# ---------------------------------------------------------------------------#
# -------------------- Extract and Plot the Ip Data -------------------------#

# Create One Plot with All the DPV Curves
fig, ax = plt.subplots(math.ceil(len(CV_CSV_Data_List)/numSubplotWidth), numSubplotWidth, sharey=False, sharex = True, figsize=(figWidth,figHeight))
fig.tight_layout(pad=3.0)
data = {}  # Store Results ina Dictionary for Later Analaysis
# For Each CSV File, Extract the Important Data and Plot
for figNum, CV_CSV_Data in enumerate(sorted(CV_CSV_Data_List)):
    if CV_CSV_Data in ignoreFiles:
        continue
    
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
    # ----------------------- Extract Run Info ------------------------------#
    
    # Set Initial Variables from last Run to Zero
    deltaV = None; endVolt = None; initialVolt = None; Vp = None; Ip = None
    # Loop Through the Info Section and Extract the Needxed Run Info from Excel
    for cell in Main['A']:
        # Get Cell Value
        cellVal = cell.value
        if cellVal == None:
            continue
        
        # Find the deltaV for Each Step (Volts)
        if cellVal.startswith("Incr E (V) = "):
            deltaV = float(cellVal.split(" = ")[-1])
        # Find the Final Voltage
        elif cellVal.startswith("Final E (V) = "):
            endVolt = float(cellVal.split(" = ")[-1])
        # Find the Initial Voltage
        elif cellVal.startswith("Init E (V) = "):
            initialVolt = float(cellVal.split(" = ")[-1])
        # If Peak Found by CHI, Get Peak Potential
        elif cellVal.startswith("Ep = "):
            Vp = float(cellVal.split(" = ")[-1][:-1])
        # If Peak Found by CHI, Get Peak Current
        elif cellVal.startswith("ip = "):
            Ip = float(cellVal.split(" = ")[-1][:-1])
        elif cellVal == "Potential/V":
            startDataRow = cell.row + 2
            break
    # Find the X Axis Width
    xRange = (endVolt - initialVolt)
    # Find Point/Scan
    pointsPerScan = int(xRange/deltaV)
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
        
    # Plot the Initial Data
    fig1 = plt.figure(2+figNum) # Leaving 2 Figures Free for Other plots
    plt.plot(potential, current, label="True Data: " + base, color='C0')
    
    # If We use the CHI Peaks, Skip Peak Detection
    if useCHIPeaks and Ip != None and Vp != None:
        # Set Axes Limits
        axisLimits = [min(current) - min(current)/10, max(current) + max(current)/10]
        
    # Else, Perform baseline Subtraction to Find the peak
    else:
        # Get Baseline
        baseline = getBase(potential, current, Iterations, order)
        
        # Plot Subtracted baseline
        baselineCurrent = current - baseline
        plt.plot(potential, baselineCurrent, label="Current After Baseline Subtraction", color='C2')
        plt.plot(potential, baseline, label="Baseline Current", color='C1')  
        
        # Find Where Data Begins to Deviate from the Edges
        minimums = argrelextrema(baselineCurrent, np.less)[0]
        stopInitial = minimums[0]
        stopFinal = minimums[-1]
        
        # Get the Peak Current (Max Differenc between the Data and the Baseline)
        IpIndex = np.argmax(baselineCurrent[stopInitial:stopFinal+1])
        Ip = baselineCurrent[stopInitial+IpIndex]
        Vp = potential[stopInitial+IpIndex]
    
        # Plot the Peak Current (Verticle Line) for Visualization
        axisLimits = [min(baselineCurrent) - min(baselineCurrent)/10, max(current) + max(current)/10]
        plt.axvline(x=Vp, ymin=normalize(baseline[stopInitial+IpIndex], axisLimits[0], axisLimits[1]), ymax=normalize(float(Ip+baseline[stopInitial+IpIndex]), axisLimits[0], axisLimits[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
    
    # Save Figure
    saveplot(fig1, axisLimits, base, outputData)
    
    # Keep Running Subplots Order
    if numSubplotWidth == 1 and len(CV_CSV_Data_List) == 1:
        currentAxes = ax
    elif numSubplotWidth == 1:
        currentAxes = ax[figNum]
    elif numSubplotWidth == len(CV_CSV_Data_List):
        currentAxes = ax[figNum]
    elif numSubplotWidth > 1:
        currentAxes = ax[figNum//numSubplotWidth][figNum%numSubplotWidth]
    else:
        print("numSubplotWidth CANNOT be < 1. Currently it is: ", numSubplotWidth)
        exit
    
    # Plot Data in Subplots
    if useCHIPeaks and Ip != None and Vp != None:
        currentAxes.plot(potential, current, label="True Data: " + base, color='C0')
        currentAxes.axvline(x=Vp, ymin=normalize(max(current) - Ip, currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), ymax=normalize(max(current), currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
        currentAxes.legend(loc='upper left')  
    elif displayOnlyBaselineSubtraction:
        currentAxes.plot(potential, baselineCurrent, label="Current After Baseline Subtraction", color='C1')
        currentAxes.axvline(x=Vp, ymin=normalize(0, currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), ymax=normalize(float(Ip), currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
        currentAxes.legend(loc='upper left')  
    else:
        currentAxes.plot(potential, current, label="True Data: " + base, color='C0')
        currentAxes.plot(potential, baselineCurrent, label="Current After Baseline Subtraction", color='C2')
        currentAxes.axvline(x=Vp, ymin=normalize(baseline[stopInitial+IpIndex], currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), ymax=normalize(float(Ip+baseline[stopInitial+IpIndex]), currentAxes.get_ylim()[0], currentAxes.get_ylim()[1]), linewidth=2, color='r', label="Peak Current: " + "%.4g"%Ip)
        currentAxes.plot(potential, baseline, label="Baseline Current", color='C1')  
        currentAxes.legend(loc='best')  


    currentAxes.set_xlabel("Potential (V)")
    currentAxes.set_ylabel("Current (Amps)")
    currentAxes.set_title(base)
    
    # Save Data in a Dictionary for Plotting Later
    data[base] = Ip
    
            
# ---------------------------------------------------------------------------#
# --------------------- Plot and Save the Data ------------------------------#

saveSubplot(fig)
plt.title(base + " DPV Graph") # Need this Line as we Change the Title When we Save Subplots
plt.show() # Must be the Last Line

# ---------------------------------------------------------------------------#
# -------------- Specific Plotting Method for This Data ---------------------#
# ----------------- USER SPECIFIC (USER SHOULD EDIT) ------------------------#


fig = plt.figure(0)
#fig.tight_layout(pad=3) #tight margins
fig.set_figwidth(6.5)
#ax = fig.add_axes([0.1, 0.1, 0.7, 0.9])
for i,filename in enumerate(sorted(data.keys())):
    # Extract Data from Name
    stringDigits = re.findall(r'\d+', filename) 
    digitsInName = list(map(int, stringDigits))
    if len(digitsInName) == 2:
        concentration = digitsInName[0]
        timePoint = digitsInName[1]
    elif len(digitsInName) == 1:
        concentration = 0
        timePoint = digitsInName[0]
    else:
        print("Found Too Many Numbers in the FileName")
        exit
    print(filename, timePoint, concentration)
    
    # Get Peak Current
    Ip = data[filename]
    
    
    if i == 8:
        i += 1
        time = []
        current = []
    
    if i%2 == 0:
        time = [timePoint]
        current  = [Ip]
    else:
        time.append(timePoint)
        current.append(Ip)
        
        # Plot Ip
        plt.plot(time, current, 'o-', label=filename.split("-")[0])
    
    
# Plot Curves
plt.title("Time Dependant DPV Peak Current: Cortisol")
plt.xlabel("Time (minutes)")
plt.ylabel("DPV Peak Current (Amps)")
plt.legend(loc=9, bbox_to_anchor=(1.2, 1))
plt.savefig(outputData + "Time Dependant DPV Curve Cortisol.png", dpi=300)
plt.show()
      






"""
Deleted Code:
    
    # Fit Lines to Ends of Graph
    m0, b0 = np.polyfit(potential[0:edgeCollectionLeft], current[0:edgeCollectionLeft], 1)
    mf, bf = np.polyfit(potential[-edgeCollectionRight:-1], current[-edgeCollectionRight:-1], 1)
    potentialNumpy = np.array(potential)
    y0 = m0*potentialNumpy+b0
    yf = mf*potentialNumpy+bf
    
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
    """
    