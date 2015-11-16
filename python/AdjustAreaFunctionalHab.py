# Marine InVEST: Adjust Area of Functional Habitat
# Author: Gregg Verutes
# 05/06/15

# import modules
import sys, string, os, datetime, csv
import arcgisscripting
from math import *

# create the geoprocessor object
gp = arcgisscripting.create()
# set output handling
gp.OverwriteOutput = 1
# check out any necessary extensions
gp.CheckOutExtension("management")
gp.CheckOutExtension("analysis")
gp.CheckOutExtension("conversion")
gp.CheckOutExtension("spatial")

# error messages
msgArguments = "\nProblem with arguments."

try:
    # get parameters
    parameters = []
    now = datetime.datetime.now()
    parameters.append("Date and Time: "+ now.strftime("%Y-%m-%d %H:%M"))
    gp.workspace = gp.GetParameterAsText(0)
    parameters.append("Workspace: "+ gp.workspace)
    ZonesScenarioStr = gp.GetParameterAsText(1)
    parameters.append("Scenario Name: "+ ZonesScenarioStr)  
    PRegions = gp.GetParameterAsText(2)
    parameters.append("Planning Regions Layer: "+ PRegions)
    PR_IDField = gp.GetParameterAsText(3)
    parameters.append("Planning Regions Identifier Field: "+ PR_IDField)
    Hab1 = gp.GetParameterAsText(4)
    parameters.append("Habitat 1 Output from HRA: "+ Hab1)
    Hab2 = gp.GetParameterAsText(5)
    parameters.append("Habitat 2 Output from HRA: "+ Hab2)
    Hab3 = gp.GetParameterAsText(6)
    parameters.append("Habitat 3 Output from HRA: "+ Hab3)
    LowRisk = float(gp.GetParameterAsText(7))
    parameters.append("Low Risk Percent Reduction: "+ str(LowRisk))
    MedRisk = float(gp.GetParameterAsText(8))
    parameters.append("Medium Risk Percent Reduction: "+ str(MedRisk))
    HighRisk = float(gp.GetParameterAsText(9))
    parameters.append("High Risk Percent Reduction: "+ str(HighRisk))   
except:
    raise Exception, msgArguments + gp.GetMessages(2)

try:
    thefolders=["intermediate","Output"]
    for folder in thefolders:
        if not gp.exists(gp.workspace+folder):
            gp.CreateFolder_management(gp.workspace, folder)
except:
    raise Exception, "Error creating folders"

# local variables 
outputws = gp.workspace + os.sep + "Output" + os.sep
interws = gp.workspace + os.sep + "intermediate" + os.sep

Hab1Dissolve = interws + "Hab1Dissolve.shp"
unionHab1Zones = interws + "unionHab1Zones.shp"
Hab2Dissolve = interws + "Hab2Dissolve.shp"
unionHab2Zones = interws + "unionHab2Zones.shp"
Hab3Dissolve = interws + "Hab3Dissolve.shp"
unionHab3Zones = interws + "unionHab3Zones.shp"
PlanningRegions = outputws + "PlanningRegions.shp"
Hab1Zones = outputws + "Hab1Zones.shp"
Hab2Zones = outputws + "Hab2Zones.shp"
Hab3Zones = outputws + "Hab3Zones.shp"

##############################################
###### COMMON FUNCTION AND CHECK INPUTS ######
##############################################

def AddField(FileName, FieldName, Type, Precision, Scale):
    fields = gp.ListFields(FileName, FieldName)
    field_found = fields.Next()
    if field_found:
        gp.DeleteField_management(FileName, FieldName)
    gp.AddField_management(FileName, FieldName, Type, Precision, Scale, "", "", "NON_NULLABLE", "NON_REQUIRED", "")
    return FileName

def ckProjection(data):
    dataDesc = gp.describe(data)
    spatreflc = dataDesc.SpatialReference
    if spatreflc.Type <> 'Projected':
        gp.AddError(data +" does not appear to be projected.  It is assumed to be in meters.")
        raise Exception
    if spatreflc.LinearUnitName <> 'Meter':
        gp.AddError("This model assumes that "+data+" is projected in meters for area calculations.  You may get erroneous results.")
        raise Exception


###########################################
######## CHECK INPUTS & DATA PREP #########
###########################################

gp.AddMessage("\nChecking and preparing the inputs...")
ckProjection(PRegions)
ckProjection(Hab1)
if Hab2:
    ckProjection(Hab2)
if Hab3:
    ckProjection(Hab3)

gp.Dissolve_management(PRegions, PlanningRegions, PR_IDField)

# area of input 'Zones'
PlanningRegions = AddField(PlanningRegions, "AREA", "FLOAT", "0", "0")
gp.CalculateField_management(PlanningRegions, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")

# create 'AreaStatsArray' to store calcs
AreaStatsArray = []
for i in range(gp.GetCount_management(PlanningRegions)):
    AreaStatsArray.append([0]*14)


PR_IDFieldList = []  
cur = gp.UpdateCursor(PlanningRegions)
row = cur.Next()
count = 0
while row:
    PR_IDFieldList.append(row.GetValue(PR_IDField))
    AreaStatsArray[count][0] = row.GetValue(PR_IDField)
    AreaStatsArray[count][1] = round(row.GetValue("AREA"), 2)
    count += 1
    row = cur.next()
del row, cur

gp.AddMessage("\nCalculating and adjusting habitat area within each planning region...")  
gp.Extent = PlanningRegions

# habitat 1
PlanningRegions = AddField(PlanningRegions, "HAB1", "SHORT", "0", "0")
gp.CalculateField_management(PlanningRegions, "HAB1", "1", "VB")
gp.Dissolve_management(Hab1, Hab1Dissolve, "CLASSIFY")
Hab1Dissolve = AddField(Hab1Dissolve, "ZONES", "SHORT", "0", "0")
gp.CalculateField_management(Hab1Dissolve, "ZONES", "1", "VB")
UnionHab1Expr =  Hab1Dissolve+" 1; "+PlanningRegions+" 2"        
gp.Union_analysis(UnionHab1Expr, unionHab1Zones)
gp.Select_analysis(unionHab1Zones, Hab1Zones, "\"ZONES\" = 1 AND \"HAB1\" = 1")
Hab1Zones = AddField(Hab1Zones, "AREA", "FLOAT", "0", "0")
Hab1Zones = AddField(Hab1Zones, "AREA_ADJ", "FLOAT", "0", "0")
gp.CalculateField_management(Hab1Zones, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")
cur = gp.UpdateCursor(Hab1Zones)
row = cur.Next()
while row:
    indexID = PR_IDFieldList.index(row.GetValue(PR_IDField))
    if row.GetValue("CLASSIFY") == 'MED':
        row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(MedRisk/100.0))))
        AreaStatsArray[indexID][3] = AreaStatsArray[indexID][3] + row.GetValue("AREA")
    elif row.GetValue("CLASSIFY") == 'HIGH':
        row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(HighRisk/100.0))))
        AreaStatsArray[indexID][4] = AreaStatsArray[indexID][4] + row.GetValue("AREA")
    else:
        row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(LowRisk/100.0))))
        AreaStatsArray[indexID][2] = AreaStatsArray[indexID][2] + row.GetValue("AREA")
    AreaStatsArray[indexID][11] = AreaStatsArray[indexID][11] + row.GetValue("AREA_ADJ")
    cur.UpdateRow(row)
    row = cur.next()
del row, cur

# habitat 2
if Hab2:
    PlanningRegions = AddField(PlanningRegions, "HAB2", "SHORT", "0", "0")
    gp.CalculateField_management(PlanningRegions, "HAB2", "1", "VB")
    gp.Dissolve_management(Hab2, Hab2Dissolve, "CLASSIFY")
    Hab2Dissolve = AddField(Hab2Dissolve, "ZONES", "SHORT", "0", "0")
    gp.CalculateField_management(Hab2Dissolve, "ZONES", "1", "VB")
    UnionHab2Expr =  Hab2Dissolve+" 1; "+PlanningRegions+" 2"        
    gp.Union_analysis(UnionHab2Expr, unionHab2Zones)
    gp.Select_analysis(unionHab2Zones, Hab2Zones, "\"ZONES\" = 1 AND \"HAB2\" = 1")
    Hab2Zones = AddField(Hab2Zones, "AREA", "FLOAT", "0", "0")
    Hab2Zones = AddField(Hab2Zones, "AREA_ADJ", "FLOAT", "0", "0")
    gp.CalculateField_management(Hab2Zones, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")
    cur = gp.UpdateCursor(Hab2Zones)
    row = cur.Next()
    while row:
        indexID = PR_IDFieldList.index(row.GetValue(PR_IDField))
        if row.GetValue("CLASSIFY") == 'MED':
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(MedRisk/100.0))))
            AreaStatsArray[indexID][6] = AreaStatsArray[indexID][6] + row.GetValue("AREA")
        elif row.GetValue("CLASSIFY") == 'HIGH':
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(HighRisk/100.0))))
            AreaStatsArray[indexID][7] = AreaStatsArray[indexID][7] + row.GetValue("AREA")
        else:
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(LowRisk/100.0))))
            AreaStatsArray[indexID][5] = AreaStatsArray[indexID][5] + row.GetValue("AREA")

        AreaStatsArray[indexID][12] = AreaStatsArray[indexID][12] + row.GetValue("AREA_ADJ")
        cur.UpdateRow(row)
        row = cur.next()
    del row, cur

# habitat 3
if Hab3:
    PlanningRegions = AddField(PlanningRegions, "HAB3", "SHORT", "0", "0")
    gp.CalculateField_management(PlanningRegions, "HAB3", "1", "VB")
    gp.Dissolve_management(Hab3, Hab3Dissolve, "CLASSIFY")
    Hab3Dissolve = AddField(Hab3Dissolve, "ZONES", "SHORT", "0", "0")
    gp.CalculateField_management(Hab3Dissolve, "ZONES", "1", "VB")
    UnionHab3Expr =  Hab3Dissolve+" 1; "+PlanningRegions+" 2"        
    gp.Union_analysis(UnionHab3Expr, unionHab3Zones)
    gp.Select_analysis(unionHab3Zones, Hab3Zones, "\"ZONES\" = 1 AND \"HAB3\" = 1")
    Hab3Zones = AddField(Hab3Zones, "AREA", "FLOAT", "0", "0")
    Hab3Zones = AddField(Hab3Zones, "AREA_ADJ", "FLOAT", "0", "0")
    gp.CalculateField_management(Hab3Zones, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")
    cur = gp.UpdateCursor(Hab3Zones)
    row = cur.Next()
    while row:
        indexID = PR_IDFieldList.index(row.GetValue(PR_IDField))
        if row.GetValue("CLASSIFY") == 'MED':
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(MedRisk/100.0))))
            AreaStatsArray[indexID][9] = AreaStatsArray[indexID][9] + row.GetValue("AREA")
        elif row.GetValue("CLASSIFY") == 'HIGH':
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(HighRisk/100.0))))
            AreaStatsArray[indexID][10] = AreaStatsArray[indexID][10] + row.GetValue("AREA")
        else:
            row.SetValue("AREA_ADJ", (row.GetValue("AREA")-(row.GetValue("AREA")*(LowRisk/100.0))))
            AreaStatsArray[indexID][8] = AreaStatsArray[indexID][8] + row.GetValue("AREA")
        AreaStatsArray[indexID][13] = AreaStatsArray[indexID][13] + row.GetValue("AREA_ADJ")
        cur.UpdateRow(row)
        row = cur.next()
    del row, cur

# write to CSV
gp.AddMessage("\nWriting area calculations and adjustments in CSV...")
gp.AddMessage("CSV Location: "+outputws+"AreaStats_"+ZonesScenarioStr+".csv")
AreaStatsCSV  = open(outputws+"AreaStats_"+ZonesScenarioStr+".csv", "wb")
writer = csv.writer(AreaStatsCSV, delimiter=',', quoting=csv.QUOTE_NONE)
count = -1
while count < gp.GetCount_management(PlanningRegions):
    if count == -1:
        if Hab2 and Hab3:
            writer.writerow(['PREGION', 'PREGION AREA', 'H1 LOW RISK', 'H1 MED RISK', 'H1 HIGH RISK', 'H2 LOW RISK', 'H2 MED RISK', 'H2 HIGH RISK', 'H3 LOW RISK', 'H3 MED RISK', 'H3 HIGH RISK', 'H1 ADJ AREA', 'H2 ADJ AREA', 'H3 ADJ AREA'])
            writer.writerow(['(ID FIELD)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)'])
        elif Hab2:
            writer.writerow(['PREGION', 'PREGION AREA', 'H1 LOW RISK', 'H1 MED RISK', 'H1 HIGH RISK', 'H2 LOW RISK', 'H2 MED RISK', 'H2 HIGH RISK', '', '', '', 'H1 ADJ AREA', 'H2 ADJ AREA', ''])
            writer.writerow(['(ID FIELD)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '', '', '', '(SQ. KM)', '(SQ. KM)', ''])
        else:
            writer.writerow(['PREGION', 'PREGION AREA', 'H1 LOW RISK', 'H1 MED RISK', 'H1 HIGH RISK', '', '', '', '', '', '', 'H1 ADJ AREA', '', ''])
            writer.writerow(['(ID FIELD)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '(SQ. KM)', '', '', '', '', '', '', '(SQ. KM)', '', ''])
    else:
        writer.writerow(AreaStatsArray[count])
    count += 1
AreaStatsCSV.close()

# create parameter file
parameters.append("Script location: "+os.path.dirname(sys.argv[0])+"\\"+os.path.basename(sys.argv[0]))
parafile = open(gp.GetParameterAsText(0)+"\\Output\\parameters_"+now.strftime("%Y-%m-%d-%H-%M")+".txt","w") 
parafile.writelines("ADJUST FUNCTIONAL HABITAT AREA\n")
parafile.writelines("______________________________\n\n")
for para in parameters:
    parafile.writelines(para+"\n")
    parafile.writelines("\n")
parafile.close()