# Marine InVEST: Area of Zones of Human Use
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
    ZonesDirectory = gp.GetParameterAsText(4)
    parameters.append("Directory of Human Use Layers: "+ ZonesDirectory)
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

PRegions_Diss = outputws + "PRegions_Diss.shp"

##############################################
###### COMMON FUNCTION AND CHECK INPUTS ######
##############################################

def checkGeometry(thedata):
    if gp.Describe(thedata).ShapeType <> "Polygon":
        raise Exception, "\nInvalid input: Zones must be polygons in order to calculate area."
    
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

def checkPRegionID(thestring):
    if thestring == 'Id':
        gp.AddError("Please use a different planning region identifier field.  The field name 'Id' is reserved for new shapefiles and conflicts with this tool's area calculations.")
        raise Exception

###########################################
######## CHECK INPUTS & DATA PREP #########
###########################################

gp.AddMessage("\nChecking and preparing the inputs...")
# check projection of inputs
ckProjection(PRegions)
checkPRegionID(PR_IDField)

gp.workspace = ZonesDirectory
fcList = gp.ListFeatureClasses("*", "all")
fc = fcList.Next()

ZonesLyrList = []
ZonesLyrAbbrevList = []
ZonesIDList = []
ZonesCount = 0

while fc:
    checkGeometry(fc)
    ckProjection(fc)
    ZonesLyrList.append(fc)
    fc0 = fc.replace(".", "")
    fc1 = fc0.replace("_", "")
    ZonesLyrAbbrevList.append(fc1[:7]) 
    ZonesCount = ZonesCount + 1
    ZonesIDList.append(ZonesCount)
    fc = fcList.Next()
del fc

gp.workspace = interws

# area of planning regions
gp.Dissolve_management(PRegions,  PRegions_Diss, PR_IDField)
PRegions_Diss = AddField(PRegions_Diss, "AREA", "FLOAT", "0", "0")
PRegions_Diss = AddField(PRegions_Diss, "P_REGION", "SHORT", "0", "0")
gp.CalculateField_management(PRegions_Diss, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")
gp.CalculateField_management(PRegions_Diss, "P_REGION", "1", "VB")
    
# create 'AreaStatsArray' to store calcs
AreaStatsArray = []
for i in range(gp.GetCount_management(PRegions_Diss)):
    AreaStatsArray.append([0]*(2+ZonesCount))

PRFieldList = []  
cur = gp.UpdateCursor(PRegions_Diss)
row = cur.Next()
count = 0
while row:
    PRFieldList.append(row.GetValue(PR_IDField))
    AreaStatsArray[count][0] = row.GetValue(PR_IDField)
    AreaStatsArray[count][1] = round(row.GetValue("AREA"), 2)
    count += 1
    row = cur.next()
del row, cur


gp.AddMessage("\nCalculating area of each human activity within each planning region...")  
gp.Extent = PRegions_Diss
gp.workspace = interws

# iterate through the zones
ZonesFolderName = "zones"
ZonesPath = os.path.join(interws, ZonesFolderName)
gp.CreateFolder_management(interws, ZonesFolderName)
              
for i in range(0,len(ZonesLyrList)):
    ZoneVector = ZonesDirectory+"\\"+ZonesLyrList[i]
    # check that zone doesn't have same field as 'PR_IDField', if so delete    
    fields = gp.ListFields(ZoneVector, PR_IDField)
    field_found = fields.Next()
    if field_found:
        pass
    else:   
        ZoneVector = AddField(ZoneVector, "ZONES", "SHORT", "0", "0") 
    gp.CalculateField_management(ZoneVector, "ZONES", "1", "VB")
    UnionExpr = ZoneVector+" 1; "+PRegions_Diss+" 2"
    gp.Union_analysis(UnionExpr, ZonesPath+"\\UnionZ"+str(i+1)+"PRegions.shp")
    ZoneSelect = ZonesPath+"\\Z"+str(i+1)+"_PRegions.shp"
    gp.Select_analysis(ZonesPath+"\\UnionZ"+str(i+1)+"PRegions.shp", ZoneSelect, "\"ZONES\" = 1 AND \"P_REGION\" = 1")
    ZoneDissolve = ZonesPath+"\\"+ZonesLyrAbbrevList[i]+"_Diss.shp"
    gp.Dissolve_management(ZoneSelect, ZoneDissolve, PR_IDField)
    ZoneDissolve = AddField(ZoneDissolve, "AREA", "FLOAT", "0", "0")
    gp.CalculateField_management(ZoneDissolve, "AREA", "!shape.area@squarekilometers!", "PYTHON", "")
    cur = gp.UpdateCursor(ZoneDissolve)
    row = cur.Next()
    while row:
        indexID = PRFieldList.index(row.GetValue(PR_IDField))
        AreaStatsArray[indexID][2+i] = round(row.GetValue("AREA"), 2)
        row = cur.next()
    del row, cur

# write to CSV
gp.AddMessage("\nCreating output CSV...")
AreaStatsCSV  = open(outputws+"AreaStats_"+ZonesScenarioStr+".csv", "wb")
writer = csv.writer(AreaStatsCSV, delimiter=',', quoting=csv.QUOTE_NONE)
count = -1
while count < gp.GetCount_management(PRegions_Diss):
    if count == -1:
        Line1List = ['PREGION', 'PREGION AREA']
        for j in range(0,len(ZonesLyrAbbrevList)):
            Line1List.append(ZonesLyrAbbrevList[j].upper())
        writer.writerow(Line1List)
        Line2List = ['(ID FIELD)', '(SQ. KM)']
        for k in range(0,ZonesCount):
            Line2List.append('(SQ. KM)')
        writer.writerow(Line2List)
    else:
        writer.writerow(AreaStatsArray[count])
    count += 1
AreaStatsCSV.close()
gp.AddMessage("CSV location:"+outputws+"AreaStats_"+ZonesScenarioStr+".csv") # print CSV location

# create parameter file
parameters.append("Script location: "+os.path.dirname(sys.argv[0])+"\\"+os.path.basename(sys.argv[0]))
parafile = open(outputws+"parameters_"+now.strftime("%Y-%m-%d-%H-%M")+".txt","w") 
parafile.writelines("AREA OF ZONES OF HUMAN USE\n")
parafile.writelines("__________________________\n\n")
for para in parameters:
    parafile.writelines(para+"\n")
    parafile.writelines("\n")
parafile.close()