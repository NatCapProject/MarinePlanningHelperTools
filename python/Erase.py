# Marine InVEST: Erase
# Author: Gregg Verutes
# 05/04/15

# import modules
import sys, string, os, datetime
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
msgEraseFunction = "\nError utilizing GIS tools to erase input feature class."

try:
    try:
        # get parameters
        inputFC = gp.GetParameterAsText(0)
        eraseFC = gp.GetParameterAsText(1)
        outputFC = gp.GetParameterAsText(2)
    except:
        raise Exception, msgArguments + gp.GetMessages(2)

    gp.workspace = outputFC[:-(1+len(outputFC.split("\\")[-1]))]
    unionFC = gp.workspace + "\\interUnion.shp"
    gp.Extent = "MAXOF"

    def AddField(FileName, WghtFieldName, Type, Precision, Scale):
        fields = gp.ListFields(FileName, WghtFieldName)
        field_found = fields.Next()
        if field_found:
            gp.DeleteField_management(FileName, WghtFieldName)
        gp.AddField_management(FileName, WghtFieldName, Type, Precision, Scale, "", "", "NON_NULLABLE", "NON_REQUIRED", "")
        return FileName

    def checkGeometry(thedata, Type, Message):
        if gp.Describe(thedata).ShapeType <> Type:
            raise Exception, "\nInvalid input: "+thedata+"\n"+Message+" must be of geometry type "+Type+"."

    def difference(a, b): # show whats in list b which isn't in list a
        return list(set(b).difference(set(a)))
    
    def intersect(a, b): # return the intersection of two lists
        return list(set(a) & set(b))

    try:
        # check geometry of inputs
        checkGeometry(inputFC, "Polygon", "Input feature class")
        checkGeometry(eraseFC, "Polygon", "Erase feature class")

        # add marker to eraseFC and unite the two FCs
        fields = gp.ListFields(inputFC, "ERASE")
        field_found = fields.Next()
        if field_found:
            gp.DeleteField_management(inputFC, "ERASE")
        
        eraseFC = AddField(eraseFC, "ERASE", "SHORT", "0", "0")
        gp.CalculateField_management(eraseFC, "ERASE", "1", "VB")
        UnionExpr = inputFC+" 1; "+eraseFC+" 2"        
        gp.Union_analysis(UnionExpr, unionFC)

        # select features where "ERASE = 0"
        gp.Select_analysis(unionFC, outputFC, "\"ERASE\" = 0")

        # create a list of all fields in eraseFC and inputFC
        inputFieldList = []
        fields = gp.ListFields(inputFC, "*")
        fc_field = fields.Next()
        while fc_field:
            inputFieldList.append(fc_field.name)
            fc_field = fields.Next()
        del fc_field
        eraseFieldList = []
        fields = gp.ListFields(eraseFC, "*")
        fc_field = fields.Next()
        while fc_field:
            eraseFieldList.append(fc_field.name)
            fc_field = fields.Next()
        del fc_field

        diffList = difference(inputFieldList, eraseFieldList)
        intersectList = intersect(inputFieldList, eraseFieldList)

        for i in range(0,len(intersectList)):
            intersectList[i] = intersectList[i]+"_1"
        eraseFieldList = diffList + intersectList
    
        EraseFieldExpr = eraseFieldList[0]
        for i in range(1,len(eraseFieldList)):
            EraseFieldExpr = EraseFieldExpr+";"+str(eraseFieldList[i])
        gp.DeleteField_management(outputFC, EraseFieldExpr)
        
    except:
        gp.AddError(msgEraseFunction)
        raise Exception

    # delete intermediate data
    gp.delete_management(unionFC)
    gp.AddMessage("")
    del gp

except Exception, ErrorDesc:
    gp.AddMessage(gp.GetMessages())