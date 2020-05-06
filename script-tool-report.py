# Version: 2.0.0
# Autor: Luis Romero
# Web: https://luishromero.com
# Date: 6/05/2020

"""
This Arcpy script creates a xlsx file that shows the number of
null and blank rows by field of a gdb. It creates a sheet by table
and feature class, inside each sheet populate extra information.

**This tool requires an active ArcGIS PRO licence to run
"""

from arcpy.da import SearchCursor
import arcpy
import xlsxwriter
import os

# get params
gdbPath = arcpy.GetParameterAsText(0)
excelName = arcpy.GetParameterAsText(1)

# set env
arcpy.env.workspace = gdbPath

# create xlsx
outWorkbook = xlsxwriter.Workbook(excelName)

# function used to build reports from feature classes
def fc_stats():
    outSheet = outWorkbook.add_worksheet(fc[0:30])
    outSheet.set_column(0, 4, 15)
    totalRows = arcpy.GetCount_management(fc)
    spatialRef = arcpy.Describe(fc).spatialReference
    fields = arcpy.ListFields(fc)
    stats_fields = []
    out_geom = "memory" + "\\" + str(fc) + "_" + "geom"
    arcpy.management.CheckGeometry(fc, out_geom)
    totalGeom = arcpy.management.GetCount(out_geom)
    output = "memory" + "\\" + str(fc)
    outSheet.write(0, 0, "NAME")
    outSheet.write(0, 1, fc)
    outSheet.write(1, 0, "TYPE")
    outSheet.write(1, 1, "Feature Class")
    outSheet.write(2, 0, "GCS name")
    outSheet.write(2, 1, spatialRef.name)
    outSheet.write(3, 0, "GCS type")
    outSheet.write(3, 1, spatialRef.type)
    outSheet.write(4, 0, "ROWS")
    outSheet.write(4, 1, int(str(totalRows)))
    outSheet.write(5, 0, "FIELDS")
    outSheet.write(5, 1, int(str(len(fields))))
    outSheet.write(6, 0, "GEOM ERROR")
    outSheet.write(6, 1, int(str(totalGeom)))
    outSheet.write(8, 0, "FIELD")
    outSheet.write(8, 1, "ALIAS")
    outSheet.write(8, 2, "TYPE")
    outSheet.write(8, 3, "COUNT NULL")
    outSheet.write(8, 4, "COUNT BLANK")
    arcpy.management.Delete(out_geom)
    for field in fields:
        if field.type not in ("OID", "Geometry"):
            outSheet.write(fields.index(field)+7, 0, field.name)
            outSheet.write(fields.index(field)+7, 1, field.aliasName)
            outSheet.write(fields.index(field)+7, 2, field.type)
            stats_fields.append([field.name, "COUNT"])
        if field.type not in ("OID", "Geometry", "Double", "Integer", "SmallInteger", "Single"):
            out_fc = "memory" + "\\" + str(fc) + "_" + str(field.name)
            expression = str(field.name) + ' IN (\'\', \' \')'
            arcpy.Select_analysis(fc, out_fc, expression)
            totalBlank = arcpy.GetCount_management(out_fc)
            if int(str(totalBlank)) > 0:
                outSheet.write(fields.index(field)+7, 4, int(str(totalBlank)))
            arcpy.management.Delete(out_fc)
    arcpy.Statistics_analysis(fc, output, stats_fields)
    fieldsOutput = arcpy.ListFields(output)
    for field in fieldsOutput:
        with SearchCursor(output, [field.name]) as cursor:
            for row in cursor:
                if fieldsOutput.index(field) > 1:
                    outSheet.write(fieldsOutput.index(field)+7, 3, int(totalRows[0]) - row[0])
    arcpy.management.Delete(output)

# function used to build report from tables
def tb_stats():
    outSheet = outWorkbook.add_worksheet(tb[0:30])
    outSheet.set_column(0, 4, 15)
    totalRows = arcpy.GetCount_management(tb)
    fields = arcpy.ListFields(tb)
    stats_fields = []
    output = "memory" + "\\" + str(tb)
    outSheet.write(0, 0, "NAME")
    outSheet.write(0, 1, tb)
    outSheet.write(1, 0, "TYPE")
    outSheet.write(1, 1, "Table")
    outSheet.write(2, 0, "GCS name")
    outSheet.write(2, 1, "N/A")
    outSheet.write(3, 0, "GCS type")
    outSheet.write(3, 1, "N/A")
    outSheet.write(4, 0, "ROWS")
    outSheet.write(4, 1, int(str(totalRows)))
    outSheet.write(5, 0, "FIELDS")
    outSheet.write(5, 1, int(str(len(fields))))
    outSheet.write(6, 0, "GEOM ERROR")
    outSheet.write(6, 1, "N/A")
    outSheet.write(8, 0, "FIELD")
    outSheet.write(8, 1, "ALIAS")
    outSheet.write(8, 2, "TYPE")
    outSheet.write(8, 3, "COUNT NULL")
    outSheet.write(8, 4, "COUNT BLANK")
    for field in fields:
        if field.type not in ("OID", "Geometry"):
            outSheet.write(fields.index(field)+8, 0, field.name)
            outSheet.write(fields.index(field)+8, 1, field.aliasName)
            outSheet.write(fields.index(field)+8, 2, field.type)
            stats_fields.append([field.name, "COUNT"])
        if field.type not in ("OID", "Geometry", "Double", "Integer", "SmallInteger", "Single"):
            out_tb = "memory" + "\\" + str(tb) + "_" + str(field.name)
            expression = str(field.name) + ' IN (\'\', \' \')'
            arcpy.TableSelect_analysis(tb, out_tb, expression)
            totalBlank = arcpy.GetCount_management(out_tb)
            if int(str(totalBlank)) > 0:
                outSheet.write(fields.index(field)+8, 4, int(str(totalBlank)))
            arcpy.management.Delete(out_tb)
    arcpy.Statistics_analysis(tb, output, stats_fields)
    fieldsOutput = arcpy.ListFields(output)
    for field in fieldsOutput:
        with SearchCursor(output, [field.name]) as cursor:
            for row in cursor:
                if fieldsOutput.index(field) > 1:
                    outSheet.write(fieldsOutput.index(field)+7, 3, int(totalRows[0]) - row[0])
    arcpy.management.Delete(output)

# list feature classes inside datasets and add sheets to report
fds = arcpy.ListDatasets()
for fd in fds:
    fcs = arcpy.ListFeatureClasses(feature_dataset=fd)
    for fc in fcs:
        fc_stats()

# list stand alone feature classes and add sheets to report
fcs = arcpy.ListFeatureClasses()
for fc in fcs:
    fc_stats()

# list tables and add sheets to report
tbs = arcpy.ListTables()
for tb in tbs:
    tb_stats()

# consolidate and open xlsx file
outWorkbook.close()
os.startfile(excelName)