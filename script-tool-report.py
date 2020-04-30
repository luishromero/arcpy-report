from arcpy.da import SearchCursor
import arcpy
import xlsxwriter
import os

gdbPath = arcpy.GetParameterAsText(0)
excelName = arcpy.GetParameterAsText(1)

# delete temp fgdb if exist
if arcpy.Exists("tempStats.gdb"):
    arcpy.management.Delete("tempStats.gdb")
# set env
arcpy.env.workspace = gdbPath
# create temp fgdb
arcpy.CreateFileGDB_management(os.getcwd(), "tempStats.gdb")
path = "tempStats.gdb"
# create xlsx
outWorkbook = xlsxwriter.Workbook(excelName)

# function used to build reports from feature classes
def fc_stats():
    outSheet = outWorkbook.add_worksheet(fc[0:30])
    totalRows = arcpy.GetCount_management(fc)
    spatialRef = arcpy.Describe(fc).spatialReference
    fields = arcpy.ListFields(fc)
    stats_fields = []
    output = path + "\\" + str(fc)
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
    outSheet.write(7, 0, "FIELD")
    outSheet.write(7, 1, "ALIAS")
    outSheet.write(7, 2, "TYPE")
    outSheet.write(7, 3, "COUNT NULL")
    outSheet.write(7, 4, "COUNT BLANK")
    for field in fields:
        if field.type not in ("OID", "Geometry"):
            outSheet.write(fields.index(field)+6, 0, field.name)
            outSheet.write(fields.index(field)+6, 1, field.aliasName)
            outSheet.write(fields.index(field)+6, 2, field.type)
            stats_fields.append([field.name, "COUNT"])
        if field.type not in ("OID", "Geometry", "Double", "Integer", "SmallInteger", "Single"):
            out_fc = path + "\\" + str(fc) + "_" + str(field.name)
            expression = str(field.name) + ' IN (\'\', \' \')'
            arcpy.Select_analysis(fc, out_fc, expression)
            totalBlank = arcpy.GetCount_management(out_fc)
            if int(str(totalBlank)) > 0:
                outSheet.write(fields.index(field)+6, 4, int(str(totalBlank)))
    arcpy.Statistics_analysis(fc, output, stats_fields)
    fieldsOutput = arcpy.ListFields(output)
    for field in fieldsOutput:
        with SearchCursor(output, [field.name]) as cursor:
            for row in cursor:
                if fieldsOutput.index(field) > 1:
                    outSheet.write(fieldsOutput.index(field)+6, 3, int(totalRows[0]) - row[0])

# function used to build report from tables
def tb_stats():
    outSheet = outWorkbook.add_worksheet(tb[0:30])
    totalRows = arcpy.GetCount_management(tb)
    fields = arcpy.ListFields(tb)
    stats_fields = []
    output = path + "\\" + str(tb)
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
    outSheet.write(7, 0, "FIELD")
    outSheet.write(7, 1, "ALIAS")
    outSheet.write(7, 2, "TYPE")
    outSheet.write(7, 3, "COUNT NULL")
    outSheet.write(7, 4, "COUNT BLANK")
    for field in fields:
        if field.type not in ("OID", "Geometry"):
            outSheet.write(fields.index(field)+7, 0, field.name)
            outSheet.write(fields.index(field)+7, 1, field.aliasName)
            outSheet.write(fields.index(field)+7, 2, field.type)
            stats_fields.append([field.name, "COUNT"])
        if field.type not in ("OID", "Geometry", "Double", "Integer", "SmallInteger", "Single"):
            out_tb = path + "\\" + str(tb) + "_" + str(field.name)
            expression = str(field.name) + ' IN (\'\', \' \')'
            arcpy.TableSelect_analysis(tb, out_tb, expression)
            totalBlank = arcpy.GetCount_management(out_tb)
            if int(str(totalBlank)) > 0:
                outSheet.write(fields.index(field)+7, 4, int(str(totalBlank)))
    arcpy.Statistics_analysis(tb, output, stats_fields)
    fieldsOutput = arcpy.ListFields(output)
    for field in fieldsOutput:
        with SearchCursor(output, [field.name]) as cursor:
            for row in cursor:
                if fieldsOutput.index(field) > 1:
                    outSheet.write(fieldsOutput.index(field)+6, 3, int(totalRows[0]) - row[0])

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
