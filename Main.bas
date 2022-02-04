Attribute VB_Name = "Main"
Option Explicit

'@Folder("Main")
'@Description "Main Loop"

Sub Main()

'UPSFIN Workbook Objects
    Dim wbUpsfin            As Workbook:    Set wbUpsfin = ThisWorkbook
    Dim wsProjectWb         As Worksheet:   Set wsProjectWb = wbUpsfin.Worksheets("Project WB Generator")

'Global Variables from UPSFIN Workbook
    Dim dtReportingPeriod   As Date

'Global Object Collections
    Dim collPls             As Collection
    Dim collActivies        As Collection
    Dim collProjects        As Collection

'test data
    dtReportingPeriod = "1-Apr-2021"

'Load object data from data model
    DataLoad.Main wbUpsfin, wsProjectWb, dtReportingPeriod, collPls, collActivies, collProjects

'Generate workbook for each P&L
    'For now we only want one P&L to test so create new collection and assign Oman P&L
    Dim collTempPl As New Collection
    collTempPl.Add Key:=collPls("OMAN").strName, Item:=collPls("OMAN")
    'collTempPl.Add key:=collPls("ADCO-CHS").strName, item:=collPls("ADCO-CHS")
    
    WorkbookGen.Main collTempPl, collActivies, collProjects, dtReportingPeriod
    
End Sub
