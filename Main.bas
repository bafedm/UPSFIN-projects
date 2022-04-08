Attribute VB_Name = "Main"
Option Explicit
'@Folder("Main")
'@Description "Main Loop"

Sub Main()

    Dim i As Long

'UPSFIN Workbook Objects
    Dim wbUpsfin            As Workbook:    Set wbUpsfin = ThisWorkbook
    Dim wsProjectWb         As Worksheet:   Set wsProjectWb = wbUpsfin.Worksheets(WS_PAF_GEN)

'Global Variables from UPSFIN Workbook
    Dim dtAccountingPeriod   As Date

'Global Object Collections
    Dim collPls             As Collection
    Dim collActivies        As Collection
    Dim collProjects        As Collection

'test data
    'test for one month
        dtAccountingPeriod = "1-Apr-2021"
    
    'testing for each month - need to comment in/out the "next i" at bottom
'        For i = 1 To 4
'            dtReportingPeriod = CDate(CStr(2021) & "-" & MonthName(i, True) & "-" & "01")

        'Load object data from data model
            DataLoad.Main wbUpsfin, wsProjectWb, dtAccountingPeriod, collPls, collActivies, collProjects
        
            'For now we only want one P&L to test so create new collection and assign Oman P&L
                Dim collTempPl As Collection
                Set collTempPl = New Collection
                collTempPl.Add Key:=collPls("OMAN").strName, Item:=collPls("OMAN")
                'collTempPl.Add key:=collPls("ADCO-CHS").strName, item:=collPls("ADCO-CHS")
        
        'Generate workbook for each P&L
            WorkbookGen.Main collTempPl, collActivies, collProjects, dtAccountingPeriod
    
    'testing for each month'
'        Next i
End Sub
