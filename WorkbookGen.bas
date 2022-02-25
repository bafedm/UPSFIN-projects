Attribute VB_Name = "WorkbookGen"
'@Folder("Generate PAF Workbook")
'@Description("Hold main loop for generating the PAF workbooks for each P&L")
Option Explicit

Sub Main( _
            ByRef collPls As Collection, _
            ByRef collActivities As Collection, _
            ByRef collProjects As Collection, _
            ByVal dtReportingPeriod As Date)

Dim wbPaf               As Workbook     'A PAF workbook to be added to the wb collection
Dim collPafWorkbooks    As Collection   'Collection containing all PAF workbooks
Dim objPl               As clsPandL     'P&L object from collPls
Dim strSaveName         As String       'File name for a PAF workbook
Dim strSavePath         As String       'Path to store the PAF workbooks


Set collPafWorkbooks = New Collection
'home path
    strSavePath = "C:\Users\Blake\OneDrive\98 Misc\Special Assignments\UPSFIN\Project UPSFIN\PAF Workbooks\"
'work path
    'strSavePath = "C:\Users\blake_fudge\OneDrive - SGS\Ops\Misc\Special Assignments\UPSFIN\Project UPSFIN\PAF Workbooks\"

'Main loop for each P&L in collection
    For Each objPl In collPls
    
        Set wbPaf = OpenNewPafWorkbookTemplate(collPafWorkbooks, objPl, dtReportingPeriod, strSavePath)
        
        'write to projects worksheet
            ProjectsWorksheet.PopulateProjectsWorksheet wbPaf, dtReportingPeriod, objPl, collActivities, collProjects
            
        'write to allocations worksheet
            AllocationsWorksheet.PopulateAllocationsWorksheet wbPaf, dtReportingPeriod, objPl, collActivities
            
        'write to lc forecast worksheet
            LcForecastWorksheet.Main wbPaf, objPl, dtReportingPeriod, collActivities
            
        
    
    Next objPl

'close workbooks
    CloseAllPafWorkbooks collPafWorkbooks

End Sub

'@Description "Opens a new PAF Template and saves it based on the P&L name, returns a workbook obj"
Private Function OpenNewPafWorkbookTemplate( _
                                        ByRef collPafWorkbooks As Collection, _
                                        ByRef objPl As clsPandL, _
                                        ByVal dtReportingPeriod As Date, _
                                        ByVal strSavePath As String) As Workbook
                                        
Dim wbPaf           As Workbook 'Workbook returned to caller
Dim strFileName     As String   'generated name based on month and p&l
Dim strSaveName     As String   'full path and file name of new PAF
                                        
'Disable save alerts
    Application.DisplayAlerts = False

'Open template workbook
    Set wbPaf = Workbooks.Open(strSavePath & "Project Allocation and Forecast Template.xlsm")

'Save template as new Paf based on P&L name and Reporting month
    strFileName = "PAF " & objPl.strName & " " & Format(dtReportingPeriod, "MMMYYYY")
    strSaveName = strSavePath & strFileName
    wbPaf.SaveAs Filename:=strSaveName

'Add workbook to collection
    collPafWorkbooks.Add Key:=objPl.strName, Item:=wbPaf
    
'Resize and move window
    ResizeWindowForTesting strFileName

'Turn alerts back on
    Application.DisplayAlerts = True

'Return workbook to caller
    Set OpenNewPafWorkbookTemplate = wbPaf

End Function

'@Description "Resizes window upon completion to fit nicely on screen"
Private Sub ResizeWindowForTesting( _
                                        ByVal strFileName As String)

With Application.Windows(strFileName & ".xlsm")
    .WindowState = xlNormal
    .Top = 0
    .Left = 1280
    .Height = 800
    .Width = 640
End With

End Sub


'@Description "Closes all workbooks in the collection"
Private Sub CloseAllPafWorkbooks( _
                                    ByRef collPafWorkbooks As Collection)

Dim wbPaf   As Workbook

For Each wbPaf In collPafWorkbooks
    wbPaf.Close savechanges:=True
Next wbPaf

End Sub

