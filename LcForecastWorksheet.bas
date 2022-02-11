Attribute VB_Name = "LcForecastWorksheet"
'@Folder("Generate PAF Workbook.WS Lc Forecast")
Option Explicit

'@Description "Generate subtotals, write tables to ws, add values to ws, add formulas to ws"
Sub Main( _
                        wbPaf As Workbook, _
                        objPl As clsPandL, _
                        dtReportingPeriod As Date, _
                        collActivities As Collection)
                        

Dim wsProjectList               As Worksheet
Dim arrVarActivityProjectList   As Variant      '2d array (0)activity name (1)array of project names
Dim arrVarPlTotalsByProject     As Variant      '2d array holding project amounts.  See below for schema
Dim arrVarPlTotalsByActivity    As Variant      '2d array holding activity amounts.  See below for schema
Dim arrVarPlTotals              As Variant      '2d array of months, (0)Rev Amount USD, (1)Cost Amount USD


'arrVarPlTotalsByProject array schema
'(i, 0) Activity Name
'(i, 1)(j, 0) Project Names Array
'(i, 1)(j, 1)(x, 0) x = month (1 to dtReportingPeriod), 0 = Rev Amount USD
'(i, 1)(j, 1)(x, 1) x = month (1 to dtReportingPeriod), 1 = Cost Amount
'example:
'   arrVarPlTotalsByProject(0, 0) = OMAN SLICKLINE
'   arrVarPlTotalsByProject(0, 1)(0, 0) = HCF Call Off
'   arrVarPlTotalsByProject(0, 1)(0, 1)(1, 0) = (Jan-21 Rev Amount USD)
'   arrVarPlTotalsByProject(0, 1)(0, 1)(1, 1) = (Jan-21 Costs Amount USD)
'Full path example:
'   arrVarPlTotalsByProject(0, 1)(0, 1)(1, 1) = Oman Slickline, HCF Call OFf, Jan-21 Cost Amount USD

'arrVarPlTotalsByActivity array schema
'(i, 0) Activity Name
'(i, 1)(x, 0) x = month (1 to dtReportingPeriod), 0 = Rev Amount USD
'(i, 1)(x, 1) x = month (1 to dtReportingPeriod), 1 = Cost Amount

'Assign worksheet to object
    Set wsProjectList = wbPaf.Worksheets("Project List")

'Get activites and projects from project list ws
    arrVarActivityProjectList = GetProjectAndActivitiesFromProjectListWs(wsProjectList)
    
'Populate arrVarPlTotalsByProject array
    arrVarPlTotalsByProject = GenerateProjectPlSubTotals(objPl, dtReportingPeriod, collActivities, arrVarActivityProjectList)
    

    
'Populate worksheet




End Sub

'@Description "To check arrVarPlTotalsByProject"
Private Sub DebugCheck_arrVarPlTotalsByProject(arrVarPlTotalsByProject As Variant)

Dim i As Long, j As Long, k As Long, m As Long
Dim strMonthRevCost As String
Dim strRevCost As String

For i = 0 To UBound(arrVarPlTotalsByProject, 1)
    For j = 0 To UBound(arrVarPlTotalsByProject(i, 1), 1)
        For k = 1 To UBound(arrVarPlTotalsByProject(i, 1)(j, 1), 1)
            For m = 0 To 1
                If m = 0 Then strRevCost = " Rev: " Else strRevCost = " Costs: "
                strMonthRevCost = " Month: " & MonthName(k, True) & strRevCost & arrVarPlTotalsByProject(i, 1)(j, 1)(k, m)
                Debug.Print _
                    "Activity: " & arrVarPlTotalsByProject(i, 0) & _
                    " Project: " & arrVarPlTotalsByProject(i, 1)(j, 0) & _
                    strMonthRevCost
            Next m
        Next k
    Next j
Next i

End Sub

'@Description "Gets the activities and projects from the projects list ws
Private Function GetProjectAndActivitiesFromProjectListWs( _
                            ByRef wsProjectList As Worksheet) _
                            As Variant

'With Project list ws
'for each named range check if it meets activity range pattern
'get the activity name from the range offset
'for each row after the header rows get the project name
'write the (0)activity and (1)project to the array

Dim i As Long, j As Long, k As Long
Dim arrVarActivityProjectList()     As Variant
Dim arrStrProjectList()             As String
Dim n                               As Name

'Get number of activities and redim activity list
    i = 0
    For Each n In wsProjectList.Names
        If GenericFunctions.StringSearch(1, n.Name, "Project.List_Activity.Name_") > 0 Then i = i + 1
    Next n

    ReDim arrVarActivityProjectList(0 To i - 1, 0 To 1) 'Size + 1 to add "Not Assigned"

'For each activity get number of projects, redim project array, get project names, store project
'array to activityproject array along with activity name
    i = 0 'counter for activities
    
    For Each n In wsProjectList.Names
        If GenericFunctions.StringSearch(1, n.Name, "Project.List_Activity.Name_") > 0 Then
            
            ReDim arrStrProjectList(0 To wsProjectList.Range(n).Rows.Count - 4)
            
            For j = 3 To wsProjectList.Range(n).Rows.Count - 1
                arrStrProjectList(j - 3) = wsProjectList.Range(n)(j, 2).Value
            Next j
            
            ReDim Preserve arrStrProjectList(UBound(arrStrProjectList, 1) + 1)
            arrStrProjectList(UBound(arrStrProjectList, 1)) = "Not Assigned"
            
            arrVarActivityProjectList(i, 0) = wsProjectList.Range(n)(1, 2).Value
            arrVarActivityProjectList(i, 1) = arrStrProjectList

            i = i + 1
        
        End If
    Next n
    
GetProjectAndActivitiesFromProjectListWs = arrVarActivityProjectList
        
End Function



