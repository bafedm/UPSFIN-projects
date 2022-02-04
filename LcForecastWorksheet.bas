Attribute VB_Name = "LcForecastWorksheet"
'@Folder("Generate PAF Workbook.WS Lc Forecast")

Sub PopulateLcForecastWorksheet( _
                                ByRef wbPaf As Workbook, _
                                ByVal dtReportingPeriod As Date, _
                                ByRef objPl As clsPandL, _
                                ByRef collActivities As Collection)

'create array with dm lc data
'get activities from dm
'get activity projects from projects list ws
'for each activity
'get p&l finance/finance header objects
'for each month sum each project rev/cost and store to array
'next activity


Dim wsProjectList               As Worksheet                    'Assign Project List ws to variable
Dim wsAllocations               As Worksheet                    'Assign Allocations ws to variable
Dim wsLcForecast                As Worksheet
Dim arrVarLcData                As Variant

Set wsProjectList = wbPaf.Worksheets("Project List")
Set wsAllocations = wbPaf.Worksheets("Allocations")
Set wsLcForecast = wbPaf.Worksheets("LC Forecast")
    
arrVarLcData = GenerateLcArray(wsProjectList, objPl, dtReportingPeriod, collActivities)

'arrVarLcData structure
'0: Activity Name
'1: Project Name
'2: Revenue/Cost
'3...: Monthly Rev or cost total AmountUSD



'write p&l forecast actual
'for each month
'sum monthly values from array for all projects from array
'write for month data
'next month
'set range name

'write activity forecast actual
'for each activity
'for each month
'sum monthly values for activity projects from array
'write month data
'next month
'set name range
'if no projects set grouping

'write project forecast actual
'for each project
'for each month
'get values from array for project
'write to month
'next month
'set name range
'if last project set grouping for activity

'next activity

'write formulas
'write lc formulas
'for each lc named range
'for each month
'write LC, LC%
'next month
'next lc named range

'write p&l totals
'for each month
'for each activity
'get month rev/cost cell locations, concate into sum formula
'next activity
'write formula to p&l total rev/cost cell for month
'next month

'write activity
'for each month
'for each activity project
'get month rev/cost cell locations, concate into sum formula
'next project
'write formula to rev/cost cell for month
'next month




End Sub

'@Description "Builds and array that holds the rev/cost totals for each project"
Private Function GenerateLcArray( _
                                    ByRef wsProjects As Worksheet, _
                                    ByRef objPl As clsPandL, _
                                    ByVal dtReportingPeriod As Date, _
                                    ByRef collActivities As Collection)

'create array with dm lc data
'get activities from dm
'get activity projects from projects list ws
'for each activity
'get p&l finance/finance header objects
'for each month sum each project rev/cost and store to array
'next activity

Dim i As Long, j As Long, k As Long, m As Long
Dim arrVarLcData            As Variant      'Holds LC values for each activity/project/month
Dim objActivity             As New clsActivity  'an individual activity object
Dim objProject              As clsProject   'an individual project object
Dim nmeProjectListActivity  As Name         'a worksheet named range
Dim rngActivity             As Range



'arrVarLcData structure
'0: Activity Name
'1: Project Name
'2: Revenue/Cost
'3...: Monthly Rev or cost total AmountUSD

'redim array with number of projects * 2 and number of months + 3
    i = 0
    For Each nmeProjectListActivity In wsProjects.Names
        If GenericFunctions.StringSearch(1, nmeProjectListActivity.Name, "Project.List_Activity.Name_") > 0 Then
            i = i + 1 + (wsProjects.Range(nmeProjectListActivity).Rows.Count - 3)
        End If
    Next nmeProjectListActivity
            
    ReDim arrVarLcData(0 To (i * 2) - 1, 0 To 3 + (Month(dtReportingPeriod) - 1))
       
'loop to fill array.
'for each activity that has p&l
'get finance table/headers
'generate unique list of projects from finance table
'for each project
'for each monht sum all rev lines, then all costs lines
'write activity, project, cost/rev, amount to array
'next month
'next project

'get finance table and header from activity/p&l
Dim arrVarFinanceTable As Variant
Dim arrVarFinanceHeader As Variant
Dim collFinTableProjects As Collection
Dim intFinTableTargetMonthIndex As Integer
Dim dblAmountUSD As Double
Dim strRevCost As String
Dim varFinTableProject As Variant


'reset lc array row counter
    i = 0

For Each nmeProjectListActivity In wsProjects.Names
    If GenericFunctions.StringSearch(1, nmeProjectListActivity.Name, "Project.List_Activity.Name_") > 0 Then
        
        'set activity object based on project list activity name
            Set objActivity = collActivities(wsProjects.Range(nmeProjectListActivity)(1, 2).Value)
        
        'get finance data from activity
            With objActivity
                If Not GenericFunctions.HasKey(.collParentPl, objPl.strName) Then
                    arrVarFinanceTable = .dictFinanceDataTable(objPl.strName)
                    arrVarFinanceHeader = .dictFinanceDataTableHeader(objPl.strName)
                    'Debug.Print .strName, .dictFinanceDataTable(objPl.strName)(1, 1), .dictFinanceDataTableHeader(objPl.strName)(1)
                End If
            End With
        
        'reset project collection
            Set collFinTableProjects = New Collection
        
        'get unique list of projects from finance table
            For j = 0 To UBound(arrVarFinanceTable, 1)
                If Not GenericFunctions.HasKey(collFinTableProjects, CStr(arrVarFinanceTable(j, 0))) Then
                    collFinTableProjects.Add Key:=arrVarFinanceTable(j, 0), Item:=arrVarFinanceTable(j, 0)
                End If
            Next j
        
        'get index of reporting month from fin table header
            For j = 0 To UBound(arrVarFinanceHeader, 1)
                If arrVarFinanceHeader(j) = Format(dtReportingPeriod, "MMM-YYYY") Then intFinTableTargetMonthIndex = j
            Next j
            
        'loop months and sum values for rev/cost values for each project
         For Each varFinTableProject In collFinTableProjects
                For k = 0 To 1
                    For j = 4 To intFinTableTargetMonthIndex
                        If k = 1 Then strRevCost = "Revenue" Else strRevCost = "Costs"
                            dblAmountUSD = 0
                            For m = 0 To UBound(arrVarFinanceTable, 1)
                                If arrVarFinanceTable(m, 0) = varFinTableProject And arrVarFinanceTable(m, 1) = strRevCost Then
                                    If Not IsNull(arrVarFinanceTable(m, j)) Then dblAmountUSD = dblAmountUSD + arrVarFinanceTable(m, j)
                                End If
                            Next m
                            
                            Debug.Print i, objActivity.strName, varFinTableProject, strRevCost, dblAmountUSD
                            
'                            arrVarLcData(i, 0) = objActivity.strName
'                            arrVarLcData(i, 1) = varFinTableProject
'                            arrVarLcData(i, 2) = strRevCost
'                            arrVarLcData(i, j - 1) = dblAmountUSD
                    Next j
                    i = i + 1
                Next k
 

        Next varFinTableProject
                            
        
    End If
Next nmeProjectListActivity
        

End Function
