Attribute VB_Name = "LcForecastWorksheet2"
'@Folder("Generate PAF Workbook.WS Lc Forecast")
Option Explicit
Sub GenerateLcArray2( _
                        wbPaf As Workbook, _
                        objPl As clsPandL, _
                        dtReportingPeriod As Date, _
                        collActivities As Collection)

'From project list page get activities and projects from each named range
'store them as 2d array

Dim wsProjectList               As Worksheet
Dim arrVarActivityProjectList   As Variant      '2d array (0)activity name (1)array of project names
Dim arrVarPlMontlyTotals        As Variant      '2d array (0)Revenue - AmountUSD (1)Cost - Amount USD
Dim arrVarActivityMonthlyTotals As Variant      '2d array (0)Activity (1)Revenue - AmountUSD (2)Cost - AmountUSD


Set wsProjectList = wbPaf.Worksheets("Project List")

'Get activites and projects from project list ws
    arrVarActivityProjectList = GetProjectAndActivitiesFromProjectListWs(wsProjectList)
    
'Calculate P&L monthly totals up month prior to reporting monht
    arrVarPlMontlyTotals = CalcMonthlyPlTotals(objPl, arrVarActivityProjectList, dtReportingPeriod, collActivities)
    
'Per activity calculate monthly totals up to month priort to reporting month
    arrVarActivityMonthlyTotals = CalcMonthlyActivityTotals(objPl, arrVarActivityProjectList, dtReportingPeriod, collActivities)



End Sub

'@Description "Returns an array with the rev/cost totals for each month by activity"
Private Function CalcMonthlyActivityTotals( _
                                                ByRef objPl As clsPandL, _
                                                ByRef arrVarActivityProjectList As Variant, _
                                                ByVal dtReportingPeriod As Date, _
                                                ByRef collActivities As Collection) _
                                                As Variant
                                                
'Redim array for reporting month - 1
'for each activity pull the P&L finance table/header
'for each month ->
'for each project sum the rev values and store to array as activity, rev/cost amount usd

Dim i As Long, j As Long, k As Long, m As Long
Dim arrVarActivityMonthlyTotals     As Variant  'Return array of monthly amounts for all activities.  see below for structure
Dim arrVarActivitiesForMonth        As Variant  'Array of all activities for the month
Dim arrVarActivtyRevCost            As Variant  'Array of (0)Rev Amount USD and (1)Costs Amount USD for an activity
Dim dblAmountUSD                    As Double   'holds the total amount for a loop
Dim dtTargetMonth                   As Date
Dim strRevCost                      As String

'arrVarActivityMonthlyTotals nested array structure
'(0) Array of activity names for month
'(0)(0,0) Activity Name
'(0)(0,1)(0) Revenue Amount USD
'(0)(0,1)(1) Costs Amount USD

'redim activity rev/cost array number of months * number of projects
    ReDim arrVarActivityMonthlyTotals(0 To (Month(dtReportingPeriod) - 2))
    
'calc loop
    For i = 0 To UBound(arrVarActivityMonthlyTotals, 1) 'Month accumulator
        
        'Redim activities for month list
            ReDim arrVarActivitiesForMonth(0 To UBound(arrVarActivityProjectList, 1), 0 To 1)
        
        'convert the month loop number (i) into a full date
            dtTargetMonth = CDate(CStr(Year(dtReportingPeriod)) & "-" & MonthName(i + 1, True) & "-" & "01")
        
        '----Activities Loop
        For j = 0 To UBound(arrVarActivitiesForMonth, 1)
            
            'set the activity name
                arrVarActivitiesForMonth(j, 0) = arrVarActivityProjectList(j, 0)
            
            'reset the rev/cost array
                ReDim arrVarActivtyRevCost(0 To 1)
            
            '----Rev/Cost Loop
            For k = 0 To 1
                
                dblAmountUSD = 0
                
                If k = 0 Then
                    strRevCost = "Revenue"
                Else
                    strRevCost = "Costs"
                End If
                
                '----Project Loop
                For m = 0 To UBound(arrVarActivityProjectList(j, 1), 1)
                    'if "no projects" goto next project, otherwise call function to calc either rev or cost subtotal
                    If Not arrVarActivityProjectList(j, 1)(m) = "No Projects" Then
                        dblAmountUSD = GetFinanceTableRevCostSubTotal(objPl, collActivities(arrVarActivitiesForMonth(j, 0)), dtTargetMonth, arrVarActivityProjectList(j, 1)(m), strRevCost, dblAmountUSD)
                    End If
                Next m 'Next project
                
                'Project "Not Assigned" sub total
                    dblAmountUSD = GetFinanceTableRevCostSubTotal(objPl, collActivities(arrVarActivitiesForMonth(j, 0)), dtTargetMonth, "Not Assigned", strRevCost, dblAmountUSD)
                
                'store sub total to rev/cost array
                    arrVarActivtyRevCost(k) = dblAmountUSD
            
            Next k 'Next Rev/Cost
            
            'assign rev/cost array to month activity array
                arrVarActivitiesForMonth(j, 1) = arrVarActivtyRevCost
            
        Next j 'next activity
        
        'assign month activity array to totals array
            arrVarActivityMonthlyTotals(i) = arrVarActivitiesForMonth
    
    Next i 'next month

CalcMonthlyActivityTotals = arrVarActivityMonthlyTotals

End Function

'@Description "Returns an array with the Rev/Cost totals for each month"
Private Function CalcMonthlyPlTotals( _
                                        ByRef objPl As clsPandL, _
                                        ByRef arrVarActivityProjectList As Variant, _
                                        ByVal dtReportingPeriod As Date, _
                                        ByRef collActivities As Collection) _
                                        As Variant

'Redim array for reporting month - 1
'for each activity pull the P&L finance table/header
'for each month ->
'for each project sum the rev values and store to array, repeat for costs
'return total

Dim i As Long, j As Long, k As Long, m As Long
Dim arrVarPlMonthlyTotals   As Variant  '(0)Revenue Amount USD, (1)Costs Amount USD
Dim dblAmountUSD            As Double   'holds the total amount for a loop
Dim dtTargetMonth           As Date
Dim strRevCost              As String   'Holds Rev or Cost for sub total filtering
Dim intRevCostArrAssign     As Integer  'holds 1(Rev) or 2(Cost) for array assignment

ReDim arrVarPlMonthlyTotals(0 To Month(dtReportingPeriod) - 2, 0 To 1)


'Revenue Amount
For i = 0 To UBound(arrVarPlMonthlyTotals, 1)

    dtTargetMonth = CDate(CStr(Year(dtReportingPeriod)) & "-" & MonthName(i + 1, True) & "-" & "01")
    
    For m = 0 To 1
        
        If m = 0 Then
            strRevCost = "Revenue"
            intRevCostArrAssign = 0
        Else
            strRevCost = "Costs"
            intRevCostArrAssign = 1
        End If
        
        dblAmountUSD = 0
        
        'j = activity, k = projects
        For j = 0 To UBound(arrVarActivityProjectList, 1)
            
            For k = 0 To UBound(arrVarActivityProjectList(j, 1), 1)
                
                'Get amount for project
                If Not arrVarActivityProjectList(j, 1)(k) = "no projects" Then
                    dblAmountUSD = GetFinanceTableRevCostSubTotal(objPl, collActivities(arrVarActivityProjectList(j, 0)), dtTargetMonth, arrVarActivityProjectList(j, 1)(k), strRevCost, dblAmountUSD)
                End If
                
            Next k 'Next Project
            
            'Get amount for "Not Assigned"
            dblAmountUSD = GetFinanceTableRevCostSubTotal(objPl, collActivities(arrVarActivityProjectList(j, 0)), dtTargetMonth, "Not Assigned", strRevCost, dblAmountUSD)
        
        Next j 'Next activity
        
        arrVarPlMonthlyTotals(i, intRevCostArrAssign) = dblAmountUSD
    
    Next m 'next rev/cost
        
Next i 'next month

CalcMonthlyPlTotals = arrVarPlMonthlyTotals

End Function

'@Description "Sums the Rev or Cost sub total for a given project/activity/pl/month
Private Function GetFinanceTableRevCostSubTotal( _
                                                ByRef objPl As clsPandL, _
                                                ByRef objActivity As clsActivity, _
                                                ByVal dtTargetDate As Date, _
                                                ByVal strProjectName As String, _
                                                ByVal strRevCost As String, _
                                                Optional ByVal dblIncomingAmountUSD As Double = 0) _
                                                As Double

'With activity
'get p&l finance table/header arrays
'get month column index from header array
'for target month were rows match project, rev/cost
'Header cols (0)Project Name (1)Rev/Cost (2)Desc Group (3)Desc (4)...Month/AmountUSD

Dim i As Long, j As Long, k As Long
Dim arrVarFinanceHeader         As Variant
Dim arrVarFinanceTable          As Variant
Dim intTargetMonthColumnIndex   As Integer
Dim dblAmountUSD                As Double

'Assign dictionary arrays to variables
    arrVarFinanceHeader = objActivity.dictFinanceDataTableHeader(objPl.strName)
    arrVarFinanceTable = objActivity.dictFinanceDataTable(objPl.strName)

'Get the column number of the target month
    For i = 0 To UBound(arrVarFinanceHeader, 1)
        If arrVarFinanceHeader(i) = Format(dtTargetDate, "MMM-YYYY") Then intTargetMonthColumnIndex = i
    Next i

'Accumlate total based on criteria
    dblAmountUSD = 0
    
    For i = 0 To UBound(arrVarFinanceTable, 1)

        If arrVarFinanceTable(i, 0) = strProjectName And arrVarFinanceTable(i, 1) = strRevCost And _
            Not IsNull(arrVarFinanceTable(i, intTargetMonthColumnIndex)) Then
                    dblAmountUSD = dblAmountUSD + arrVarFinanceTable(i, intTargetMonthColumnIndex)
        End If
        
    Next i
       
GetFinanceTableRevCostSubTotal = dblAmountUSD + dblIncomingAmountUSD

End Function

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

    ReDim arrVarActivityProjectList(0 To i - 1, 0 To 1)

'For each activity get number of projects, redim project array, get project names, store project
'array to activityproject array along with activity name
    i = 0 'counter for activities
    
    For Each n In wsProjectList.Names
        If GenericFunctions.StringSearch(1, n.Name, "Project.List_Activity.Name_") > 0 Then
            
            ReDim arrStrProjectList(0 To wsProjectList.Range(n).Rows.Count - 4)
            
            For j = 3 To wsProjectList.Range(n).Rows.Count - 1
                arrStrProjectList(j - 3) = wsProjectList.Range(n)(j, 2).Value
            Next j
            
            arrVarActivityProjectList(i, 0) = wsProjectList.Range(n)(1, 2).Value
            arrVarActivityProjectList(i, 1) = arrStrProjectList
            i = i + 1
        
        End If
    Next n
    
GetProjectAndActivitiesFromProjectListWs = arrVarActivityProjectList
        
End Function
