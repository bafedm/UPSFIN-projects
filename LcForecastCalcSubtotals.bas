Attribute VB_Name = "LcForecastCalcSubtotals"
'@Folder("Generate PAF Workbook.WS Lc Forecast")
Option Explicit

'@Description "Calculates the rev/cost subtotals on P&L basis"
Function GeneratePlSubTotals( _
                            ByRef arrVarPlTotalsByActivity As Variant, _
                            ByVal dtReportingPeriod As Date) _
                            As Variant
                            
'loop activities
'loop months
'loop rev/costs
'sum rev/costs
'next rev/cost
'next month
'next activity

Dim i As Long, j As Long, k As Long
Dim arrVarPlTotals      As Variant

ReDim arrVarPlTotals(1 To Month(dtReportingPeriod) - 1, 0 To 1)

'Activity Loop
For i = 0 To UBound(arrVarPlTotalsByActivity, 1)
    
    'Month Loop
    For j = 1 To Month(dtReportingPeriod) - 1
    
        'Rev/Cost Loop
        For k = 0 To 1
        
            arrVarPlTotals(j, k) = arrVarPlTotals(j, k) + arrVarPlTotalsByActivity(i, 1)(j, k)
            
        Next k 'Next Rev/Cost
            
    Next j 'Next Month
    
Next i 'Next Activity

GeneratePlSubTotals = arrVarPlTotals

End Function

'@Description "Calculates the rev/cost subtotals on activity basis"
Function GenerateActivityPlSubTotals( _
                                    ByRef arrVarPlTotalsByProject As Variant) _
                                    As Variant

'loop activity
'   loop project
'       loop month
'           loop rev/cost
'               sum rev/cost for all months
'           next rev/cost
'       next month
'   next project
'next activity

'arrVarPlTotalsByActivity array schema
'(i, 0) Activity Name
'(i, 1)(x, 0) x = month (1 to dtReportingPeriod), 0 = Rev Amount USD
'(i, 1)(x, 1) x = month (1 to dtReportingPeriod), 1 = Cost Amount

Dim i As Long, j As Long, k As Long, m As Long
Dim arrVarPlTotalsByActivity    As Variant
Dim arrVarPlTotalsByMonth       As Variant

'loop
'Activity Loop
ReDim arrVarPlTotalsByActivity(UBound(arrVarPlTotalsByProject, 1), 0 To 1)

For i = 0 To UBound(arrVarPlTotalsByProject, 1)
    arrVarPlTotalsByActivity(i, 0) = arrVarPlTotalsByProject(i, 0)
        
        'Project Loop
        For j = 0 To UBound(arrVarPlTotalsByProject(i, 1), 1)
        
            'Month Loop
            ReDim arrVarPlTotalsByMonth(1 To UBound(arrVarPlTotalsByProject(i, 1)(j, 1), 1), 0 To 1)
            
            For k = 1 To UBound(arrVarPlTotalsByProject(i, 1)(j, 1), 1)
            
                'Rev/Cost Loop
                For m = 0 To 1
                    
                    arrVarPlTotalsByMonth(k, m) = arrVarPlTotalsByMonth(k, m) + arrVarPlTotalsByProject(i, 1)(j, 1)(k, m)
                
                Next m 'Next Rev/Cost
            
            Next k 'Next Month
            
        Next j 'Next Project
        
        arrVarPlTotalsByActivity(i, 1) = arrVarPlTotalsByMonth
        
Next i 'Activity Loop

GenerateActivityPlSubTotals = arrVarPlTotalsByActivity

End Function

'@Description "iterate through activities and projects to generate PL sub totals"
Function GenerateProjectPlSubTotals( _
                                                ByRef objPl As clsPandL, _
                                                ByVal dtReportingPeriod As Date, _
                                                ByRef collActivities As Collection, _
                                                ByVal arrVarActivityProjectList As Variant) _
                                                As Variant

                                                
'loop activities (max = ubound [activity]/project list)
'   loop projects (max = ubound activity/[project] list)
'       loop months (max = dtReportingPeriod - 1) NOTE: One less then current month as current month coming from allocations sheet
'           loop rev/cost (0 to 1)
'               call GetFinanceTableRevCostSubTotal
'               save to array
'           next rev/cost
'       next month
'   next project
'next activity

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

'GetFinanceTableRevCostSubTotal arguments
'   objPl As clsPandL                       - Pass through
'   objActivity As clsActivity              - Get based on loop, from collActivities
'   dtTargetDate As Date                    - Get based on month loop
'   strProjectName As String                - Get based on project loop
'   strRevCost As String                    - Get based on rev/cost loop
'   dblIncomingAmountUSD As Double = 0      - Optional, not required

Dim i As Long, j As Long, k As Long, m As Long
Dim objActivity                     As clsActivity
Dim dtTargetMonth                   As Date
Dim arrVarPlTotalsByProject         As Variant
Dim arrVarProjectList               As Variant
Dim arrVarPlMonthlyRevCostTotals    As Variant
Dim strRevCost                      As String

'Resize array to number of activities
    ReDim arrVarPlTotalsByProject(0 To UBound(arrVarActivityProjectList, 1), 0 To 1)

'Loop
    'Activity Loop
    For i = 0 To UBound(arrVarActivityProjectList, 1) 'Activity loop
        
        'Assign Activity Name
            arrVarPlTotalsByProject(i, 0) = arrVarActivityProjectList(i, 0)
        
        'Resize/reset project list array
            ReDim arrVarProjectList(0 To UBound(arrVarActivityProjectList(i, 1), 1), 0 To 1)
        
        'Project Loop
        For j = 0 To UBound(arrVarActivityProjectList(i, 1), 1) 'Project loop
            
            'Assign Project Name
                If arrVarActivityProjectList(i, 1)(j) = "Not Assigned" Then
                    arrVarProjectList(j, 0) = arrVarActivityProjectList(i, 0) & " " & arrVarActivityProjectList(i, 1)(j)
                Else
                    arrVarProjectList(j, 0) = arrVarActivityProjectList(i, 1)(j)
                End If
            
            'Resize/reset Month/Amount USD array
                ReDim arrVarPlMonthlyRevCostTotals(1 To Month(dtReportingPeriod) - 1, 0 To 1)
            
            'Loop Month
            For k = 1 To Month(dtReportingPeriod) - 1 'month loop
                
                'Generate target month
                    dtTargetMonth = CDate(CStr(Year(dtReportingPeriod)) & "-" & MonthName(k, True) & "-" & "01")
                
                'Rev/Cost Loop
                For m = 0 To 1 'rev/cost loop
                
                'Set Rev/Cost variable
                    If m = 0 Then strRevCost = "Revenue" Else strRevCost = "Costs"
                    
                'Get subtotal
                    arrVarPlMonthlyRevCostTotals(k, m) = GetFinanceTableRevCostSubTotal( _
                                                        objPl, _
                                                        collActivities(arrVarActivityProjectList(i, 0)), _
                                                        dtTargetMonth, _
                                                        arrVarActivityProjectList(i, 1)(j), _
                                                        strRevCost)
            
                Next m 'Next Rev/Cost
                
            Next k ' Next Month
            
            'Assign monthly rev/cost amount to project
                arrVarProjectList(j, 1) = arrVarPlMonthlyRevCostTotals
        
        Next j 'Next project
        
        'Assign project to activity
            arrVarPlTotalsByProject(i, 1) = arrVarProjectList
        
    Next i 'Next Activity

GenerateProjectPlSubTotals = arrVarPlTotalsByProject

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


