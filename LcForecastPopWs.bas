Attribute VB_Name = "LcForecastPopWs"
'@Folder("Generate PAF Workbook.WS Lc Forecast")
Option Explicit

'@Description "Methods for writing to LC worksheet"
Sub WriteToLcWorksheet( _
                        ByRef objPl As clsPandL, _
                        ByRef wbPaf As Workbook, _
                        ByVal arrVarPlTotalsByProject As Variant, _
                        ByVal dtReportingPeriod As Date)

'Write tables, assign name range, group
'populate sub total amounts
'write formulas to calculate allocations amount for projects
'write formulas to calculate activity/pl amounts from projects/activities
'format ranges

Dim wsLcForecast As Worksheet
Dim wsAllocations As Worksheet

'Set worksheet to object
    Set wsLcForecast = wbPaf.Worksheets("LC Forecast")
    Set wsAllocations = wbPaf.Worksheets("Allocations")
    
'Write activity/project blank tables and set range names
    WriteBlankTablesToLcWorksheet objPl, wsLcForecast, arrVarPlTotalsByProject, dtReportingPeriod
    
'Write project revenue/costs values to tables
    WriteRevCostValuesToTables wsLcForecast, wsAllocations, arrVarPlTotalsByProject, dtReportingPeriod
    
'Write formulas to worksheet
    WriteFormulasToTables wsLcForecast, dtReportingPeriod, arrVarPlTotalsByProject

End Sub

'@Description "Writes all formulas to tables - activities/projects/P&L sums, LC&LC% calc
Private Sub WriteFormulasToTables( _
                                    ByRef wsLcForecast As Worksheet, _
                                    ByVal dtReportingPeriod As Date, _
                                    ByVal arrVarPlTotalsByProject As Variant)
                                    
'for each range type (PL/Activity/Project)
'   Write LC/LC%
'For each Activity
'   for rev/costs create sum formula for each project under the activity
'For PL
'   for rev/costs create sum formula for each activity in the P&L

'Looping months
'   set local anchor
'   loop 0 to 11 for months


Dim i As Long, j As Long, k As Long
Dim n As Name
Dim rngLocalAnchor As Range
Dim arrVarProjectRangeNames As Variant
Dim strRevenueSumFormula As String
Dim strCostsSumFormula As String


For Each n In wsLcForecast.Names
    Debug.Print n.Name
        If GenericFunctions.StringSearch(1, n.Name, "Lc.Forecasts_Project.Name_") > 0 Then
            Set rngLocalAnchor = Range(n)(4, 4)
            For i = 0 To 11
                WriteLcFormulasToTable Range(rngLocalAnchor(3, i), rngLocalAnchor(4, i))
                WriteLcFormulasToTable Range(rngLocalAnchor(8, i), rngLocalAnchor(9, i))
            Next i
        ElseIf GenericFunctions.StringSearch(1, n.Name, "Lc.Forecasts_Activity.Name_") > 0 Then
            'get activity proper name from ws table
            'get list of projects from array
            'loop worksheet names again and generate string of ranges
            'write to activity table month
            Set rngLocalAnchor = Range(n)(3, 4)
            arrVarProjectRangeNames = GetArrayOfProjectRangeNames(arrVarPlTotalsByProject, Range(n)(1, 3).Value)
            
            For i = 0 To Month(dtReportingPeriod)
                strRevenueSumFormula = "=SUM("
                strCostsSumFormula = "=SUM("
                For j = 0 To UBound(arrVarProjectRangeNames, 1)
                    If Not IsEmpty(arrVarProjectRangeNames(j)) Then
                        strRevenueSumFormula = strRevenueSumFormula & wsLcForecast.Range(arrVarProjectRangeNames(j))(4 + 0, 3 + i).Address & ","
                        strCostsSumFormula = strCostsSumFormula & wsLcForecast.Range(arrVarProjectRangeNames(j))(4 + 1, 3 + i).Address & ","
                    End If
                Next j
                strRevenueSumFormula = Left(strRevenueSumFormula, Len(strRevenueSumFormula) - 1) & ")"
                rngLocalAnchor(1, i).Formula = strRevenueSumFormula
                strCostsSumFormula = Left(strCostsSumFormula, Len(strCostsSumFormula) - 1) & ")"
                rngLocalAnchor(2, i) = strCostsSumFormula
                'write lc,lc%
                    WriteLcFormulasToTable Range(rngLocalAnchor(3, i), rngLocalAnchor(4, i))
            Next i
        End If
Next n


End Sub

'@Description "return array of projects worksheet names based on activity"
Private Function GetArrayOfProjectRangeNames( _
                                        ByVal arrVarPlTotalsByProject As Variant, _
                                        ByVal strActivityName As String) As Variant

Dim i As Long, j As Long, k As Long
Dim arrStrProjectRangeNames() As Variant

'loop activities until match with current activity
'redim arr to match number of projects in PlTotals array
'concat range name for each project in activity

For i = 0 To UBound(arrVarPlTotalsByProject, 1)
    If arrVarPlTotalsByProject(i, 0) = strActivityName Then
        ReDim arrStrProjectRangeNames(0 To UBound(arrVarPlTotalsByProject(i, 1), 1))
        k = 0 'implement unique arrStrProjectRangeNames counter since there might be some values skipped from the PlTotals array
        For j = 0 To UBound(arrStrProjectRangeNames)
            If Not arrVarPlTotalsByProject(i, 1)(j, 0) = "no projects" Then 'ignore any "no projects"
                arrStrProjectRangeNames(k) = "Lc.Forecasts_Project.Name_" & _
                        GenericFunctions.replaceIllegalNamedRangeCharacters(arrVarPlTotalsByProject(i, 1)(j, 0))
                k = k + 1
            End If
            'Debug.Print arrStrProjectRangeNames(j)
        Next j
    End If
Next i

GetArrayOfProjectRangeNames = arrStrProjectRangeNames

End Function

'@Description "Writes the LC and LC% formulas to a give range"
Private Sub WriteLcFormulasToTable( _
                                      ByRef rngTarget As Range)
                                        
'rngTarget should be 2 row x 1 column range, indicating the LC and LC% rows for a given month
'LC formula = Rev + Cost (costs are normally a negative number)
'LC % formula = Rev / LC

rngTarget(1, 1).FormulaR1C1 = "=IF(AND(ISBLANK(R[-2]C[0]),ISBLANK(R[-1]C[0])),"""",SUM(R[-2]C[0],R[-1]C[0]))"
rngTarget(2, 1).FormulaR1C1 = "=IF(AND(ISBLANK(R[-3]C[0]),ISBLANK(R[-2]C[0])),"""",R[-3]C[0]/R[-1]C[0])"
         
End Sub


'@Description "Writes LC values for each activity/project prior to current month"
Private Sub WriteRevCostValuesToTables( _
                                    ByRef wsLcForecast As Worksheet, _
                                    ByRef wsAllocations As Worksheet, _
                                    ByVal arrVarPlTotalsByProject As Variant, _
                                    ByVal dtReportingPeriod As Date)
                                    
Dim i As Long, j As Long, k As Long, m As Long
Dim intTargetMonth As Integer
Dim n As Name
Dim rngLcTableAmountAnchor As Range
                                    
'loop activities
'   loop projects
'       populate months rev, cost for prior months from array
'       else populate from allocations ws
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

'Target month is month prior to reporting period month
    intTargetMonth = Month(dtReportingPeriod) - 1

For i = 0 To UBound(arrVarPlTotalsByProject, 1)
    For j = 0 To UBound(arrVarPlTotalsByProject(i, 1), 1)
        For Each n In wsLcForecast.Names
            If GenericFunctions.StringSearch(1, n.Name, GenericFunctions.replaceIllegalNamedRangeCharacters(arrVarPlTotalsByProject(i, 1)(j, 0))) > 0 Then
                Set rngLcTableAmountAnchor = Range(n)(4, 4)
                For k = 0 To intTargetMonth
                    For m = 0 To 1
                        If Not k = intTargetMonth Then
                            rngLcTableAmountAnchor(m + 1, k).Value = arrVarPlTotalsByProject(i, 1)(j, 1)(k + 1, m)
                            
                        ElseIf k = intTargetMonth And GenericFunctions.StringSearch(1, arrVarPlTotalsByProject(i, 1)(j, 0), "Not Assigned") = 0 Then
                            rngLcTableAmountAnchor(m + 1, k).Value = GetRevCostTotalFromAllocationsWs( _
                                                                        wsAllocations, _
                                                                        arrVarPlTotalsByProject(i, 0), _
                                                                        arrVarPlTotalsByProject(i, 1)(j, 0), _
                                                                        m)
                        Else
                            rngLcTableAmountAnchor(m + 1, k).Value = GetRevCostTotalForNotAllocated( _
                                                                        wsAllocations, _
                                                                        arrVarPlTotalsByProject(i, 0), _
                                                                        arrVarPlTotalsByProject(i, 1)(j, 1), _
                                                                        m)
                        End If
                        'Debug.Print arrVarPlTotalsByProject(i, 0), arrVarPlTotalsByProject(i, 1)(j, 0), arrVarPlTotalsByProject(i, 1)(j, 1)(k + 1, m)
                    Next m 'rev/cost
                Next k ' monhts
            End If
        Next n 'worksheet names
    Next j 'project loop
Next i 'activity loop

End Sub

'@Description "Returns a Rev/Cost total for "not allocated" amounts based on Allocations WS input"
Private Function GetRevCostTotalForNotAllocated( _
                                            ByRef wsAllocations As Worksheet, _
                                            ByVal strActivityName As String, _
                                            ByVal arrVarProjects As Variant, _
                                            ByVal intRevCostIndicator As Integer _
                                            ) As Variant
                        
'Get activity range on allocations ws
'set projects range by calculating activity range width
'set criteria range
'set activity total range
'generate [activity total] formula for rev/cost using criteria range and activity total range
'generate [allocated] formula for rev/cost using criteria range and projects range
'typical function format:
'   =SUMPRODUCT((([criteria range]=[critera_1])+([criteria range]=[critera_2])+...)*[sum range])
'general unallocated calculation:
'   unallocated = [activity total] - [allocated]

Dim i As Long, j As Long, k As Long
Dim rngActivity As Range
Dim rngProjects As Range
Dim rngActivityCriteria As Range
Dim rngActivityTotals As Range
Dim intProjectsCount As Integer
Dim arrVarCriteria As Variant
Dim strActivityTotalsFunction As String
Dim strProjectsTotalsFunction As String

'Set rev/cost criteria from constant
    If intRevCostIndicator = 0 Then arrVarCriteria = ARRAY_DESC_GROUPS_REV Else arrVarCriteria = ARRAY_DESC_GROUPS_COSTS

'Get number of projects from PlTotals array that are not "Not Allocated"
'    intProjectsCount = 0
'    For i = 0 To UBound(arrVarProjects, 1)
'        If GenericFunctions.StringSearch(1, arrVarProjects(i), "Not Allocated") = 0 Then intProjectsCount = intProjectsCount + 1
'    Next i

'Set activity range
    Set rngActivity = wsAllocations.Range("Allocations_Activity.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strActivityName))
    
'Calculate Projects Range and set
    Set rngProjects = Range(rngActivity(3, 6), rngActivity(rngActivity.Rows.Count, rngActivity.Columns.Count - 1))

'Set Criteria Range
    Set rngActivityCriteria = Range(rngActivity(3, 1), rngActivity(rngActivity.Rows.Count, 1))

'Set Activity Total Amount range
    Set rngActivityTotals = Range(rngActivity(3, 3), rngActivity(rngActivity.Rows.Count, 3))
    
'Generate functions
    strActivityTotalsFunction = "SUMPRODUCT(("
    strProjectsTotalsFunction = "SUMPRODUCT(("
    For i = 0 To UBound(arrVarCriteria, 1)
        strActivityTotalsFunction = strActivityTotalsFunction & _
                                    "(" & _
                                    rngActivityCriteria.Address(, , , True) & _
                                    "=" & _
                                    """" & arrVarCriteria(i) & """" & _
                                    ")+"
        strProjectsTotalsFunction = strProjectsTotalsFunction & _
                                    "(" & _
                                    rngActivityCriteria.Address(, , , True) & _
                                    "=" & _
                                    """" & arrVarCriteria(i) & """" & _
                                    ")+"
    Next i
    strActivityTotalsFunction = Left(strActivityTotalsFunction, Len(strActivityTotalsFunction) - 1) & ")*" & rngActivityTotals.Address(, , , True) & ")"
    strProjectsTotalsFunction = Left(strProjectsTotalsFunction, Len(strProjectsTotalsFunction) - 1) & ")*" & rngProjects.Address(, , , True) & ")"
        
    GetRevCostTotalForNotAllocated = "=" & strActivityTotalsFunction & "-" & strProjectsTotalsFunction



End Function


'@Description "Generates a formula that returns a value from the allocations ws based on the project"
Private Function GetRevCostTotalFromAllocationsWs( _
                                                    ByRef wsAllocations As Worksheet, _
                                                    ByVal strActivityName As String, _
                                                    ByVal strProjectName As String, _
                                                    ByVal intRevCostIndicator As Integer) _
                                                    As Variant

'generate sum range name based on project name
'generate criteria range name based on activity name
'generate sum range from sum range name
'generate criteria range from range name

Dim rngActivityRange As Range
Dim rngProjectRange As Range
Dim rngSumRange As Range
Dim rngCriteriaRange As Range
Dim arrVarCriteria As Variant

If intRevCostIndicator = 0 Then arrVarCriteria = ARRAY_DESC_GROUPS_REV Else arrVarCriteria = ARRAY_DESC_GROUPS_COSTS

Set rngActivityRange = wsAllocations.Range("Allocations_Activity.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strActivityName))
Set rngProjectRange = wsAllocations.Range("Allocations_Project.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strProjectName))

Set rngSumRange = wsAllocations.Range(rngProjectRange(3, 1), rngProjectRange(rngProjectRange.Rows.Count, 1))
Set rngCriteriaRange = wsAllocations.Range(rngActivityRange(3, 1), rngActivityRange(rngActivityRange.Rows.Count, 1))

GetRevCostTotalFromAllocationsWs = "=SUM(SUMIFS(" & rngSumRange.Address(, , , True) & ", " & rngCriteriaRange.Address(, , , True) & ", {""" & Join(arrVarCriteria, """, """) & """}))"

End Function

                               


'@Description "Loops activities/projects to write blank table and set named range for each"
Private Sub WriteBlankTablesToLcWorksheet( _
                                        ByRef objPl As clsPandL, _
                                        ByRef wsLcForecast As Worksheet, _
                                        ByVal arrVarPlTotalsByProject As Variant, _
                                        ByVal dtReportingPeriod As Date)
                                        
'loop activities
'   write project header
'   write row headers
'   write month header
'   assign to named range
'
'loop projects
'   write project header
'   write row headers
'   write month header
'   assign to named range
'next project
'group activites/projects
'next activity

Dim i As Long, k As Long, j As Long
Dim rngTopAnchor            As Range
Dim intLcTableRowOffset        As Integer
Dim intPlTableRowOffset        As Integer
Dim intAnchorRowOffset         As Integer
Dim intMonthStartCol        As Integer
Dim intRowHeaderStartRow    As Integer
Dim intWsRowOffset          As Integer
Dim intTableRowOffset       As Integer

'Constants
    intLcTableRowOffset = 2
    intPlTableRowOffset = 11
    intMonthStartCol = 3
    'intRowHeaderStartRow = 3  (probably not required)


intAnchorRowOffset = (intLcTableRowOffset * 2) + intPlTableRowOffset

'Set top anchor range
    Set rngTopAnchor = wsLcForecast.Range("Lc.Forecast_Top.Anchor")
    
'Set PL named range
    wsLcForecast.Names.Add Name:="LC.Forecast_Pl.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(objPl.strName), _
                                RefersTo:=Range(rngTopAnchor(2, 1), rngTopAnchor(12, intMonthStartCol + 11))

    
'loop activities
'   write table for activity and set named range
'   advance table row offset
'   loop projects
'       write table for project and set named range
'       advance table row offset
'   next project
'next activity

For i = 0 To UBound(arrVarPlTotalsByProject, 1)
    intTableRowOffset = writeBlankTable(wsLcForecast, dtReportingPeriod, rngTopAnchor(intAnchorRowOffset, 1), intMonthStartCol, arrVarPlTotalsByProject(i, 0))
    intAnchorRowOffset = intAnchorRowOffset + intTableRowOffset + intLcTableRowOffset
    
    For j = 0 To UBound(arrVarPlTotalsByProject(i, 1), 1)
        If Not arrVarPlTotalsByProject(i, 1)(j, 0) = "no projects" Then
            intTableRowOffset = writeBlankTable(wsLcForecast, dtReportingPeriod, rngTopAnchor(intAnchorRowOffset, 1), intMonthStartCol, arrVarPlTotalsByProject(i, 0), arrVarPlTotalsByProject(i, 1)(j, 0))
            intAnchorRowOffset = intAnchorRowOffset + intTableRowOffset + intLcTableRowOffset
        End If
    Next j
    
Next i
    
End Sub

'@Description "Writes a blank LC table to the worksheet and sets the named range"
Private Function writeBlankTable( _
                                ByRef wsLcForecast As Worksheet, _
                                ByVal dtReportingDate As Date, _
                                ByRef rngLocalAnchor As Range, _
                                ByVal intMonthStartCol As Integer, _
                                ByVal strActivityName As String, _
                                Optional ByVal strProjectName As String) _
                                As Integer

'Write header (activity, project)
'Write months
'Write row header section title (forecast/actual)
'Write row headers
'set range

Dim i As Long, j As Long, k As Long
Dim intRowOffset As Integer

intRowOffset = 1

'Write Activity Header
    rngLocalAnchor(intRowOffset, 1) = "Activity Name"
    PAFCellFormats.FormatProjectListHeaderActivityTitle rngLocalAnchor(intRowOffset, 1)
    PAFCellFormats.FormatAllBordersWhiteThin rngLocalAnchor(intRowOffset, 1)
    rngLocalAnchor(intRowOffset, 3) = strActivityName
    intRowOffset = intRowOffset + 1
    
'Write Project Header if present
    If Not strProjectName = "" Then
        rngLocalAnchor(intRowOffset, 1) = "Project Name"
        PAFCellFormats.FormatProjectListHeaderActivityTitle rngLocalAnchor(intRowOffset, 1)
        PAFCellFormats.FormatAllBordersWhiteThin rngLocalAnchor(intRowOffset, 1)
        
        'Because "Not Assigned" shows up multiple times we need to a special activity+not assigned named range
        'this doesnt look good in the ws so we override the typical naming from the array with fixed value when
        '"Not Assigned" is found in the project name array
            If GenericFunctions.StringSearch(1, strProjectName, "Not Assigned") Then
                rngLocalAnchor(intRowOffset, 3) = "Not Assigned"
            Else
                rngLocalAnchor(intRowOffset, 3) = strProjectName
            End If
            
        intRowOffset = intRowOffset + 1
    End If
    
'Write month column headers based on current year
    PAFCellFormats.FormatLcMonthColumnHeader _
        Range(rngLocalAnchor(intRowOffset, intMonthStartCol), rngLocalAnchor(intRowOffset, intMonthStartCol + 11))
    For i = 1 To 12
        rngLocalAnchor(intRowOffset, intMonthStartCol + (i - 1)).Value = MonthName(i, True) & "-" & Year(dtReportingDate)
    Next i
    intRowOffset = intRowOffset + 1
    
'Write Row Header Titles
    rngLocalAnchor(intRowOffset, 1) = "Actual"
        LcForecastMergeRowHeaderTitle rngLocalAnchor(intRowOffset, 1)
        
    rngLocalAnchor(intRowOffset + 5, 1) = "Forecast"
        LcForecastMergeRowHeaderTitle rngLocalAnchor(intRowOffset + 5, 1)

'Write Row Headers
    For i = 1 To 2
        rngLocalAnchor((intRowOffset), 2) = "Revenue"
        rngLocalAnchor((intRowOffset) + 1, 2) = "Cost"
        rngLocalAnchor((intRowOffset) + 2, 2) = "LC"
        rngLocalAnchor((intRowOffset) + 3, 2) = "LC%"
        
        'Apply amount USD/percentage formatting for the montly values
            PAFCellFormats.FormatAmountUsd _
                        Range(rngLocalAnchor((intRowOffset), intMonthStartCol), rngLocalAnchor((intRowOffset) + 2, intMonthStartCol + 11)), 11
            PAFCellFormats.FormatLcPercentage _
                    Range(rngLocalAnchor((intRowOffset) + 3, intMonthStartCol), rngLocalAnchor((intRowOffset) + 3, intMonthStartCol + 11))
        
        If i = 1 Then intRowOffset = intRowOffset + 5 Else intRowOffset = intRowOffset + 3
    Next i
    
    If Not strProjectName = "" Then
        SetTableRange wsLcForecast, rngLocalAnchor, intRowOffset, intMonthStartCol, strActivityName, strProjectName
    Else
        SetTableRange wsLcForecast, rngLocalAnchor, intRowOffset, intMonthStartCol, strActivityName
    End If
    
writeBlankTable = intRowOffset

End Function

'@Description "Merges and applies formatting for the Row Header Title"
Private Sub LcForecastMergeRowHeaderTitle( _
                                                rngTopCell As Range)
                                                
Range(rngTopCell, rngTopCell(4, 1)).Merge
PAFCellFormats.FormatLcRowHeaderTitle Range(rngTopCell, rngTopCell(4, 1))


End Sub

'@Description "Calculates, generates name, and sets a named range for caller table"
Private Sub SetTableRange( _
                            ByRef wsLcForecast As Worksheet, _
                            ByRef rngLocalAnchor As Range, _
                            ByVal intRowOffset As Integer, _
                            ByVal intMonthStartCol As Integer, _
                            ByVal strActivityName As String, _
                            Optional ByVal strProjectName As String)
                            
'get local anchor
'get rowoffset
'get column month offest
'calc right bottom corner = local anchor (rowoffset, col month offset + 11)

Dim rngTarget As Range
Dim strRangeName As String

'Set range
    Set rngTarget = Range(rngLocalAnchor(1, 1), rngLocalAnchor(intRowOffset, intMonthStartCol + 11))

'Set range name
    If strProjectName = "" Then
        strRangeName = "Lc.Forecasts_Activity.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strActivityName)
    Else
        If strProjectName = "Not Assigned" Then
            strRangeName = "Lc.Forecasts_Project.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strActivityName) & _
                            "_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strProjectName)
        Else
            strRangeName = "Lc.Forecasts_Project.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(strProjectName)
        End If
    End If

'Set named range
    wsLcForecast.Names.Add Name:=strRangeName, RefersTo:=rngTarget
    
End Sub


