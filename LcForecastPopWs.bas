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

'Set worksheet to object
    Set wsLcForecast = wbPaf.Worksheets("LC Forecast")
    
'Write activity/project blank tables and set range names
    WriteBlankTablesToLcWorksheet objPl, wsLcForecast, arrVarPlTotalsByProject, dtReportingPeriod
    
'Write project Lc/Lc% values to tables
    WriteLcValuesToTables wsLcForecast, arrVarPlTotalsByProject, dtReportingPeriod

End Sub

'@Description "Writes LC values for each activity/project prior to current month"
Private Sub WriteLcValuesToTables( _
                                    ByRef wsLcForecast As Worksheet, _
                                    ByVal arrVarPlTotalsByProject As Variant, _
                                    ByVal dtReportingPeriod As Date)
                                    
Dim i As Long, j As Long, k As Long, m As Long
Dim intTargetMonth As Integer
Dim n As Name
Dim rngLcTableAmountAnchor As Range

                                    
'loop activities
'   loop projects
'       populate months lc, lc%
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
                Set rngLcTableAmountAnchor = Range(n)(4, 3)
                For k = 0 To intTargetMonth - 1
                    For m = 0 To 1
                        rngLcTableAmountAnchor(m + 1, k).Value = arrVarPlTotalsByProject(i, 1)(j, 1)(k + 1, m)
                        Debug.Print arrVarPlTotalsByProject(i, 0), arrVarPlTotalsByProject(i, 1)(j, 0), arrVarPlTotalsByProject(i, 1)(j, 1)(k + 1, m)
                    Next m 'rev/cost
                Next k ' monhts
            End If
        Next n 'worksheet names
    Next j 'project loop
Next i 'activity loop

End Sub

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
    intMonthStartCol = 2
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
    rngLocalAnchor(intRowOffset, 0) = "Activity Name"
    rngLocalAnchor(intRowOffset, 2) = strActivityName
    intRowOffset = intRowOffset + 1
    
'Write Project Header if present
    If Not strProjectName = "" Then
        rngLocalAnchor(intRowOffset, 0) = "Project Name"
        
        'Because "Not Assigned" shows up multiple times we need to a special activity+not assigned named range
        'this doesnt look good in the ws so we override the typical naming from the array with fixed value when
        '"Not Assigned" is found in the project name array
            If GenericFunctions.StringSearch(1, strProjectName, "Not Assigned") Then
                rngLocalAnchor(intRowOffset, 2) = "Not Assigned"
            Else
                rngLocalAnchor(intRowOffset, 2) = strProjectName
            End If
        intRowOffset = intRowOffset + 1
    End If
    
'Write month column headers based on current year
    For i = 1 To 12
        rngLocalAnchor(intRowOffset, intMonthStartCol + (i - 1)).Value = MonthName(i, True) & "-" & Year(dtReportingDate)
    Next i
    intRowOffset = intRowOffset + 1
    
'Write Row Header Titles
    rngLocalAnchor(intRowOffset, 0) = "Actual"
    rngLocalAnchor(intRowOffset + 5, 0) = "Forecast"

'Write Row Headers
    For i = 1 To 2
        rngLocalAnchor((intRowOffset), 1) = "Revenue"
        rngLocalAnchor((intRowOffset) + 1, 1) = "Cost"
        rngLocalAnchor((intRowOffset) + 2, 1) = "LC"
        rngLocalAnchor((intRowOffset) + 3, 1) = "LC%"
        If i = 1 Then intRowOffset = intRowOffset + 5 Else intRowOffset = intRowOffset + 3
    Next i
    
    If Not strProjectName = "" Then
        SetTableRange wsLcForecast, rngLocalAnchor, intRowOffset, intMonthStartCol, strActivityName, strProjectName
    Else
        SetTableRange wsLcForecast, rngLocalAnchor, intRowOffset, intMonthStartCol, strActivityName
    End If
    
writeBlankTable = intRowOffset

End Function

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


