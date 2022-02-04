Attribute VB_Name = "AllocationsWriteToWs"
'@Folder("Generate PAF Workbook.WS Allocations")
'@Description "Methods related to writing PL data to worksheet"
Option Explicit

'@Description "Main loop"
Sub Main( _
            ByRef wsAllocations As Worksheet, _
            ByRef wsProjects As Worksheet, _
            ByRef objTargetPl As clsPandL, _
            ByRef dtReportingPeriod As Date, _
            ByRef collActivities As Collection, _
            ByVal arrVarMasterPlList As Variant)

'Write P&L total and set range
'loop activities
'write activity
'check project list from project ws
'write project column headers and set range
'set activity range
'next activity
'set allocation % cf and formulas

Dim rngLeftAnchor                       As Range        'Fixed page anchor for ranges
Dim intAllocationsOffsetLeft            As Integer      'The column offset from left anchor
Dim arrVarSingleColumnDescGroupDesc     As Variant      'Single column containing desc group and desc item
Dim objActivity                         As clsActivity  'Activity object from collActivity
Dim objPl                               As clsPandL     'P&L object from a collPls
Dim nmeProjectListActivity              As Variant      'Named ranges from the Project List worksheet
    
'set anchor variable and initial offset value
    Set rngLeftAnchor = wsAllocations.Range("Allocations_Left.Anchor")
    intAllocationsOffsetLeft = 0

'Write the full PL array to ws and increment offset
    arrVarSingleColumnDescGroupDesc = GenerateSingleColumnDescGroupDesc(arrVarMasterPlList)
    WritePlAllocationAndReturnLeftOffset objTargetPl, arrVarMasterPlList, arrVarSingleColumnDescGroupDesc, rngLeftAnchor, intAllocationsOffsetLeft
    intAllocationsOffsetLeft = intAllocationsOffsetLeft + 5
    
'write activites/projects retrieved from Project List worksheet then increment array
    For Each nmeProjectListActivity In wsProjects.Names
        If GenericFunctions.StringSearch(1, nmeProjectListActivity.Name, "Project.List_Activity.Name_") Then
            
            'Write the activity header cells
                WriteActivityAllocationHeader rngLeftAnchor.Offset(0, intAllocationsOffsetLeft)
                
          
            'Write and format the activity name
                PAFCellFormats.FormatAllocationsDGDListFontDefault rngLeftAnchor.Offset(2, intAllocationsOffsetLeft)
                rngLeftAnchor.Offset(2, intAllocationsOffsetLeft).Value = wsProjects.Range(nmeProjectListActivity)(1, 2).Value
                           
            'Write the activity details to the ws
                WriteActivityToAllocationsWs _
                        wsProjects, objTargetPl, dtReportingPeriod, _
                        collActivities(wsProjects.Range(nmeProjectListActivity)(1, 2).Value), _
                        arrVarSingleColumnDescGroupDesc, rngLeftAnchor, intAllocationsOffsetLeft
                                   
            'Increment offset for projects column
                intAllocationsOffsetLeft = intAllocationsOffsetLeft + 3
                        
            'Write the project columns
                intAllocationsOffsetLeft = intAllocationsOffsetLeft + WriteProjectColumnsAndReturnLeftOffset(rngLeftAnchor.Offset(0, intAllocationsOffsetLeft), wsProjects.Range(nmeProjectListActivity))
                           
            'Increment offset
                'intAllocationsOffsetLeft = intAllocationsOffsetLeft + 5
                
    End If
    
Next nmeProjectListActivity

End Sub

'@Description "Writes a column header for each activity"
Private Function WriteProjectColumnsAndReturnLeftOffset( _
                                    ByRef rngLocalAnchor As Range, _
                                    ByRef rngProjectListActivity As Range)

Dim i As Long, j As Long, k As Long
Dim arrVarProjectsList  As Variant  'An array of projects from the Project List range passed from caller
Dim strRangeName        As String   'Project name for the named range
Dim rngNamedRange       As Range    'Project range for the named range
Dim nCurrentRegion      As Name     'Named range that intersects current cell

'Resize array based on number of projects
    ReDim arrVarProjectsList(0 To (rngProjectListActivity.Rows.Count - 4))

'build projects list
    j = 0
    For i = 3 To rngProjectListActivity.Rows.Count - 1
        arrVarProjectsList(j) = rngProjectListActivity(i, 2)
        j = j + 1
    Next i

'projects loop
For i = 0 To UBound(arrVarProjectsList, 1)
    With rngLocalAnchor
        'insert new column
            .Offset(0, i).EntireColumn.Insert
        'Set project name and format.  If "no projects" then custom format
            .Offset(2, i).Value = arrVarProjectsList(i)
                If arrVarProjectsList(i) = "no projects" Then
                    PAFCellFormats.FormatHiddenText .Offset(2, i)
                    PAFCellFormats.FormatAllBordersWhiteThin .Offset(2, i)
                    .Offset.Font.Italic = True
                Else
                    PAFCellFormats.FormatDefaultFont .Offset(2, i)
                    PAFCellFormats.FormatAllBordersWhiteThin .Offset(2, i)
                    .Offset(2, i).HorizontalAlignment = xlCenter
                End If
        'set text "Amount USD" and format
            .Offset(3, i).Value = "Amount USD"
                .Offset(3, i).ColumnWidth = 16.3
                PAFCellFormats.FormatActivityListHeaderRow .Offset(3, i)
                PAFCellFormats.FormatBordersBlackBottomThinWhiteThin .Offset(3, i)
                .Offset(3, i).HorizontalAlignment = xlCenter
        
        'Get the name of the activity range, used to determine the number of rows for the project range
            Set nCurrentRegion = GetIntersectingRangeName(rngLocalAnchor.Offset(3, i))

        'Build named range name using prefix & the project name
            If arrVarProjectsList(i) = "no projects" Then
                strRangeName = "Allocations_Project.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(.Offset(2, -4).Value) & "_" & _
                                GenericFunctions.replaceIllegalNamedRangeCharacters(arrVarProjectsList(i))
            Else
                strRangeName = "Allocations_Project.Name_" & GenericFunctions.replaceIllegalNamedRangeCharacters(arrVarProjectsList(i))
            End If
            
        'set the named range range using the current column start cell and the end with row obtained from activity range row count
            Set rngNamedRange = .Parent.Range(.Offset(2, i), .Offset(nCurrentRegion.RefersToRange.Rows.Count + 1, i))
    
        'add the worksheet name
            .Parent.Names.Add Name:=strRangeName, RefersTo:=rngNamedRange
    
    End With
Next i

'resize the column between the activity totals and the project columns
    rngLocalAnchor.Offset(0, -1).ColumnWidth = 2

'Add a "Projects" header to the first column of the projects section and format
    rngLocalAnchor.Offset(1, 0).Value = "Projects"
    PAFCellFormats.FormatProjectListHeaderActivityTitle rngLocalAnchor.Offset(1, 0)
    PAFCellFormats.FormatAllBordersWhiteThin rngLocalAnchor.Offset(1, 0)
    
    WriteProjectColumnsAndReturnLeftOffset = UBound(arrVarProjectsList, 1) + 4

End Function

'@Description "Returns the name of the range where a cell resides"
Private Function GetIntersectingRangeName( _
                                            ByRef rngLocalAnchor As Range) As Name

Dim nWorksheetNamedRange    As Name     'A named range in the worksheet
Dim nCurrentRegion          As Name     'The named range to be returned to caller
Dim varIntersectTest        As Variant  'returns a range where the cells intersect


For Each nWorksheetNamedRange In rngLocalAnchor.Parent.Names
    Set varIntersectTest = Application.Intersect(rngLocalAnchor, nWorksheetNamedRange.RefersToRange)
    If Not varIntersectTest Is Nothing Then
        Set nCurrentRegion = nWorksheetNamedRange
    End If
Next nWorksheetNamedRange

Set GetIntersectingRangeName = nCurrentRegion
            
End Function

'@Description "writes an individual activity finance table to the ws"
Private Sub WriteActivityToAllocationsWs( _
                                                                    ByRef wsProjects As Worksheet, _
                                                                    ByRef objTargetPl As clsPandL, _
                                                                    ByVal dtReportingPeriod As Date, _
                                                                    ByRef objActivity As clsActivity, _
                                                                    ByVal arrVarSingleColumnDescGroupDesc As Variant, _
                                                                    ByRef rngLeftAnchor As Range, _
                                                                    ByVal intAllocationsOffsetLeft As Integer _
                                                                    )

Dim arrVarMasterActivityPlList  As Variant  'Holds the full activity array

'Create master activity sorted array from dictionary item
    arrVarMasterActivityPlList = AllocationsFinanceTableArray.GenerateActivityCombinedPlTableAsArray(objTargetPl, objActivity, dtReportingPeriod)

'write array to ws
    WritePlArrayToWs rngLeftAnchor, intAllocationsOffsetLeft, arrVarSingleColumnDescGroupDesc, arrVarMasterActivityPlList, "Allocations_Activity.Name_", _
                                GenericFunctions.replaceIllegalNamedRangeCharacters(objActivity.strName)
                                                
End Sub
'@Description "Writes the PL Total Allocation Range"
Private Function WritePlAllocationAndReturnLeftOffset( _
                                                        ByRef objTargetPl As clsPandL, _
                                                        ByVal arrVarMasterPlList As Variant, _
                                                        ByVal arrVarSingleColumnDescGroupDesc As Variant, _
                                                        ByRef rngLeftAnchor As Range, _
                                                        ByVal intAllocationsOffsetLeft As Integer) _
                                                        As Integer

'Writes the array to the worksheet
    WritePlArrayToWs rngLeftAnchor, intAllocationsOffsetLeft, arrVarSingleColumnDescGroupDesc, arrVarMasterPlList, "Allocations_PL.Name_", _
                   GenericFunctions.replaceIllegalNamedRangeCharacters(objTargetPl.strName)
   
End Function

'@Description "Writes the standard activity header, does not include project columns
Private Sub WriteActivityAllocationHeader( _
                                            ByRef rngLocalAnchor As Range)

With rngLocalAnchor
    .Offset(1, 0).Value = "Activity"
        PAFCellFormats.FormatProjectListHeaderActivityTitle .Offset(1, 0)
        PAFCellFormats.FormatAllBordersWhiteThin .Offset(1, 0)
    .Offset(3, 0).Value = "Description"
    .Offset(3, 1).Value = "Amount USD"
    .Offset(3, 2).Value = "% Allocated"
        PAFCellFormats.FormatActivityListHeaderRow .Parent.Range(.Offset(3, 0), .Offset(3, 2))
        PAFCellFormats.FormatBordersBlackBottomThinWhiteThin .Parent.Range(.Offset(3, 0), .Offset(3, 2))
End With

End Sub


'@Description "Writes a combined PL array to the worksheet and sets the named range"
Private Sub WritePlArrayToWs( _
                                ByRef rngLeftAnchor As Range, _
                                ByVal intAllocationsOffsetLeft As Integer, _
                                ByVal arrVarSingleColumnDescGroupDesc As Variant, _
                                ByVal arrVarCombinedPlList As Variant, _
                                ByVal strNamedRangePrefix As String, _
                                ByVal strRangeReferenceName As String, _
                                Optional ByVal intLocalRowOffset = 4, _
                                Optional ByVal intLocalColOffset = 0)


Dim i As Long, j As Long, k As Long
Dim rngLocalAnchor      As Range    'The local anchor used for offset locations

'Set the local anchor to the main achor + the passed column offset
    Set rngLocalAnchor = rngLeftAnchor.Offset(0, intAllocationsOffsetLeft)

With rngLocalAnchor
    'loop all desc group, desc items (dg/d)
    For i = 0 To UBound(arrVarSingleColumnDescGroupDesc, 1)
        'write dg/d element to cell and format
            .Offset(intLocalRowOffset + i, intLocalColOffset).Value = arrVarSingleColumnDescGroupDesc(i, 1)
            .Offset(intLocalRowOffset + i, intLocalColOffset - 1).Value = arrVarSingleColumnDescGroupDesc(i, 0)
            PAFCellFormats.FormatAllocationsDGDListFontDefault .Offset(intLocalRowOffset + i, intLocalColOffset)
            PAFCellFormats.FormatAllBordersWhiteThin .Offset(intLocalRowOffset + i, intLocalColOffset)
        'for each dg/d write, check against master pl list to see if it's a dg or a d
            For j = 0 To UBound(arrVarCombinedPlList, 1)
                'if it's a d, write it's dg to the left (for lc forecast) and the amount usd to the write.  format.
                    If arrVarCombinedPlList(j, 2) = arrVarSingleColumnDescGroupDesc(i, 1) Then
                        'amount usd col
                            .Offset(intLocalRowOffset + i, intLocalColOffset + 1) = arrVarCombinedPlList(j, 3)
                                PAFCellFormats.FormatAmountUsd .Offset(intLocalRowOffset + i, intLocalColOffset + 1)
                                PAFCellFormats.FormatAllBordersWhiteThin .Offset(intLocalRowOffset + i, intLocalColOffset + 1)
                'if its a dg, format
                    ElseIf .Offset(intLocalRowOffset + i, intLocalColOffset) = .Offset(intLocalRowOffset + i, intLocalColOffset - 1) Then
                        PAFCellFormats.FormatDescGroupHeader .Offset(intLocalRowOffset + i, intLocalColOffset)
                    End If
                    
                    
            Next j
    Next i

End With 'rngLocalAnchor

'Set named range
    Dim strRangeName As String
    Dim rngNamed As Range
    
    strRangeName = strNamedRangePrefix & strRangeReferenceName
    Set rngNamed = rngLeftAnchor.Parent.Range(rngLocalAnchor.Offset(2, -1), rngLocalAnchor(4 + UBound(arrVarSingleColumnDescGroupDesc, 1) + 1, 5))
    
    rngLeftAnchor.Parent.Names.Add Name:=strRangeName, RefersTo:=rngNamed

'hide desc_group column
    rngLocalAnchor.Offset(0, -1).EntireColumn.Hidden = True

'resize dg/d and amount USD column to fit text
    rngLocalAnchor.Parent.Range(rngLocalAnchor(1, 1), rngLocalAnchor(1, 3)).EntireColumn.AutoFit



End Sub

'@Description "Creates a single column with desc group followed by desc group desc items"
Private Function GenerateSingleColumnDescGroupDesc( _
                                                    arrVarMasterPlList As Variant) As Variant
                                                    
Dim i As Long, j As Long, k As Long
Dim arrVarSingleColumnDescGroupDesc     As Variant  'single dimension array with desc group and desc items
Dim intDescGroupCount                   As Integer  'counts the number of desc group items for a given list for an array redim

'Start at one then compare desc group column items, increment when there is a change (assumes order preserved)
    intDescGroupCount = 1

    For i = 1 To UBound(arrVarMasterPlList, 1)
        If Not arrVarMasterPlList(i, 1) = arrVarMasterPlList(i - 1, 1) Then intDescGroupCount = intDescGroupCount + 1
    Next i

'redim array based on desc group count + the number of rows in the full array
    ReDim arrVarSingleColumnDescGroupDesc(0 To intDescGroupCount + UBound(arrVarMasterPlList, 1), 0 To 1)

'set initial desc group
    arrVarSingleColumnDescGroupDesc(0, 0) = arrVarMasterPlList(0, 1)
    arrVarSingleColumnDescGroupDesc(1, 0) = arrVarMasterPlList(0, 1)
    arrVarSingleColumnDescGroupDesc(0, 1) = arrVarMasterPlList(0, 1)
    arrVarSingleColumnDescGroupDesc(1, 1) = arrVarMasterPlList(0, 2)

'populate remainder of array.  compare master list desc group item with current
'if match then write only the desc and advance single col counter
'if not match then write desc group, advance, write desc, advance
    j = 2
    For i = 1 To UBound(arrVarMasterPlList, 1)
        If arrVarMasterPlList(i, 1) = arrVarMasterPlList(i - 1, 1) Then
            arrVarSingleColumnDescGroupDesc(j, 0) = arrVarMasterPlList(i, 1)
            arrVarSingleColumnDescGroupDesc(j, 1) = arrVarMasterPlList(i, 2)
            j = j + 1
        Else
            arrVarSingleColumnDescGroupDesc(j, 0) = arrVarMasterPlList(i, 1)
            arrVarSingleColumnDescGroupDesc(j + 1, 0) = arrVarMasterPlList(i, 1)
            arrVarSingleColumnDescGroupDesc(j, 1) = arrVarMasterPlList(i, 1)
            arrVarSingleColumnDescGroupDesc(j + 1, 1) = arrVarMasterPlList(i, 2)
            j = j + 2
        End If
    Next i

'return result
    GenerateSingleColumnDescGroupDesc = arrVarSingleColumnDescGroupDesc
 
End Function
