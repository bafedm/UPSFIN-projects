Attribute VB_Name = "AllocationsProjectInputFormat"
'@Folder("Generate PAF Workbook.WS Allocations")
Option Explicit

'@Description "Formats a cell with a blue input box when there is a value in the associated activity USD column
Sub Main( _
            ByRef wsAllocations As Worksheet)

Dim i As Long, j As Long, k As Long
Dim nActivityRange          As Name     'A ws Range Name with the prefix "Allocations_Activity.Name_"
Dim rngActivityRange        As Range    'The range associated with nActivityRange
Dim nProjectRange           As Name     'A ws Range Name with the prefix "Allocations_Project.Name_"
Dim rngActivityAmountUSD    As Range    'The Amount USD column range from an Activity Range


'Main loop, for each name search for project name prefix
    For Each nProjectRange In wsAllocations.Names
        If GenericFunctions.StringSearch(1, nProjectRange.Name, "Allocations_Project.Name_") > 0 And Not GenericFunctions.StringSearch(1, nProjectRange.Name, "no.projects") > 0 Then
            
            'if project range found call sub to determine parent activity range name
                Set nActivityRange = GetIntersectingRangeNameConditional(wsAllocations.Range(nProjectRange)(1, 1))
            
            'Get the range from the associated activity name
                Set rngActivityRange = wsAllocations.Range(nActivityRange)
            
            'Get the range containing the Amount USD column for the activity
                Set rngActivityAmountUSD = wsAllocations.Range(rngActivityRange(1, 1).Offset(0, 2), rngActivityRange(1, 1).Offset(rngActivityRange.Rows.Count - 1, 2))
            
            'For each cell in the project range check to see if associated amount USD cell has a value
            'If it has a value set background blue and change number format
            For i = 3 To Range(nProjectRange).Rows.Count - 1
                If Not IsEmpty(rngActivityAmountUSD(1, 1).Offset(i, 0).Value) Then
                    With Range(nProjectRange)(1, 1)
                        PAFCellFormats.FormatAmountUsd .Offset(i, 0)
                        PAFCellFormats.FormatAllBordersWhiteThin .Offset(i, 0)
                        .Offset(i, 0).Interior.Color = RGB(204, 244, 255)
                    End With
                End If
            Next i
            
        End If
    Next nProjectRange

End Sub

'@Description "Returns the name of the range where a cell resides"
Private Function GetIntersectingRangeNameConditional( _
                                            ByRef rngLocalAnchor As Range) As Name

Dim nWorksheetNamedRange    As Name     'A named range in the worksheet
Dim nCurrentRegion          As Name     'The named range to be returned to caller
Dim varIntersectTest        As Variant  'returns a range where the cells intersect


For Each nWorksheetNamedRange In rngLocalAnchor.Parent.Names
    If GenericFunctions.StringSearch(1, nWorksheetNamedRange.Name, "Allocations_Activity.Name_") Then
        Set varIntersectTest = Application.Intersect(rngLocalAnchor, nWorksheetNamedRange.RefersToRange)
        If Not varIntersectTest Is Nothing Then
            Set nCurrentRegion = nWorksheetNamedRange
        End If
    End If
Next nWorksheetNamedRange

Set GetIntersectingRangeNameConditional = nCurrentRegion
            
End Function

