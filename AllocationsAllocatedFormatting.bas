Attribute VB_Name = "AllocationsAllocatedFormatting"
'@Folder("Generate PAF Workbook.WS Allocations")
'@Description "Writes the formulas and cf for the allocated% columns"
Option Explicit

'@Description "Main Loop.  build formula write the allocated% formula to the associated column for each range"
Sub Main( _
            ByRef wsAllocations As Worksheet)

'for each activity range
'get number of projects by range size calc
'for each row
'create formula to add projects amount

Dim i As Long, j As Long, k As Long
Dim nWsRange                        As Name     'A named range in the worksheet
Dim rngAllocated                    As Range    'The allocated% range for a given activity
Dim intProjectCount                 As Integer  'The number projects for an acitivty based on activity column count
Dim arrVarProjectCountPerActivity   As Variant  'An array that holds an activity name and the number of projects

'Resize array based on number of activity ranges in ws
    ReDim arrVarProjectCountPerActivity(0 To CountActivityRanges(wsAllocations, "Allocations_Activity.Name_") - 1, 0 To 1)
   
'Loop activites, check for activity range names
    i = 0
    For Each nWsRange In wsAllocations.Names
        If GenericFunctions.StringSearch(1, nWsRange.Name, "Allocations_Activity.Name_") > 0 Then
            'Set allocated column of activity range to a variable
                Set rngAllocated = wsAllocations.Range(wsAllocations.Range(nWsRange)(3, 4), wsAllocations.Range(nWsRange)(wsAllocations.Range(nWsRange).Rows.Count, 4))
            
            'If the first project is NOT named "no projects" then calculate the number of projects
            'by subtracting the activity columns from the number of columns without project columns
            'If it is no projects then set project count to zero
            'Also store activity name and project count to array for the pl total allocated% formula
                If Not rngAllocated(1, 1).Offset(-2, 2).Value = "no projects" Then
                    intProjectCount = rngAllocated.Range(nWsRange).Columns.Count - 6: arrVarProjectCountPerActivity(i, 0) = nWsRange: arrVarProjectCountPerActivity(i, 1) = intProjectCount
                    i = i + 1
                Else
                    intProjectCount = 0: arrVarProjectCountPerActivity(i, 0) = nWsRange: arrVarProjectCountPerActivity(i, 1) = intProjectCount
                    i = i + 1
                End If
        End If
        
        'For each cell in the allocated% range build the formula based on the project count and row number
            WriteAcitivityAllocationFormulaAndFormat rngAllocated, intProjectCount
        
    Next nWsRange
    
'write Pl allocated% column formulas and format
    WritePlActivityAllocationFormulaAndFormat wsAllocations, arrVarProjectCountPerActivity

End Sub

'@Description "generates and applies allocated% formula and sets formatting for the pl total"
Private Sub WritePlActivityAllocationFormulaAndFormat( _
                                                        ByRef wsAllocations As Worksheet, _
                                                        ByVal arrVarProjectCountPerActivity As Variant)
Dim i As Long, j As Long, k As Long
Dim strAllocatedFormula As String
Dim n As Name
Dim rngPlAmountUSD As Range

'find PL total allocatin range and save its Amount USD column as a range
    For Each n In wsAllocations.Names
        If GenericFunctions.StringSearch(1, n.Name, "Allocations_PL.Name_") > 0 Then
            Set rngPlAmountUSD = wsAllocations.Range(Range(n)(3, 3), Range(n)(wsAllocations.Range(n).Rows.Count, 3))
        End If
    Next n

'Loop cells in amount usd range
    For i = 0 To rngPlAmountUSD.Rows.Count - 1
        'Generate a new formula.  first build sum section that covers project ranges.  second divide it by pl total amount usd and wrap in iferror
            strAllocatedFormula = "("
            'build sum portion
                For j = 0 To UBound(arrVarProjectCountPerActivity, 1)
                    'add + between activity sections
                        If j > 0 Then strAllocatedFormula = strAllocatedFormula & "+"
                    'for each activity if project count = 0 then refer to the activity amount usd (fully allocated)
                    'else create a SUM() function that covers the range of projects in the row for that activity
                    'note:  for the .Offset(2 + i, 2) the 2 + is the offset from the top of the range to first value row
                        If arrVarProjectCountPerActivity(j, 1) = 0 Then
                            strAllocatedFormula = strAllocatedFormula & Range(arrVarProjectCountPerActivity(j, 0))(1, 1).Offset(2 + i, 2).Address(False, False)
                        Else
                            strAllocatedFormula = strAllocatedFormula & "SUM(" & Range(arrVarProjectCountPerActivity(j, 0))(1, 1).Offset(2 + i, 5).Address(False, False) & ":" & _
                                Range(arrVarProjectCountPerActivity(j, 0))(1, 1).Offset(2 + i, 4 + (arrVarProjectCountPerActivity(j, 1))).Address(False, False) & ")"
                        End If
                Next j
                
            'build divide by pl total amount usd and error wrap
                strAllocatedFormula = "=IFERROR(" & strAllocatedFormula & ")/" & rngPlAmountUSD(1, 1).Offset(i, 0).Address(False, False) & ","""")"
       
    
        'set cell format and formula
            PAFCellFormats.FormatAllocatedPercCell rngPlAmountUSD(1, 1).Offset(i, 1)
            rngPlAmountUSD(1, 1).Offset(i, 1).Formula = strAllocatedFormula
    Next i

'Set cf and white borders
    PAFCellFormats.FormatAllBordersWhiteThin rngPlAmountUSD.Offset(0, 1)
    SetAllocatedCf rngPlAmountUSD.Offset(0, 1)

End Sub


'@Description "generates and applies allocated% formula and sets formatting for an activity range"
Private Sub WriteAcitivityAllocationFormulaAndFormat( _
                                                ByRef rngAllocated As Range, _
                                                ByVal intProjectCount As Integer)
  
Dim rngCell As Range
  
'For each cell in the allocated% range build the formula based on the project count and row number
    For Each rngCell In rngAllocated
        With rngCell
            'apply formatting
                PAFCellFormats.FormatAllocatedPercCell rngCell
            'if project count >= 1 then formula covers range of project columns in activity
            'if project count = 0 then formula references the amount usd of the activity
            If intProjectCount >= 1 Then
                .FormulaR1C1 = "=IFERROR(SUM(R[0]C[2]:R[0]C[" & intProjectCount + 1 & "]) / SUM(R[0]C[-1]),"""")"
            Else
                .FormulaR1C1 = "=IFERROR(SUM(R[0]C[-1]) / SUM(R[0]C[-1]),"""")"
            End If
        End With
    Next rngCell

'Call sub to set Cf for column
    PAFCellFormats.FormatAllBordersWhiteThin rngAllocated
    SetAllocatedCf rngAllocated
    
End Sub



'@Description "searches range names for a string and returns a count of occurences"
Private Function CountActivityRanges( _
                                        ByRef wsAllocations As Worksheet, _
                                        ByVal strSearchString As String) _
                                        As Integer

Dim i           As Integer  'Counter for number of times the string is found
Dim nWsRange    As Name     'A Name object in the Worksheet

For Each nWsRange In wsAllocations.Names
    If GenericFunctions.StringSearch(1, nWsRange.Name, strSearchString) > 0 Then i = i + 1
Next nWsRange

CountActivityRanges = i

End Function


'@Description "Sets the cf for the allocated% column

Private Sub SetAllocatedCf( _
                            rangeTarget As Range)

'Clear any existing conditional formatting
    rangeTarget.FormatConditions.Delete

'Add/format data bar based on 0% to 100% range
    rangeTarget.FormatConditions.AddDatabar

    With rangeTarget.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With

'Add/Format an icon set that shows a check if value is exactly 100% and a 'x' if it's over 100%
'and nothing if below 100%
    rangeTarget.FormatConditions.AddIconSetCondition
    
    With rangeTarget.FormatConditions(2)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Symbols2)
    End With
    
    rangeTarget.FormatConditions(2).IconCriteria(1).Icon = xlIconNoCellIcon
    
    With rangeTarget.FormatConditions(2).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 1
        .Operator = 7
        .Icon = xlIconGreenCheck
    End With
    
    With rangeTarget.FormatConditions(2).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 1
        .Operator = 5
        .Icon = xlIconRedCross
    End With

End Sub
