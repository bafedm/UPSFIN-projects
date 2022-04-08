Attribute VB_Name = "PAFCellFormats"
'@Folder("Generate PAF Workbook")
Option Explicit

Sub FromatInputCell( _
                        rngTarget As Range)
                        
With rngTarget
    .Interior.Color = RGB(204, 244, 255)
End With
    
End Sub

Sub FormatAllocatedPercCell( _
                                rngTarget As Range)
      
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = vbBlack
    .WrapText = True
    .VerticalAlignment = xlCenter
    .NumberFormat = "0%"
    .HorizontalAlignment = xlCenter
End With

End Sub


Sub FormatDescGroupHeader( _
                            rngTarget As Range)
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = RGB(237, 125, 49)
End With
                            
End Sub

'@Description "Sets the cell formating for project list activities header title
Sub FormatProjectListHeaderActivityTitle( _
                                        rngTarget As Range)

With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .Font.Color = RGB(237, 125, 49)
    .VerticalAlignment = xlCenter
End With

End Sub

Sub FormatAllocationsDGDListFontDefault( _
                                            rngTarget As Range)
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = vbBlack
    .VerticalAlignment = xlCenter
End With
                                            
End Sub


'@Description "sets a default cell font"
Sub FormatDefaultFont( _
                                        rngTarget As Range)
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = vbBlack
    .WrapText = True
    .VerticalAlignment = xlCenter
End With

End Sub

'@Description "sets the activity project table header cell font"
Sub FormatActivityListHeaderRow( _
                                    rngTarget As Range)

With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = vbBlack
    .Font.Bold = True
    '.Font.Underline = True
    '.HorizontalAlignment = xlCenter
End With

End Sub
'@Description "sets the activity project table header cell font"
Sub FormatProjectListHeaderRow( _
                                    rngTarget As Range)

With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = vbBlack
    .Font.Bold = True
    .Font.Underline = True
    '.HorizontalAlignment = xlCenter
End With

End Sub

'@Description "sets the activity project table header cell font"
Sub FormatNoValue( _
                    rngTarget As Range)

With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Italic = True
    .Font.Color = RGB(174, 170, 170)
End With

End Sub

'@Description "sets the cell borders for each cell in range"
Sub FormatAllBordersWhiteThin( _
                                ByRef rngTarget As Range)

Dim rngCells As Range

For Each rngCells In rngTarget
    rngCells.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(255, 255, 255)
Next rngCells

End Sub

'@Description "sets the cell borders for each cell in range"
Sub FormatBordersBlackBottomThinWhiteThin( _
                                ByRef rngTarget As Range)

Dim rngCells As Range

For Each rngCells In rngTarget
    rngCells.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(255, 255, 255)
    With rngCells.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
Next rngCells

End Sub
'@Description "sets format for hidden cells"
Sub FormatHiddenText( _
                        ByRef rngTarget As Range)
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Color = RGB(174, 170, 170)
End With

End Sub
                        
'@Description "sets formatting for cells containing dollar figures"
Sub FormatAmountUsd( _
                        ByRef rngTarget As Range, _
                        Optional intFontSize As Integer = 10)
                                            
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = intFontSize
    .Font.Color = vbBlack
    .NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "" - ""??_-;_-@_-"
End With

PAFCellFormats.FormatAllBordersWhiteThin rngTarget

End Sub

'@Description "set formatting for LC% cells"
Sub FormatLcPercentage( _
                        ByRef rngTarget As Range)
With rngTarget
    .ClearFormats
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Color = vbBlack
    .NumberFormat = "0.00%"
End With

PAFCellFormats.FormatAllBordersWhiteThin rngTarget
            

End Sub

'@Description "set formatting for LC Allocation Column Header Month"
Sub FormatLcMonthColumnHeader( _
                                ByRef rngTarget As Range)
                                
With rngTarget
    .ClearContents
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Color = vbBlack
    .HorizontalAlignment = xlCenter
    .NumberFormat = "MMM-YY"
End With

End Sub
                        
Sub FormatLcRowHeaderTitle( _
                            ByRef rngTarget As Range)
With rngTarget
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 90
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Color = vbBlack
End With


                            


End Sub
