Attribute VB_Name = "LcForecastPopWs"
'@Folder("Generate PAF Workbook.WS Lc Forecast")
Option Explicit

'@Description "Methods for writing to LC worksheet"
Sub WriteToLcWorksheet( _
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

End Sub

Private Sub WriteTablesToLcWorksheet( _
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
Dim wsLcForecast            As Worksheet
Dim rngTopAnchor            As Range
Dim intLcTableOffset        As Integer
Dim intPlTableOffset        As Integer
Dim intAnchorOffset         As Integer
Dim intMonthStartCol        As Integer
Dim intRowHeaderStartRow    As Integer

intLcTableOffset = 2
intPlTableOffset = 11
intMonthStartCol = 3
intRowHeaderStartRow = 3


intAnchorOffset = (intLcTableOffset * 2) + intPlTableOffset

'Set top anchor range
    Set rngTopAnchor = wsLcForecast.Range("Lc.Forecast_Top.Anchor")
    




End Sub

