Attribute VB_Name = "WriteSourceData"
'@Folder("Generate PAF Workbook.WS Source Data")
Option Explicit

'@Description "Main loop for writing source data to the PAF workbook"
Sub Main( _
            wbPaf As Workbook, _
            collActivities As Collection, _
            collProjects As Collection, _
            objPl As clsPandL)
            
Dim wsSourceData As Worksheet

Set wsSourceData = wbPaf.Worksheets("Source Data")
            
            
'write activity table
    WriteActivitiesTable wsSourceData, collActivities, objPl

'write projects table
    WriteProjectTable wsSourceData, collProjects
    



End Sub

'@Description "Write pl activities to activity table"
Private Sub WriteActivitiesTable( _
                            wsSourceData As Worksheet, _
                            collActivities As Collection, _
                            objTargetPl As clsPandL)

Dim i As Long, j As Long, k As Long
Dim lsoActivityTable As ListObject
Dim objActivity As clsActivity
Dim objPl As clsPandL


Set lsoActivityTable = wsSourceData.ListObjects("tbl_srcActivityList")

For Each objActivity In collActivities
    'With lsoActivityTable
        
    For Each objPl In objActivity.collParentPl
    
        If objPl.strName = objTargetPl.strName Then
            
            lsoActivityTable.ListRows.Add
            i = lsoActivityTable.DataBodyRange.Rows.Count
            lsoActivityTable.DataBodyRange(i, 1).Value = objActivity.strName
            
        End If
        
    Next objPl
    
    'End With 'lsoActivityTable
Next objActivity

End Sub


'@Description "Writes project details to the related table"
Private Sub WriteProjectTable( _
                        wsSourceData As Worksheet, _
                        collProjects As Collection)

'with project collection
'with project table
'clear table
'for each project object
'write each row to table
'next object

'Project Table Structure
'   1. Activity Name
'   2. Project Name
'   3. Project Description
'   4. Start Date
'   5. End Date

Dim i As Long, j As Long, k As Long
Dim lsoProjectTable As ListObject
Dim objProject As clsProject

'Set project table to variable
    Set lsoProjectTable = wsSourceData.ListObjects("tbl_srcProjectList")

'Clear contents just to be safe
    'lsoProjectTable.DataBodyRange.Rows.Delete

For Each objProject In collProjects
    With lsoProjectTable
        .ListRows.Add
        
        i = .DataBodyRange.Rows.Count
        
        .DataBodyRange(i, 1).Value = objProject.objParentActivity.strName
        .DataBodyRange(i, 2).Value = objProject.strName
        .DataBodyRange(i, 3).Value = objProject.strDescription
        .DataBodyRange(i, 4).Value = objProject.dtStartDate
        .DataBodyRange(i, 5).Value = objProject.dtEndDate
    
    End With 'lsoProjectTable

Next objProject

End Sub


'@Description "Writes the Pl Lc subtotal data to a table in the source worksheet"
Public Sub WritePlLcData( _
            arrVarPlTotalsByProject As Variant, _
            wbPaf As Workbook)

Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim strMonthRevCost As String
Dim strRevCost As String
Dim lsoLcPlData As ListObject
Dim wsSourceData As Worksheet

Set wsSourceData = wbPaf.Worksheets("Source Data")

Set lsoLcPlData = wsSourceData.ListObjects("tbl_srcLcValues")

With lsoLcPlData
    For i = 0 To UBound(arrVarPlTotalsByProject, 1)
        For j = 0 To UBound(arrVarPlTotalsByProject(i, 1), 1)
            For k = 1 To UBound(arrVarPlTotalsByProject(i, 1)(j, 1), 1)
                For m = 0 To 1
                    .ListRows.Add
                        If m = 0 Then strRevCost = "Rev" Else strRevCost = "Costs"
                            .DataBodyRange(.DataBodyRange.Rows.Count, 1).Value = arrVarPlTotalsByProject(i, 0)
                            .DataBodyRange(.DataBodyRange.Rows.Count, 2).Value = arrVarPlTotalsByProject(i, 1)(j, 0)
                            .DataBodyRange(.DataBodyRange.Rows.Count, 3).Value = MonthName(k, True)
                            .DataBodyRange(.DataBodyRange.Rows.Count, 4).Value = strRevCost
                            .DataBodyRange(.DataBodyRange.Rows.Count, 5).Value = arrVarPlTotalsByProject(i, 1)(j, 1)(k, m)
                            
                        
    '                If m = 0 Then strRevCost = "Rev" Else strRevCost = "Costs"
    '                strMonthRevCost = " Month: " & MonthName(k, True) & strRevCost & arrVarPlTotalsByProject(i, 1)(j, 1)(k, m)
    '                Debug.Print _
    '                    "Activity: " & arrVarPlTotalsByProject(i, 0) & _
    '                    " Project: " & arrVarPlTotalsByProject(i, 1)(j, 0) & _
    '                    strMonthRevCost
                Next m
            Next k
        Next j
    Next i
End With 'lsoLcPlData

End Sub
