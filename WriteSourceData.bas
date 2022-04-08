Attribute VB_Name = "WriteSourceData"
'@Folder("Generate PAF Workbook.WS Source Data")
Option Explicit

'@Description "Writes project details to the related table"
Sub WriteProjectTable( _
                        wbPaf As Workbook, _
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
Dim wsSourceData As Worksheet
Dim lsoProjectTable As ListObject
Dim objProject As clsProject

'Set worksheet and project table to variables
    Set wsSourceData = wbPaf.Worksheets("Source Data")
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
