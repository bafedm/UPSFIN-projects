Attribute VB_Name = "ProjectClassMethods"
'@Folder("Class Objects.Class Project")
Option Explicit

'@Description "Loads the projet table from UPSFIN into individual objects"
Public Function GenerateProjectObjectCollection( _
                                                    wsProjectWb As Worksheet, _
                                                    dtReportingPeriod As Date, _
                                                    collActivities As Collection, _
                                                    collPls As Collection _
                                                    ) As Collection


Dim i As Long, j As Long, k As Long
Dim collProjects                As Collection       'Projects collection, returns to caller
Dim lsoProjectsTable            As ListObject       'Project Table from worksheet as object
Dim rngProjectsTable            As Range            'Worksheet range of the projects table
Dim arrVarProjectDataBodyRange  As Variant          'Data body range from project table
Dim objProject                  As clsProject       'project object for each project that is assigned to collection

'Create new collProjects collection
    Set collProjects = New Collection

'Set table object
    Set lsoProjectsTable = wsProjectWb.ListObjects(TBL_PROJECT_LIST)

'Filter table to target reporting period, copy data body range to array
    With lsoProjectsTable
        .AutoFilter.ShowAllData
        'Set rngProjectsTable = lsoProjectsTable.Range
        'rngProjectsTable.AutoFilter Field:=1, Criteria1:=Format(dtReportingPeriod, "YYYY-MM-DD")
        '.Range.AutoFilter Field:=1, Criteria1:=Format(dtReportingPeriod, "YYYY-MM-DD")
        arrVarProjectDataBodyRange = .DataBodyRange
    End With

'Loop array, create object for each row, assign properties, add to collection
    For i = 1 To UBound(arrVarProjectDataBodyRange, 1)
       If Format(arrVarProjectDataBodyRange(i, 1), "YYYY-MM-DD") = Format(dtReportingPeriod, "YYYY-MM-DD") Then
           Set objProject = ClassFactory.newProjectObject( _
                                                        collPls(arrVarProjectDataBodyRange(i, 2)), _
                                                        arrVarProjectDataBodyRange(i, 1), _
                                                        collActivities(arrVarProjectDataBodyRange(i, 3)), _
                                                        arrVarProjectDataBodyRange(i, 4), _
                                                        arrVarProjectDataBodyRange(i, 5), _
                                                        arrVarProjectDataBodyRange(i, 6), _
                                                        arrVarProjectDataBodyRange(i, 7))
        
            
            collProjects.Add Key:=arrVarProjectDataBodyRange(i, 4), Item:=objProject
        End If
    Next i

'return collection to caller
    Set GenerateProjectObjectCollection = collProjects

End Function
