Attribute VB_Name = "ProjectsWorksheet"
'@Folder("Generate PAF Workbook.WS Projects")
'@Description "Methods related to the populating of the projects worksheet"
Option Explicit

'@Description "Main loop to populate the projects sheet"
Sub PopulateProjectsWorksheet( _
                                ByRef wbPaf As Workbook, _
                                ByVal dtReportingPeriod As Date, _
                                ByRef objTargetPl As clsPandL, _
                                ByRef collActivities As Collection, _
                                ByRef collProjects As Collection)

'Set main anchor range and offset variable
'Write pl/date headers
'loop activites
'write activity header
'write activity named range
'write projects
'next activity

Dim wsProjects          As Worksheet    'Assign projects worksheet to variable

Dim rngActivitiesAnchor As Range        'Main anchor for activity ranges, used as offset reference
Dim intActivitiesOffset As Integer      'Stores offset


Set wsProjects = wbPaf.Worksheets("Project List")
Set rngActivitiesAnchor = wsProjects.Range("Project.List_Activity.Top.Anchor")
intActivitiesOffset = 1

'Write headers
    GenericFunctions.writeProjectHeadersToPafWorksheet wsProjects, dtReportingPeriod, objTargetPl.strName, GenericFunctions.replaceIllegalNamedRangeCharacters(wsProjects.Name)
    
'Activity Loop
   WriteProjectTableToWorksheet wsProjects, objTargetPl, collActivities, collProjects, rngActivitiesAnchor, intActivitiesOffset
        
'set borders for tables
    SetBordersForActivityTable wsProjects

'Set date format and center no., date columns
    
    With wsProjects.Range("E1:F1").EntireColumn
        .HorizontalAlignment = xlCenter
        .NumberFormat = "DD-MMM-YYYY"
    End With
    
    wsProjects.Range(wsProjects.Range("Project.List_Activity.Top.Anchor"), wsProjects.Cells(wsProjects.Rows.Count, "B").End(xlUp)).HorizontalAlignment = xlCenter

    
End Sub

'@Description "Loops through all activities and projects and writes matching projects to the worksheet"
Private Sub WriteProjectTableToWorksheet( _
                                            ByRef wsProjects As Worksheet, _
                                            ByRef objTargetPl As clsPandL, _
                                            ByRef collActivities As Collection, _
                                            ByRef collProjects As Collection, _
                                            ByRef rngActivitiesAnchor As Range, _
                                            ByVal intActivitiesOffset As Integer)


Dim boolActivityWrite   As Boolean      'Indicates that an activity was written to ws and that offset should increment
Dim objPl               As clsPandL     'An individual p&L object from the collection
Dim objActivity         As clsActivity  'An individual activity object from the collection
Dim objProject          As clsProject   'An indvidual project object from the collection
Dim rngActivityAnchor   As Range        'local anchor
Dim intActivityOffset   As Integer      'local offset counter
Dim intProjectCounter   As Integer      'counts number of projects for activity

'loop all activities
For Each objActivity In collActivities
    
    'Reset activity write flag
        boolActivityWrite = False
        
    'loop all P&Ls that are in the activity collection
    For Each objPl In objActivity.collParentPl
            
        'If the current P&L has the same name as the target P&L then loop projects
        If objPl.strName = objTargetPl.strName Then
                
            'Set local anchor to sheet anchor offset
                Set rngActivityAnchor = rngActivitiesAnchor.Offset(intActivitiesOffset, 0)
                 
            'flag to write activity header, reset local offset, project counter (used for "No." column and to test if projects have been written)
                intActivityOffset = 0
                intProjectCounter = 0
                boolActivityWrite = True
                
            'activity header name
                WriteProjectTableHeader rngActivityAnchor, intActivityOffset, objActivity
                intActivityOffset = 2

            'loop projects
                For Each objProject In collProjects
                    'test if current project parent activity and parent p&l match against loop activity and parent p&l
                        If objProject.objParentActivity.strName = objActivity.strName And objProject.objParentPl.strName = objTargetPl.strName Then
                            'increment project counter, write row data, increment local offset
                                intProjectCounter = intProjectCounter + 1
                                WriteProjectRowData objProject, rngActivityAnchor, intActivityOffset, intProjectCounter
                                intActivityOffset = intActivityOffset + 1
                        End If
                Next objProject
                
            'if no project written activity then use no projects as default
                If intProjectCounter = 0 Then intActivityOffset = WriteNoProjectToTable(rngActivityAnchor, intActivityOffset)
                    
        End If 'objPl = objTargetPl
    Next objPl
        
    'if an activity was written increment the worksheet offset
        If boolActivityWrite = True Then
            SetActivityProjectTableRange wsProjects, rngActivitiesAnchor, intActivitiesOffset, intActivityOffset, objActivity.strName
            intActivitiesOffset = intActivitiesOffset + intActivityOffset + 2
        End If
    


Next objActivity
  
End Sub

'@Description "Sets white borders for each activity table"
Private Sub SetBordersForActivityTable( _
                                            ByRef wsProjects As Worksheet)

Dim varName As Variant

For Each varName In wsProjects.Names
    If StringSearch(1, varName.Name, "Project.List_Activity.Name_") Then PAFCellFormats.FormatAllBordersWhiteThin varName.RefersToRange
Next varName

End Sub


'@Description "Sets the named range for each activity table"
Private Sub SetActivityProjectTableRange( _
                                            ByRef wsProjects As Worksheet, _
                                            ByRef rngActivitiesAnchor As Range, _
                                            ByVal intActivitiesOffset As Integer, _
                                            ByVal intActivityOffset As Integer, _
                                            ByVal strActivityName As String)

'define range
'define range name
'set named range

Dim rngActivityRange        As Range
Dim strActivityRangeName    As String

Set rngActivityRange = Range(rngActivitiesAnchor.Offset(intActivitiesOffset, 0), rngActivitiesAnchor.Offset(intActivitiesOffset + intActivityOffset, 4))
strActivityRangeName = "Project.List_Activity.Name_" & replaceIllegalNamedRangeCharacters(strActivityName)

wsProjects.Names.Add Name:=strActivityRangeName, RefersTo:=rngActivityRange

End Sub

'@Description "Writes and formats no project"
Private Function WriteNoProjectToTable( _
                                    ByRef rngActivityAnchor As Range, _
                                    ByVal intActivityOffset As Integer) _
                                    As Integer
                                    
rngActivityAnchor.Offset(intActivityOffset, 1).Value = "no projects"
PAFCellFormats.FormatNoValue rngActivityAnchor.Offset(intActivityOffset, 1)
intActivityOffset = intActivityOffset + 1

WriteNoProjectToTable = intActivityOffset

End Function

'@Description "Writes the project header to the activity table"
Private Sub WriteProjectTableHeader( _
                                        ByRef rngActivityAnchor As Range, _
                                        ByVal intActivityOffset As Integer, _
                                        ByRef objActivity As clsActivity)

'Activity table layout
'0,0 Activity
'0,1 "Activity Name"
'1,0 No.
'1,1 Project Name
'1,2 Project Description
'1,3 Start Date
'1,4 End Date

With rngActivityAnchor
    .Offset(intActivityOffset, 0).Value = "Activity"
    PAFCellFormats.FormatProjectListHeaderActivityTitle .Offset(intActivityOffset, 0)
    
    .Offset(intActivityOffset, 1).Value = objActivity.strName
    PAFCellFormats.FormatDefaultFont .Offset(intActivityOffset, 1)
    
    .Offset(intActivityOffset + 1, 0).Value = "No."
    .Offset(intActivityOffset + 1, 1).Value = "Project Name"
    .Offset(intActivityOffset + 1, 2).Value = "Project Description"
    .Offset(intActivityOffset + 1, 3).Value = "Start Date"
    .Offset(intActivityOffset + 1, 4).Value = "End Date"
    PAFCellFormats.FormatProjectListHeaderRow Range(.Offset(intActivityOffset + 1, 0), .Offset(intActivityOffset + 1, 4))

End With

End Sub

'@Description "Writes the project details to a single row in the activity project table
Private Sub WriteProjectRowData( _
                                    ByRef objProject As clsProject, _
                                    ByRef rngActivityAnchor As Range, _
                                    ByVal intActivityOffset As Integer, _
                                    ByVal intProjectCounter)

With objProject
    rngActivityAnchor.Offset(intActivityOffset, 0).Value = intProjectCounter
    rngActivityAnchor.Offset(intActivityOffset, 1).Value = .strName
    rngActivityAnchor.Offset(intActivityOffset, 2).Value = .strDescription
    rngActivityAnchor.Offset(intActivityOffset, 3).Value = .dtStartDate
    rngActivityAnchor.Offset(intActivityOffset, 4).Value = .dtEndDate
    PAFCellFormats.FormatDefaultFont Range(rngActivityAnchor.Offset(intActivityOffset, 0), rngActivityAnchor.Offset(intActivityOffset, 4))
End With

End Sub
