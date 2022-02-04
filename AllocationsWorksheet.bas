Attribute VB_Name = "AllocationsWorksheet"
'@Folder("Generate PAF Workbook.WS Allocations")
'@Description "Populates P&L data for reporting period and allows for project allocation"
Option Explicit

'@Description "Main loop for allocations"
Sub PopulateAllocationsWorksheet( _
                                    ByRef wbPaf As Workbook, _
                                    ByVal dtReportingPeriod As Date, _
                                    ByRef objTargetPl As clsPandL, _
                                    ByRef collActivities As Collection)

'assign worksheets to variable
'combine all activities finance details into master array
'generate desc_group, desc list
'create array with p&l totals and write
'set named range for p&l total
'for each activity
'write sub totals
'get project list from ws and write project headers
'set named ranges for activity and projects
'next activity
'write allocation formulas and set cf

Dim wsProjectList               As Worksheet                    'Assign Project List ws to variable
Dim wsAllocations               As Worksheet                    'Assign Allocations ws to variable
Dim arrVarSortedPlFinanceList   As Variant                      'Sorted array of dictMasterPlFinanceTable by desc_group(custom) and desc(alpha)


Set wsProjectList = wbPaf.Worksheets("Project List")
Set wsAllocations = wbPaf.Worksheets("Allocations")

GenericFunctions.writeProjectHeadersToPafWorksheet wsAllocations, dtReportingPeriod, objTargetPl.strName, GenericFunctions.replaceIllegalNamedRangeCharacters(wsAllocations.Name)

Application.ScreenUpdating = False

'Generate a sorted array containing activity finance data for target P&L and reporting period
    arrVarSortedPlFinanceList = _
        AllocationsFinanceTableArray.GenerateMasterPlFinanceTableAsArray(dtReportingPeriod, objTargetPl, collActivities)
        
'Write activities and projects to worksheet
    AllocationsWriteToWs.Main wsAllocations, wsProjectList, objTargetPl, dtReportingPeriod, collActivities, arrVarSortedPlFinanceList

'Set cell formats for input, set allocated% formulas, set allocated% cf
    AllocationsAllocatedFormatting.Main wsAllocations

'Set project input cell formatting
    AllocationsProjectInputFormat.Main wsAllocations
    
Application.ScreenUpdating = True
End Sub




