Attribute VB_Name = "DataLoad"
'@Folder("DataLoad")
Option Explicit

'@Description "Generates object lists and loads data from data model and tables"
Sub Main( _
            wbUpsfin As Workbook, _
            wsProjectWb As Worksheet, _
            dtReportingPeriod As Date, _
            collPls As Collection, _
            collActivities As Collection, _
            collProjects As Collection _
            )
            

'Get list of P&Ls from data model and generate objects, combine with P&L table data
Set collPls = PlClassMethods.GeneratePlObjectCollection(wbUpsfin, wsProjectWb, dtReportingPeriod)

'Get list of activities from data model and generate objects
Set collActivities = ActivityClassMethods.GenerateActivityObjectCollection(wbUpsfin, wsProjectWb, dtReportingPeriod, collPls)

'Load table of Projects and details from data model, generate objects
Set collProjects = ProjectClassMethods.GenerateProjectObjectCollection(wsProjectWb, dtReportingPeriod, collActivities, collPls)


End Sub









