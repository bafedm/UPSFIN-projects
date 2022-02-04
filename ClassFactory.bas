Attribute VB_Name = "ClassFactory"
'@Folder("Class Objects.Factory")
'@Description "Contains constructors for class objects"
Option Explicit

Public Function newProjectObject( _
                                    ByRef objParentPl As clsPandL, _
                                    ByVal dtReportMonth As Date, _
                                    ByRef objParentActivity As clsActivity, _
                                    ByVal strName As String, _
                                    ByVal strDescription As String, _
                                    ByVal dtStartDate As Date, _
                                    ByVal dtEndDate As Date _
                                    ) As clsProject

With New clsProject
    .objParentPl = objParentPl
    .dtReportMonth = dtReportMonth
    .objParentActivity = objParentActivity
    .strName = strName
    .strDescription = strDescription
    .dtStartDate = dtStartDate
    .dtEndDate = dtEndDate
    
    Set newProjectObject = .Self
End With



End Function


'@Description "Creates a new PL Object without the child Pl data"
Public Function NewPlObject( _
                                ByVal strName As String _
                                ) As clsPandL
                                
With New clsPandL

    .strName = strName
    .boolHasChildren = False    'default value
    
    Set NewPlObject = .Self
    
End With

End Function


Public Function NewActivityObject( _
                                    ByVal strName As String _
                                    ) As clsActivity
With New clsActivity

    .strName = strName
    .collParentPl = New Collection
    .dictFinanceDataTable = New Scripting.Dictionary
    .dictFinanceDataTableHeader = New Scripting.Dictionary
    
    Set NewActivityObject = .Self

End With

End Function

