VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class Objects.Class Project")
'@Description "Holds details for each project that is loaded from the projects table in UPSFIN workbook"
Option Explicit

Private mObjParentPl            As clsPandL   'Name of the parent P&L
Private mDtReportMonth          As Date     'Date the entry was added to the data model
Private mObjParentActivity      As clsActivity   'Name of the parent Activity
Private mStrName                As String   'Project Name
Private mStrDescription         As String   'Project Description
Private mDtStartDate            As Date     'Project Start Date
Private mDtEndDate              As Date     'Project End Date

Public Property Get Self() As clsProject
    Set Self = Me
End Property

'@Description "Returns the Parent Pl Name"
Public Property Get objParentPl() As clsPandL
    Set objParentPl = mObjParentPl
End Property

'@Description "Sets the Parent Pl Name"
Public Property Let objParentPl(ByRef strObjParentPl As clsPandL)
    Set mObjParentPl = strObjParentPl
End Property

'@Description "Returns the source report month"
Public Property Get dtReportMonth() As Date
    dtReportMonth = mDtReportMonth
End Property

'@Description "Sets the source report month"
Public Property Let dtReportMonth(ByVal dtNewReportMonth As Date)
    mDtReportMonth = dtNewReportMonth
End Property

'@Description "Returns the parent activity name"
Public Property Get objParentActivity() As clsActivity
    Set objParentActivity = mObjParentActivity
End Property

'@Description "Sets the parent activity name"
Public Property Let objParentActivity(ByRef strObjParentActivity As clsActivity)
    Set mObjParentActivity = strObjParentActivity
End Property

'@Description "Returns the name of the project"
Public Property Get strName() As String
    strName = mStrName
End Property

'@Description "Sets the object name"
Public Property Let strName(ByVal strNewName As String)
    mStrName = strNewName
End Property

'@Description "Returns the project description"
Public Property Get strDescription() As String
    strDescription = mStrDescription
End Property

'@Description "Sets the objects description"
Public Property Let strDescription(ByVal strNewDescription As String)
    mStrDescription = strNewDescription
End Property

'@Description "Returns the project start date"
Public Property Get dtStartDate() As Date
    dtStartDate = mDtStartDate
End Property

'@Description "Sets the project start date"
Public Property Let dtStartDate(ByVal dtNewStartDate As Date)
    mDtStartDate = dtNewStartDate
End Property


'@Description "Returns the project end date"
Public Property Get dtEndDate() As Date
    dtEndDate = mDtEndDate
End Property

'@Description "Sets the project end date"
Public Property Let dtEndDate(ByVal dtNewEndDate As Date)
    mDtEndDate = dtNewEndDate
End Property
