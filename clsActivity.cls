VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class Objects.Class Activity")
'@Description "Interface for Activity Objects"
Option Explicit

Private mStrName                        As String       'The Activity Name from the data model
Private mCollParentPl                   As Collection   'A collection containing P&L objects that contain finance data for the Activity in for the selected reporting period.
Private mDictFinanceDataTable           As Dictionary   'A dictionary object containing finace data tables for each P&L for the related activity.  Key=String of P&L Name, Item=Variant Array with finance data table
Private mDictFinanceDataTableHeader     As Dictionary   'A dictionary object containing finace data table headers for each P&L for the related activity.  Key=String of P&L Name, Item=Variant Array of column headers



'@Description "Returns self for constructor"
Public Property Get Self() As clsActivity
    Set Self = Me
End Property

'@Description "Returns the activity name of the object"
Public Property Get strName() As String
    strName = mStrName
End Property

'@Description "Sets the activity name of the object"
Public Property Let strName(ByVal strNewString As String)
    mStrName = strNewString
End Property

'@Description "Returns a collection containing the P&L Objects that are related to the activity"
Public Property Get collParentPl() As Collection
    Set collParentPl = mCollParentPl
End Property

'@Description "Sets the parent P&L collection"
Public Property Let collParentPl(ByVal collParentPl As Collection)
    Set mCollParentPl = collParentPl
End Property

'@Description "Returns a dictionay collection containing finance data tables for each P&L related to the activity"
Public Property Get dictFinanceDataTable() As Scripting.Dictionary
    Set dictFinanceDataTable = mDictFinanceDataTable
End Property

'@Description "Sets a new dictionary collection containing finance data tables for each P&L related to the activity"
Public Property Let dictFinanceDataTable(ByVal dictNewFinanceTable As Scripting.Dictionary)
   Set mDictFinanceDataTable = dictNewFinanceTable
End Property


'@Description "Returns a dictionary collection containing finance data headers for each P&L releated to the activity"
Public Property Get dictFinanceDataTableHeader() As Scripting.Dictionary
    Set dictFinanceDataTableHeader = mDictFinanceDataTableHeader
End Property

'@Description "Sets a new dictionary collection containing finance data headers for each P&L releated to the activity"
Public Property Let dictFinanceDataTableHeader(ByVal dictNewFinanceTableHeader As Scripting.Dictionary)
    Set mDictFinanceDataTableHeader = dictNewFinanceTableHeader
End Property
