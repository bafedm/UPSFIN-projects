VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPandL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Class Objects.Class PL")
'@Description "P&L Object"
Option Explicit

Private mStrName                As String       'Object Name
Private mCollChildrenPl         As Collection   'Collection of P&L Names byRef
Private mBoolHasChildren        As Boolean      'True/False if Children present, default false
Private mIntPlLevel             As Integer      'Indicates the P&L hierarchy level (1 = highest)

'@Description "Returns self for constructor"
Public Property Get Self() As clsPandL
    Set Self = Me
End Property

'@Description "Gets the P&L object name"
Public Property Get strName() As String
    strName = mStrName
End Property

'@Description "Gets the P&L object name"
Public Property Let strName(ByVal strNewName As String)
    mStrName = strNewName
End Property

'@Description "Returns an array containing associated children P&L object references"
Public Property Get collChildrenPl() As Collection
    Set collChildrenPl = mCollChildrenPl
End Property

'@Description "Returns an array containing associated children P&L object references"
Public Property Let collChildrenPl(ByRef collChildrenPl As Collection)
    Set mCollChildrenPl = collChildrenPl
End Property

'@Description "returns True if mArrVarChildrenPl is not empty"
Public Property Get boolHasChildren() As Variant
    boolHasChildren = mBoolHasChildren
End Property

'@Description "set True if mArrVarChildrenPl is not empty"
Public Property Let boolHasChildren(ByVal boolNewValue As Variant)
    mBoolHasChildren = boolNewValue
End Property

'@Description "Returns the P&L Hierarchy Level for the object"
Public Property Get intPlLevel() As Integer
    intPlLevel = mIntPlLevel
End Property

'@Description "Sets the P&L Hierarchy Level for the object"
Public Property Let intPlLevel(ByVal intNewPlLevel As Integer)
    mIntPlLevel = intNewPlLevel
End Property
