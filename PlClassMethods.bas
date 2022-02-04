Attribute VB_Name = "PlClassMethods"
'@Folder("Class Objects.Class PL")
'@Description("Contains methods related to the PL Object")
Option Explicit

'@Description "Creates a collection P&L Objects with data from the data model and from a related table"
Function GeneratePlObjectCollection( _
                    wbUpsfin As Workbook, _
                    wsProjectWb As Worksheet, _
                    dtReportingPeriod As Date _
                    ) As Collection

Dim i As Long, j As Long, k As Long
Dim strTemp     As String

'Get P&L list from Data Model
    Dim arrVarPlNamesFromDm As Variant
    
    'Call function to return array of P&L names from the data model
        arrVarPlNamesFromDm = GetPlListFromDataModel(wbUpsfin, dtReportingPeriod)


'Get P&L table
    Dim lsoPlHierarchy                      As ListObject   'ws table object with P&L Data
    Dim arrVarPlHierarchyTableData          As Variant      'Data body range of the table
    Dim arrVarPlHierarchyTableHeader        As Variant      'header data range of the table
    Dim arrIntPlHierarchyTableHeaderIndex() As Integer      'column index of wanted columns
    
    'Set list object, get values from header/table
        Set lsoPlHierarchy = wsProjectWb.ListObjects("tblPlHierarchy")
        arrVarPlHierarchyTableData = lsoPlHierarchy.DataBodyRange
        arrVarPlHierarchyTableHeader = lsoPlHierarchy.HeaderRowRange
        
    'set column indexes
        ReDim arrIntPlHierarchyTableHeaderIndex(0 To 2)
        arrIntPlHierarchyTableHeaderIndex = GenericFunctions.GetHeaderColumnIndexes(arrVarPlHierarchyTableHeader, _
                                                    Array("PL_Name", "Child_PL", "PL_Level"), 2)
        
'Compare the dm list with the P&L table to insure that all P&L Hierarchy details are available
    CheckDmPlListAgainstPlTable arrVarPlNamesFromDm, arrVarPlHierarchyTableData, arrIntPlHierarchyTableHeaderIndex(0)
    
'Create new collection.  for each p&l from the dm create a new base P&L object
    Dim collTempPlList  As New Collection
    Dim varPlName       As Variant

    'loop p&l list, create p&l object and add to collection
        For Each varPlName In arrVarPlNamesFromDm
            collTempPlList.Add Key:=CStr(varPlName), Item:=ClassFactory.NewPlObject(CStr(varPlName))
        Next varPlName
        
'Loop through the P&L Object collection and populate child collections where required
    PopulateChildPlCollectionAndPlLevel collTempPlList, arrVarPlHierarchyTableData, arrIntPlHierarchyTableHeaderIndex

Set GeneratePlObjectCollection = collTempPlList
        
End Function

'@Description "Returns an array containing the active P&Ls from the data model based on the criteria"
Private Function GetPlListFromDataModel( _
                                            wbUpsfin As Workbook, _
                                            dtReportingPeriod As Date _
                                            ) As Variant
                                            
Dim i As Long
Dim dictDmReturnData    As New Scripting.Dictionary
Dim strMdxPath          As String
Dim arrVarPlNamesFromDm As Variant

'Pivot table setup
'Page filters:  [d_Cal_accPeriod].[MMM-YYYY], [d_Cal_reportMonth].[MMM-YY]
'Row items: [co_d_PL_toTCodeFilter].[PL_Name]
'Values items: [Measures].[(PROJ) BU Upstream P&Ls Amount USD]

'Set MDX
    strMdxPath = _
        "SELECT NON EMPTY Hierarchize({[co_d_PL_toTCodeFilter].[PL_Name].[PL_Name].AllMembers}) " & _
        "DIMENSION PROPERTIES PARENT_UNIQUE_NAME,MEMBER_VALUE,HIERARCHY_UNIQUE_NAME ON COLUMNS  " & _
        "FROM [Model] WHERE ([d_Cal_reportMonth].[MMM-YY].&[" & _
        Format(dtReportingPeriod, "MMM-YY") & _
        "],[d_Cal_accPeriod].[MMM-YYYY].&[" & _
        Format(dtReportingPeriod, "MMM-YYYY") & _
        "],[Measures].[(PROJ) BU Upstream P&Ls Amount USD]) " & _
        "CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS"

'Call function to return data from data model, store in dictionary
    Set dictDmReturnData = GetTableDataFromDataModel(wbUpsfin, strMdxPath)

'Only data in "header" element is needed, save to array for processing
    arrVarPlNamesFromDm = dictDmReturnData("header")

'Clean array elements by removing MDX formatting
    For i = 0 To UBound(arrVarPlNamesFromDm)
        arrVarPlNamesFromDm(i) = GenericFunctions.CleanMdxString(arrVarPlNamesFromDm(i), 0, Array("[", "]", "&"), True)
    Next i

'Set function to cleaned array
    GetPlListFromDataModel = arrVarPlNamesFromDm

End Function

'@Description "Compares the P&L list from the data model against the P&L table"
'If there are items on the list that don't match the table then prompt and exit program
Private Sub CheckDmPlListAgainstPlTable( _
                                            arrVarPlNamesFromDm As Variant, _
                                            arrVarPlHierarchyTableData As Variant, _
                                            intPlHierarchyPlNameColumnIndex As Integer)
    
Dim i As Long, j As Long, k As Long

'Local variables
    Dim collStrNoMatchList  As New Collection   'Holds list of P&L names that have no match
    Dim varCollItem         As Variant          'For incrementing through no match collection
    Dim strTemp             As String           'for building string to be used in msgbox

'Loop through each name in the dm P&L list, then compare it to all P&L names in the table.
'If match not found then add it to the no match collection
    For i = 0 To UBound(arrVarPlNamesFromDm)
        j = 0 'j indicates that a match has been found, if j = 0 after the 'k' loop then there are no matches
        For k = 1 To UBound(arrVarPlHierarchyTableData, 1)
            If arrVarPlNamesFromDm(i) = arrVarPlHierarchyTableData(k, intPlHierarchyPlNameColumnIndex) Then j = j + 1
        Next k
        If j = 0 Then collStrNoMatchList.Add arrVarPlNamesFromDm(i) 'when no match add to no match collection
    Next i

'If no match count is greater than 0 prompt user of missing P&Ls and exit program
    If collStrNoMatchList.Count > 0 Then
        
        strTemp = vbCrLf
        
        For Each varCollItem In collStrNoMatchList
            strTemp = strTemp & "   " & varCollItem & vbCrLf
        Next varCollItem
        
        MsgBox "The following P&Ls are missing from WB Generator Table: " & vbCrLf & strTemp & vbCrLf & "The WB Generator will now exit."
        
        End
    
    End If
    
    
End Sub

'@Description "Gets list of Child P&Ls from table and converts to collection of P&L objects to be assigned to parent P&L Object"
Private Sub PopulateChildPlCollectionAndPlLevel( _
                                collTempPlList As Collection, _
                                arrVarPlHierarchyTableData As Variant, _
                                arrIntPlHierarchyTableHeaderIndex As Variant)

Dim i As Long, j As Long, k As Long

'Local variables
    Dim objClsPandL         As clsPandL     'clsPandL object in the collTempPlList collection
    Dim arrStrChildPlName() As String       'Array formed from spliting of the table Child_PL row value
    Dim collChildPls        As Collection   'Collection generated of child PLs that is assigned to the P&L Object

'Loop through each P&L in the collection
'Get P&L name and compare it to the names on the P&L table
'When matched get the associated value from the Child_PL.  If value is "NONE" skip.  Also assign the PL_Level to the object
'Split the Child_PL by delimiter, store to array
'for each string in the Child_PL array find it's matching P&L Object and add it to the child P&L collection
'assign the collection to the P&L Object
    For Each objClsPandL In collTempPlList
    
        For i = 1 To UBound(arrVarPlHierarchyTableData, 1)
        
            If objClsPandL.strName = arrVarPlHierarchyTableData(i, arrIntPlHierarchyTableHeaderIndex(0)) Then 'compare object name vs table row name
                
                objClsPandL.intPlLevel = arrVarPlHierarchyTableData(i, arrIntPlHierarchyTableHeaderIndex(2)) 'assign PL level
                
                If arrVarPlHierarchyTableData(i, arrIntPlHierarchyTableHeaderIndex(1)) = "NONE" Then
                    GoTo noChild
                Else
                    Set collChildPls = New Collection 'new collection for each objects
                    
                    arrStrChildPlName = Split(arrVarPlHierarchyTableData(i, arrIntPlHierarchyTableHeaderIndex(1)), ",") 'split string
                    
                    'trim any leading/tailing spaces
                        For j = 0 To UBound(arrStrChildPlName)
                            arrStrChildPlName(j) = Trim(arrStrChildPlName(j))
                        Next j
                    
                    
                    'for each child P&L in string find the associated object and assign to collection
                        For j = 0 To UBound(arrStrChildPlName)
                            For k = 1 To collTempPlList.Count
                                If collTempPlList(k).strName = arrStrChildPlName(j) Then collChildPls.Add collTempPlList(k)
                            Next k
                        Next j
                    
                    'Assign collection to object
                        objClsPandL.collChildrenPl = collChildPls
                    
                    'Set HasChildren bool to True (constructor default = false)
                    objClsPandL.boolHasChildren = True
                    
                End If
noChild:
            End If
            
        Next i
        
    Next objClsPandL
    
End Sub
