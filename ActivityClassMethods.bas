Attribute VB_Name = "ActivityClassMethods"
'@Folder("Class Objects.Class Activity")
Option Explicit

'@Description "Creates a collection of activity objects with data from the data model"
Public Function GenerateActivityObjectCollection( _
                    wbUpsfin As Workbook, _
                    wsProjectWb As Worksheet, _
                    dtReportingPeriod As Date, _
                    collPls As Collection _
                    ) As Collection
                    
' x  Get list of all activities and associated p&ls
' x  create new clsActivity object for each activity in the table
' x  populate parent p&l list
'for all activity/p&l combinations get finance data
'save finance data and headers to activity dictionaries

'Get list of all activities and associated p&ls
    Dim arrVarActivityTable     As Variant  '2d array that holds the activity and p&l association table.  col_0 = activity, col_1 = p&l
    arrVarActivityTable = GetActivityAndPlTableFromDataModel(wbUpsfin, dtReportingPeriod)

'create new clsActivity object for each activity in the table
    Dim collActivities          As New Collection 'holds activities and returned to caller
    Set collActivities = GenerateInitialActivityList(arrVarActivityTable)
    
'for each activity generate it's parent p&l collection
    GenerateParentPlCollections collActivities, collPls, arrVarActivityTable
    
'for all activity/p&l combinations get finance data
    PopulateActivityFinanceTables wbUpsfin, collActivities, dtReportingPeriod

Set GenerateActivityObjectCollection = collActivities

    
End Function

'@Description "Adds dictionary items containing the finance data and headers for each activity/p&l"
Private Sub PopulateActivityFinanceTables( _
                                            wbUpsfin As Workbook, _
                                            collActivities As Collection, _
                                            dtReportingPeriod As Date)
            
'for each activity
'for each activity p&l get finance data from dm
'with finance table header create index with desired columns
'with finance table write array for each row using column index
'save header/table to appropriate dictionary items, key = p&l name

Dim i As Long, j As Long, k As Long
Dim objActivity                 As clsActivity
Dim objPl                       As clsPandL
Dim dictDmData                  As Scripting.Dictionary
Dim arrVarHeaderNamesAndIndex   As Variant
Dim arrVarDataTableClean()      As Variant

'Main activity loop
    For Each objActivity In collActivities
         
        'P&L loop for each activity
            For Each objPl In objActivity.collParentPl
            
                'call function to get table data as dictionary object
                    Set dictDmData = New Scripting.Dictionary
                    Set dictDmData = GetFinanceDataHeaderAndTableFromDm(wbUpsfin, dtReportingPeriod, objActivity.strName, objPl.strName)
                    
                'check if returned dictionary contains data.  If true then format and assign to object
                    If IsArray(dictDmData("header")) Then
                        
                    'build index array of desired columns.  result is 2d array with 0=column names, 1=column index
                        arrVarHeaderNamesAndIndex = GetMemberValueAndDateColumns(dictDmData("header"))
                    
                    'build new table array based on column indexes
                        arrVarDataTableClean = GetFinanceTableWithSelectedColumns(dictDmData("table"), arrVarHeaderNamesAndIndex)
                        
                    'Store clean header and data table to activity dictionary items
                        objActivity.dictFinanceDataTable.Add Key:=objPl.strName, Item:=arrVarDataTableClean
                        objActivity.dictFinanceDataTableHeader.Add Key:=objPl.strName, Item:=arrVarHeaderNamesAndIndex(0)
                    End If
            
            Next objPl
                   
    Next objActivity




End Sub

'@Desription "Copies selected columns from data model table into a new table using the column index"
Private Function GetFinanceTableWithSelectedColumns( _
                                                        arrVarDataModelTable As Variant, _
                                                        arrVarHeaderNamesAndIndex As Variant) _
                                                        As Variant

Dim i As Long, j As Long, k As Long
Dim arrVarDataTableClean()   As Variant

ReDim arrVarDataTableClean(LBound(arrVarDataModelTable) To UBound(arrVarDataModelTable), _
                            LBound(arrVarHeaderNamesAndIndex(1)) To UBound(arrVarHeaderNamesAndIndex(1)))

For i = LBound(arrVarDataModelTable, 1) To UBound(arrVarDataModelTable, 1)
    For j = LBound(arrVarHeaderNamesAndIndex(1), 1) To UBound(arrVarHeaderNamesAndIndex(1), 1)
        arrVarDataTableClean(i, j) = arrVarDataModelTable(i, arrVarHeaderNamesAndIndex(1)(j))
    Next j
Next i

GetFinanceTableWithSelectedColumns = arrVarDataTableClean

End Function


'@Description "Returns a 2D array containing the names and column indexes of [MEMBER_VALUE] and date columns"
Private Function GetMemberValueAndDateColumns( _
                                                arrVarRawHeaderNames) As Variant

'search columns for MEMBER_VALUE or MMM-YYYY string
'store column index for each match
'extract clean and store column name from each match
'return as 2d array 1=column name, 2=column index

Dim i As Long, j As Long, k As Long
Dim arrVarColumnNames()     As Variant
Dim arrVarColumnIndexes()   As Integer

j = 0
'Loop through all items in array, search for member_value string or mmm-yyyy
For i = LBound(arrVarRawHeaderNames, 1) To UBound(arrVarRawHeaderNames, 1)
    
    If StringSearch(1, arrVarRawHeaderNames(i), "[MEMBER_VALUE]") > 0 Then
        'If string is found resize arrays, store index, store cleaned column name, increment index counter
        'example raw text [tbl_d_DescGroup].[Desc_Group].[Desc_Group].[MEMBER_VALUE]
            ReDim Preserve arrVarColumnIndexes(j)
            ReDim Preserve arrVarColumnNames(j)
            arrVarColumnIndexes(j) = i
            arrVarColumnNames(j) = CleanMdxString(arrVarRawHeaderNames(i), 2, Array("[", "]", "&"), True)
            j = j + 1
    
    ElseIf StringSearch(1, arrVarRawHeaderNames(i), "[MMM-YYYY]") Then
        'If string is found resize arrays, store index, store cleaned column name as date, increment counter
        'Example raw text [d_Cal_accPeriod].[MMM-YYYY].&[Jan-2020]
        ReDim Preserve arrVarColumnIndexes(j)
        ReDim Preserve arrVarColumnNames(j)
        arrVarColumnIndexes(j) = i
        arrVarColumnNames(j) = CleanMdxString(arrVarRawHeaderNames(i), 0, Array("[", "]", "&"), True)
        j = j + 1
    
    End If
Next i

'return as 2d array
    GetMemberValueAndDateColumns = Array(arrVarColumnNames, arrVarColumnIndexes)

End Function

'@Description "Returns the header and table for an activity/pl set"
Private Function GetFinanceDataHeaderAndTableFromDm( _
                                                        wbUpsfin As Workbook, _
                                                        dtReportingPeriod As Date, _
                                                        strActivityName As String, _
                                                        strPlName As String)


Dim strMdxPath                  As String                       'Mdx path for data model request

'Pivot table setup
'Page Filters: [d_Cal_reportMonth].[MMM-YY], [co_d_PL_toTCodeFilter].[PL_Name], [d_tbl_tCodeNamesActivity].[Activity_Name]
'Columns: [d_Cal_accPeriod].[MMM-YYYY]
'Rows: [co_f_busUnitsAndProjects].[Project_Name], [dm_tbl_revCostRelationship].[Rev_Cost], [tbl_d_DescGroup].[Desc_Group], [co_f_busUnitsAndProjects].[Description]
'Values: [Measures].[(PROJ) BU Upstream P&Ls Amount USD]

'Set Mdx path
    strMdxPath = "SELECT NON EMPTY Hierarchize({[dm_d_AccountingPeriod_Calendar].[MMM-YYYY].[MMM-YYYY].AllMembers}) " & _
            "DIMENSION PROPERTIES PARENT_UNIQUE_NAME,MEMBER_VALUE,HIERARCHY_UNIQUE_NAME ON COLUMNS , NON EMPTY " & _
            "Hierarchize(CrossJoin({[q_dm_BU_CY_PY].[Project_Name].[Project_Name].AllMembers}," & _
            "{([tbl_d_AC_DescGroupRanges].[RevCost_Groups].[RevCost_Groups].AllMembers," & _
            "[tbl_d_AC_DescGroupRanges].[Desc_Groups].[Desc_Groups].AllMembers,[q_dm_BU_CY_PY].[Description].[Description].AllMembers)})) " & _
            "DIMENSION PROPERTIES PARENT_UNIQUE_NAME,MEMBER_VALUE,HIERARCHY_UNIQUE_NAME ON ROWS  FROM [Model] WHERE " & _
            "([dm_d_ReportingPeriod_Calendar].[MMM-YYYY].&[" & _
            Format(dtReportingPeriod, "MMM-YYYY") & _
            "],[q_co_PlRanges].[PL_Name].&[" & _
            strPlName & _
            "],[d_tbl_tCodeNamesActivity].[Activity_Name].&[" & _
            strActivityName & _
            "],[Measures].[(BU PL) Description Grouping Amount USD]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS"

'Call function to get table data and return to caller
    Set GetFinanceDataHeaderAndTableFromDm = GetTableDataFromDataModel(wbUpsfin, strMdxPath)

End Function



'@Description "Add Parent P&Ls to the Activity Object as a collection"
Private Sub GenerateParentPlCollections( _
                                            collActivities As Collection, _
                                            collPls As Collection, _
                                            arrVarActivityTable As Variant)

'for each activity check associated p&l to see if it is childless
'if p&l is childless add to activity parent p&l collection
'if no childless p&l is present add the lowest level p&ls

Dim i As Long, j As Long, k As Long

Dim objActivity     As clsActivity      'for looping through activity collection

For Each objActivity In collActivities

    'First loop
    'Check if each associated P&L is childless, if so it adds that P&L to the activity collection
        For i = LBound(arrVarActivityTable, 1) To UBound(arrVarActivityTable, 1)
            If arrVarActivityTable(i, 0) = objActivity.strName Then
                If collPls(arrVarActivityTable(i, 1)).boolHasChildren = False Then
                    objActivity.collParentPl.Add Key:=arrVarActivityTable(i, 1), Item:=collPls(arrVarActivityTable(i, 1))
                End If
            End If
        Next i
    
    'Second loop
    'If none of the P&Ls are childless then add any P&Ls of the highest hierarchial level (farthest from top)
    'Always add the first one to the empty collection
    'Compare the next one to first one
    '   if it has a higher value erase any existing and add current
    '   if it is equal then add it
        If objActivity.collParentPl.Count = 0 Then
            For i = LBound(arrVarActivityTable, 1) To UBound(arrVarActivityTable, 1)
                
                If objActivity.collParentPl.Count = 0 Then 'initial entry
                    objActivity.collParentPl.Add Key:=arrVarActivityTable(i, 1), Item:=collPls(arrVarActivityTable(i, 1))
                    
                ElseIf collPls(arrVarActivityTable(i, 1)).intPlLevel > objActivity.collParentPl(1).intPlLevel Then 'if greater
                    objActivity.collParentPl = New Collection
                    objActivity.collParentPl.Add Key:=arrVarActivityTable(i, 1), Item:=collPls(arrVarActivityTable(i, 1))
                    
                ElseIf collPls(arrVarActivityTable(i, 1)).intPlLevel = objActivity.collParentPl(1).intPlLevel Then 'if equal
                    'HasKey checks if the key is already in use.  If not (returns true) then add the P&L to the collection
                        If GenericFunctions.HasKey(objActivity.collParentPl, CStr(arrVarActivityTable(i, 1))) Then objActivity.collParentPl.Add Key:=arrVarActivityTable(i, 1), Item:=collPls(arrVarActivityTable(i, 1))
    
                End If
                
            Next i
        End If
    
    'If the Activity doesn't have any Parent P&L at this point something is wrong.  Prompt and abort execution
        If objActivity.collParentPl.Count = 0 Then
            MsgBox "Error: " & objActivity.strName & " does not have any Parent P&L, halting program"
            End
        End If
                    
Next objActivity

End Sub



'Description "Takes array of activities and generates a collection of activity objects"
Private Function GenerateInitialActivityList( _
                                                arrVarActivityTable As Variant _
                                                ) As Collection
                                                
Dim i As Long, j As Long, k As Long
Dim collActivites   As New Collection   'Holds list of activity objects
Dim objActivity     As clsActivity      'for looping through object collection
                                                
                                                
'loop though array activity list.  If the collection is empty add the create the first item as a new activity and add
'for each item after the first check the collection to see if it is already present, if not create and add
                                                
    For i = LBound(arrVarActivityTable, 1) To UBound(arrVarActivityTable, 1)
        If collActivites.Count = 0 Then
            collActivites.Add Key:=arrVarActivityTable(i, 0), Item:=ClassFactory.NewActivityObject(arrVarActivityTable(i, 0))
        Else
            j = 0 'reset counter
            For Each objActivity In collActivites
                If objActivity.strName = arrVarActivityTable(i, 0) Then j = j + 1 'if object is already present increment counter
            Next objActivity
            
            If j = 0 Then collActivites.Add Key:=arrVarActivityTable(i, 0), _
                                            Item:=ClassFactory.NewActivityObject(arrVarActivityTable(i, 0)) 'if counter is 0 then it means no objects found with that name
        End If
    Next i
    
'Set collection to function
    Set GenerateInitialActivityList = collActivites

End Function


'@Description "Returns a table with the activities and associated P&Ls"
Private Function GetActivityAndPlTableFromDataModel( _
                                                        wbUpsfin As Workbook, _
                                                        dtReportingPeriod As Date _
                                                        ) As Variant


Dim i As Long, j As Long, k As Long                             'counters
Dim dictDmReturnData            As New Scripting.Dictionary     'Returns header and table from data model
Dim strMdxPath                  As String                       'Mdx path for data model request
Dim arrVarTableDataFromDm()   As Variant                        'return array, 2d. 0 = Activity List, 1 = P&L list

'Pivot table setup
'page filters:  [d_Cal_reportMonth].[MMM-YY], [d_Cal_accPeriod].[MMM-YYYY]
'rows:  [d_tbl_tCodeNamesActivity].[Activity_Name], [co_d_PL_toTCodeFilter].[PL_Name]
'Values: [Measures].[(PROJ) BU Upstream P&Ls Amount USD]

'Set MDX
    strMdxPath = "SELECT NON EMPTY Hierarchize(CrossJoin({[d_tbl_tCodeNamesActivity].[Activity_Name].[Activity_Name].AllMembers}, " & _
        "{([q_co_PlRanges].[PL_Name].[PL_Name].AllMembers)})) DIMENSION PROPERTIES PARENT_UNIQUE_NAME,MEMBER_VALUE,HIERARCHY_UNIQUE_NAME ON COLUMNS  " & _
        "FROM [Model] WHERE ([dm_d_ReportingPeriod_Calendar].[MMM-YYYY].&[" & _
        Format(dtReportingPeriod, "MMM-YYYY") & _
        "],[dm_d_AccountingPeriod_Calendar].[MMM-YYYY].&[" & _
        Format(dtReportingPeriod, "MMM-YYYY") & _
        "],[Measures].[(BU PL) Description Grouping Amount USD]) CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE, BACK_COLOR, FORE_COLOR, FONT_FLAGS"
               
'Call function to return data from data model, store in dictionary
    Set dictDmReturnData = GetTableDataFromDataModel(wbUpsfin, strMdxPath)

'Store dictionary header into return array
'header contains both activity and p&l data.  Each needs to be isolated and stored separately
'header table activity name is position 3 from front, p&l name is position 0 from back

    'clean and store activity list
        ReDim arrVarTableDataFromDm(0 To UBound(dictDmReturnData("header")), 0 To 1)
        
        For i = LBound(dictDmReturnData("header")) To UBound(dictDmReturnData("header"))
            arrVarTableDataFromDm(i, 0) = GenericFunctions.CleanMdxString(dictDmReturnData("header")(i), 3, Array("[", "]", "&"))
            arrVarTableDataFromDm(i, 1) = GenericFunctions.CleanMdxString(dictDmReturnData("header")(i), 0, Array("[", "]", "&"), True)
        Next i
    
'set array to function and return
    GetActivityAndPlTableFromDataModel = arrVarTableDataFromDm

End Function
