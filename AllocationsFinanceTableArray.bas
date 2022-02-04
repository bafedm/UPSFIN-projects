Attribute VB_Name = "AllocationsFinanceTableArray"
'@Folder("Generate PAF Workbook.WS Allocations")
Option Explicit

'@Description "Creates a 2D array combining all project line items from the activity finance table."
Function GenerateActivityCombinedPlTableAsArray( _
                                                    ByRef objTargetPl As clsPandL, _
                                                    ByRef objActivity As clsActivity, _
                                                    ByVal dtReportingPeriod As Date) _
                                                    As Variant

'With an activity finance table and header table
'use header to find the reporting period
'create a dictonary to hold finance items.
'for each desc item (d) check to see if a dict key is present
'if key d exists is false then create a new dict with
'   key:=d, item:=array (0) rev cost, (1) desc group (2) desc (3) amount usd (target month)
'if key d exists is true then extract the array, accumulate amount USD, save array back to key
'transfer dictionary to array.  sorting not required.

Dim i As Long, j As Long, k As Long
Dim dictMasterActivityTable         As New Scripting.Dictionary
Dim arrVarOrginalActivityTable      As Variant
Dim intReportingMonthColumnIndex    As Integer
Dim arrVarTempRowData               As Variant

'get column index of reporting month from table header
For i = 0 To UBound(objActivity.dictFinanceDataTableHeader(objTargetPl.strName), 1)
    If objActivity.dictFinanceDataTableHeader(objTargetPl.strName)(i) = Format(dtReportingPeriod, "MMM-YYYY") Then intReportingMonthColumnIndex = i
Next i

'store dictionary finance table to array as easier to type...
    arrVarOrginalActivityTable = objActivity.dictFinanceDataTable(objTargetPl.strName)

'finance table loop
    For i = 0 To UBound(arrVarOrginalActivityTable, 1)
        If dictMasterActivityTable.Exists(arrVarOrginalActivityTable(i, 3)) Then
            If Not IsNull(arrVarOrginalActivityTable(i, intReportingMonthColumnIndex)) Or _
            Not arrVarOrginalActivityTable(i, intReportingMonthColumnIndex) = 0 Then
            
                arrVarTempRowData = dictMasterActivityTable(arrVarOrginalActivityTable(i, 3))
                arrVarTempRowData(3) = arrVarTempRowData(3) + arrVarOrginalActivityTable(i, intReportingMonthColumnIndex)
                dictMasterActivityTable(arrVarOrginalActivityTable(i, 3)) = arrVarTempRowData
                
            End If
        Else
            If Not IsNull(arrVarOrginalActivityTable(i, intReportingMonthColumnIndex)) Or _
            Not arrVarOrginalActivityTable(i, intReportingMonthColumnIndex) = 0 Then
                dictMasterActivityTable.Add Key:=arrVarOrginalActivityTable(i, 3), Item:=Array( _
                                                                                    arrVarOrginalActivityTable(i, 1), _
                                                                                    arrVarOrginalActivityTable(i, 2), _
                                                                                    arrVarOrginalActivityTable(i, 3), _
                                                                                    arrVarOrginalActivityTable(i, intReportingMonthColumnIndex))
            End If
        End If
    Next i
                                                                           
GenerateActivityCombinedPlTableAsArray = GenerateMasterDescListFromPlTable(dictMasterActivityTable)


End Function


'@Description "Returns an array containing the combined activity finance data for a given P&L/Reporting period
Function GenerateMasterPlFinanceTableAsArray( _
                                                dtReportingPeriod As Date, _
                                                objTargetPl As clsPandL, _
                                                collActivities As Collection) _
                                                As Variant



Dim dictMasterPlFinanceTable    As New Scripting.Dictionary     'Holds master Pl data key = desc, item = pl table array
                                                                '(0)rev_cost, (1)Desc_group, (2)Desc, (3)amountUSD
                                                                

Set dictMasterPlFinanceTable = GenerateMasterPlFinanceTable(dtReportingPeriod, objTargetPl, collActivities)
GenerateMasterPlFinanceTableAsArray = GenerateMasterDescListFromPlTable(dictMasterPlFinanceTable)

End Function

'@Description "Converts the combined P&L finance table dictionary into a sorted 2d array"
Private Function GenerateMasterDescListFromPlTable( _
                                                    ByRef dictMasterPlFinanceTable As Scripting.Dictionary) _
                                                    As Variant
Dim i As Long, j As Long, k As Long, m As Long
Dim varKey As Variant
Dim arrVarMasterList As Variant


'store dictionary items to 2d array
    ReDim arrVarMasterList(0 To dictMasterPlFinanceTable.Count - 1, 0 To UBound(dictMasterPlFinanceTable.Items(1), 1))
    
    For i = 0 To dictMasterPlFinanceTable.Count - 1
        For j = 0 To UBound(dictMasterPlFinanceTable.Items(i), 1)
            arrVarMasterList(i, j) = dictMasterPlFinanceTable.Items(i)(j)
        Next j
    Next i
    
'Change desc_group to numbers based on custom sort order
    Dim arrVarDescGroupOrder As Variant
    arrVarDescGroupOrder = Array("Revenue", "Personnel Costs", "External Services", "Travel & Vehicles", "Depreciation", _
                            "Operating Expense - 3rd Party", "Operating Expense - Group", "Split Overhead & Dir. & Ind. Costs")
    
    For i = 0 To UBound(arrVarMasterList, 1)
        For j = 0 To UBound(arrVarDescGroupOrder, 1)
            If arrVarMasterList(i, 1) = arrVarDescGroupOrder(j) Then arrVarMasterList(i, 1) = j
        Next j
    Next i
        
'Sort list alphabetically by desc
    GenericFunctions.QuickSortArray arrVarMasterList, 0, UBound(arrVarMasterList, 1), 2


'create new master array
'for each old master check if group number matches current, if so write and increment new master counter
    Dim arrVarSortedMaster As Variant
    ReDim arrVarSortedMaster(0 To UBound(arrVarMasterList, 1), 0 To 3)
    For i = 0 To UBound(arrVarDescGroupOrder, 1)
        For j = 0 To UBound(arrVarMasterList, 1)
            If arrVarMasterList(j, 1) = i Then
                For k = 0 To 3
                    arrVarSortedMaster(m, k) = arrVarMasterList(j, k)
                Next k
                m = m + 1
            End If
        Next j
    Next i
    
'convert numeric desc_group back to strings
    For i = 0 To UBound(arrVarSortedMaster, 1)
        arrVarSortedMaster(i, 1) = arrVarDescGroupOrder(arrVarSortedMaster(i, 1))
    Next i
    
GenerateMasterDescListFromPlTable = arrVarSortedMaster
    
            
End Function


Private Function GenerateMasterPlFinanceTable( _
                                                ByVal dtReportingPeriod As Date, _
                                                ByRef objTargetPl As clsPandL, _
                                                ByRef collActivities As Collection) _
                                                As Scripting.Dictionary

'with p&l
'loop activities and find p&l finance table matches by dict key
'for each line check if dictionary key exists for description
'if not create dictionary entry with key=desc, item=2d array
'2d array colums (0)Rev_Cost as string, (1)Desc_Group as string, (2)Description as string, (3)AmountUSD as double
'if exists then add the target month value to the existing key value
'return dictionary

Dim i As Long, j As Long, k As Long
Dim dictPlFinanceData                   As New Scripting.Dictionary     'Create new dictionary object
Dim objActivity                         As clsActivity                  'single activity object from collection
Dim intReportingMonthColumnIndex        As Integer                      'reporting month column from finance data table header
Dim arrIntFinanceDataTargetColumns      As Variant                      'Array holding the desired column indexs
Dim arrVarPl

'main activity loop
    For Each objActivity In collActivities
        With objActivity
            If .dictFinanceDataTable.Exists(objTargetPl.strName) Then
                'find reporting month column index
                    For i = LBound(.dictFinanceDataTableHeader(objTargetPl.strName), 1) To UBound(.dictFinanceDataTableHeader(objTargetPl.strName), 1)
                        If .dictFinanceDataTableHeader(objTargetPl.strName)(i) = Format(dtReportingPeriod, "MMM-YYYY") Then intReportingMonthColumnIndex = i
                    Next i
                    
                'build array that holds target column indexes
                    ReDim arrIntFinanceDataTargetColumns(0 To 3)
                    '(0)Rev_Cost as string, (1)Desc_Group as string, (2)Description as string, (3)Target Month as long
                    arrIntFinanceDataTargetColumns = Array(1, 2, 3, CInt(intReportingMonthColumnIndex))
                    
                'loop activity finance data, check for desc keys in dictPlFinanceData, create/add as required
                    
                    For i = LBound(.dictFinanceDataTable(objTargetPl.strName), 1) To UBound(.dictFinanceDataTable(objTargetPl.strName), 1)
                        
                        'check to make sure value is not null or zero
                        If Not IsNull(.dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(3))) Or _
                        Not .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(3)) = 0 Then

                        'Check if key exists.  If yes add to the monthly value, if not create new key and build array from activity finance table
                            If dictPlFinanceData.Exists(.dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(2))) Then
                                
                                'dictPlFinanceData(.dictFinanceDataTable(objTargetPl.strName)(i, 3))(3) = $amount stored in pl dictionary array
                                '.dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(3)) = $amount stored in activity finance array
                                Dim arrVarTemp   As Variant     'Holds the dictionary array so it can be so the value in (3) can be added
                                
                                arrVarTemp = dictPlFinanceData(.dictFinanceDataTable(objTargetPl.strName)(i, 3))
                                arrVarTemp(3) = arrVarTemp(3) + .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(3))
                                dictPlFinanceData(.dictFinanceDataTable(objTargetPl.strName)(i, 3)) = arrVarTemp
                                
                            Else
                                dictPlFinanceData.Add _
                                    Key:=.dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(2)), _
                                    Item:=Array( _
                                                    .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(0)), _
                                                    .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(1)), _
                                                    .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(2)), _
                                                    .dictFinanceDataTable(objTargetPl.strName)(i, arrIntFinanceDataTargetColumns(3)) _
                                                    )
                                
                            End If
                        End If
                    Next i
            End If
        End With 'objActivityy
    Next objActivity
    
Set GenerateMasterPlFinanceTable = dictPlFinanceData

'    Dim vKey As Variant
'    Dim strT As String
'
'    For Each vKey In dictPlFinanceData.Keys
'        strT = ""
'        Debug.Print Join(dictPlFinanceData(vKey), " | ")
'    Next vKey

End Function


