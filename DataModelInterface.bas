Attribute VB_Name = "DataModelInterface"
'@Folder("DataLoad")
'@Description "Procedures related to the loading of data and tables from the data model"
Option Explicit

'@Description "Primary ADODB Interface, returns raw table of data"
Function GetTableDataFromDataModel( _
                                wbUpsfin As Workbook, _
                                strMdxPath As String) _
                                As Scripting.Dictionary


Dim conn    As ADODB.Connection:    Set conn = wbUpsfin.Model.DataModelConnection.ModelConnection.ADOConnection
Dim rs      As ADODB.Recordset:     Set rs = New ADODB.Recordset

Dim dictOut As New Scripting.Dictionary

With rs

    'Activate connection to Data Model
        .ActiveConnection = conn
        .Open strMdxPath, conn, adOpenForwardOnly, adLockOptimistic
    
    'Retrieve table
        dictOut.Add Key:="table", Item:=AdodbGetRowsFromRecordSet(rs)
        
    'Retrieve headers
        dictOut.Add Key:="header", Item:=AdodbGetFieldsFromRecordSet(rs)
    
    'Close connection
        .Close
        conn.Close

End With 'With rs

Set GetTableDataFromDataModel = dictOut

End Function

'@Description "Returns 2D array containing transposed RecordSet Rows"
Private Function AdodbGetRowsFromRecordSet( _
                                        rs As ADODB.Recordset) _
                                        As Variant

Dim arrOut    As Variant

If rs.RecordCount = 0 Then
    arrOut = -1
Else
    arrOut = rs.GetRows
End If

If IsArray(arrOut) Then arrOut = GenericFunctions.TransposeArray(arrOut)

AdodbGetRowsFromRecordSet = arrOut
    

End Function


'@Description "Returns a 1D array containing the RecordSet Fields"
Private Function AdodbGetFieldsFromRecordSet( _
                                        rs As ADODB.Recordset) _
                                        As Variant

Dim arrOut()    As Variant
Dim i           As Long

With rs
    
    'If no fields return -1 as error indicator else iterate through list and populate array
        If .Fields.Count = 0 Then
            AdodbGetFieldsFromRecordSet = -1
        Else
            ReDim arrOut(.Fields.Count - 1)
            For i = 0 To .Fields.Count - 1
                arrOut(i) = rs.Fields(i).Name
            Next i
            AdodbGetFieldsFromRecordSet = arrOut
        End If
            
End With

End Function

