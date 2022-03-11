Attribute VB_Name = "GenericFunctions"
'@Folder("Main")
'@Description "Holds Generic Subs and Functions"
Option Explicit

'@Description "Sort a 2-D array"
Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next
    'from https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206
    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)
    
End Sub

'@Description "Replaces characters that are not allowed in named ranges with a legal character"
Function replaceIllegalNamedRangeCharacters( _
                                                vInput As Variant, _
                                                Optional strReplacement As String = ".") _
                                                As Variant

Dim vOriginal           As Variant
Dim vIllegalChrList     As Variant
Dim vElement            As Variant
Dim vReturn             As Variant

vOriginal = vInput

vIllegalChrList = Array(" ", ";", "-", "(", ")", "+", "/", "\", "_")

For Each vElement In vIllegalChrList
    vOriginal = Replace(vOriginal, vElement, strReplacement)
Next vElement

replaceIllegalNamedRangeCharacters = vOriginal

End Function

'@Description "Writes the P&L Name and the reporting month to the relevant fields in the worksheet"
Sub writeProjectHeadersToPafWorksheet( _
                                                ByRef ws As Worksheet, _
                                                ByVal dtReportingPeriod As Date, _
                                                ByVal strPlName As String, _
                                                ByVal strSheetName As String)
                                                
ws.Range(strSheetName & "_Header_PL.Name").Value = strPlName
ws.Range(strSheetName & "_Header_Reporting.Month").Value = Format(dtReportingPeriod, "MMM-YYYY")

End Sub

'@Description "Checks if a key already exists in a collection"
'from: https://stackoverflow.com/questions/38007844/generic-way-to-check-if-a-key-is-in-a-collection-in-excel-vba

Function HasKey(coll As Collection, strKey As String) As Boolean
    
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (Err.Number = 0)
    Err.Clear
    
End Function

'@Description "Returns an array with the index of selected columns"
Function GetHeaderColumnIndexes( _
                                    arrVarInput As Variant, _
                                    arrVarColumnNames As Variant, _
                                    Optional intInputArrayDimension As Integer, _
                                    Optional intInputArrayHeaderRowNumber As Integer = 1 _
                                    ) As Variant
                            
 Dim i As Long, j As Long, k As Long
                            
'placeholder array for function assignment
    Dim arrIntIndexOutput() As Integer

'Resize array for number of elements to be found
    ReDim arrIntIndexOutput(0 To UBound(arrVarColumnNames))
                                    
'If array dimension is provided use it and build index, otherwise don't and build index
    If Not intInputArrayDimension = 0 Then
        For i = 0 To UBound(arrVarColumnNames)
            For j = LBound(arrVarInput, intInputArrayDimension) To UBound(arrVarInput, intInputArrayDimension)
                If arrVarInput(intInputArrayHeaderRowNumber, j) = arrVarColumnNames(i) Then arrIntIndexOutput(i) = j
            Next j
        Next i
    Else
        For i = 0 To UBound(arrVarColumnNames)
            For j = LBound(arrVarInput) To UBound(arrVarInput)
                If arrVarInput(j) = arrVarColumnNames(i) Then arrIntIndexOutput(i) = j
            Next j
        Next i
    End If
    
GetHeaderColumnIndexes = arrIntIndexOutput

End Function

'@Description "Transposed a 2D array"
Function TransposeArray( _
                            arrIn As Variant) _
                            As Variant
                            
Dim i           As Long
Dim ii          As Long
Dim arrOut()    As Variant

'Redim array to fit incoming
    ReDim arrOut(UBound(arrIn, 2), UBound(arrIn))

'loop through incoming array and transpose
    For i = 0 To UBound(arrIn, 2)
        For ii = 0 To UBound(arrIn)
            arrOut(i, ii) = arrIn(ii, i)
        Next
    Next

TransposeArray = arrOut

End Function

'@Description "Returns a section of an MDX string and removes MDX structure characters"
Function CleanMdxString( _
                            varMdxIn As Variant, _
                            intSplitTargetIndex As Integer, _
                            arrVarReplaceCharacters As Variant, _
                            Optional boolIndexFromEnd As Boolean = False, _
                            Optional strSplitDelimiter As String = "].") _
                            As String
                            
Dim arrStrMdxSplit()    As String
Dim intSplitCount       As Integer
Dim strReturn           As String
Dim element             As Variant
                          
'Split incoming MDX string into array based on delimiter
    arrStrMdxSplit = Split(varMdxIn, strSplitDelimiter)
    intSplitCount = UBound(arrStrMdxSplit)
    If boolIndexFromEnd Then
        strReturn = arrStrMdxSplit(intSplitCount - intSplitTargetIndex)
    Else
        strReturn = arrStrMdxSplit(intSplitTargetIndex - 1)
    End If
   
'For target array element loop through list of characters to remove
    For Each element In arrVarReplaceCharacters
        strReturn = Replace(strReturn, element, "")
    Next element
    
CleanMdxString = strReturn

End Function

'@Description "I strongly dislike the wording for InStr..."
Function StringSearch( _
                        intStartLoc As Integer, _
                        strString As Variant, _
                        strSearchPhrase As Variant) As Variant

StringSearch = InStr(intStartLoc, strString, strSearchPhrase)

End Function

'@Description "Builds a new header array based on string criteria and an array with the index location in the old header array"
Function CreateNewHeaderListWithIndex( _
                                        arrIn As Variant, _
                                        arrStrSearchKeys As Variant) _
                                        As Variant
Dim i As Long, j As Long, k As Long
Dim arrVarNewHeaderList()   As Variant
Dim arrOldArrayIndexList()  As Variant

'Set array counter
    k = 0

'Loop through arrIn, search for string, if found save string and original index
    For i = 0 To UBound(arrIn)
        For j = 0 To UBound(arrStrSearchKeys)
            If StringSearch(1, arrIn(i), arrStrSearchKeys(j)) > 0 Then
                ReDim Preserve arrVarNewHeaderList(k)
                ReDim Preserve arrOldArrayIndexList(k)
                arrVarNewHeaderList(k) = arrIn(i)
                arrOldArrayIndexList(k) = i
                k = k + 1
            End If
        Next j
    Next i

'return both arrays
    CreateNewHeaderListWithIndex = Array(arrVarNewHeaderList, arrOldArrayIndexList)

End Function

