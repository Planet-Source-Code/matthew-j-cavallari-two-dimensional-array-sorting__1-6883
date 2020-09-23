VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMattsSort"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************************************
'Class:        clsMattsSort
'Description:  A class library with one public array sorting function and four
'              supporting private functions.  The primary work done by the library
'              is done by the recursively called function MultiSortRecur, which
'              sorts any two dimensional array, given the array to be sorted and
'              the order over which the array's second dimension will be sorted.
'              It uses a divide and conquer type algorithm (similar to a QuickSort).
'
'              An example of usage would be where a two dimensional array is seen
'              as a type of recordset where the first dimension describes the rows
'              and the second dimension describes the fields collection.
'              The recordset (array) can then be multisorted according to column order
'              in ascending or descending order.  (i.e.  have the the array sorted by
'              4th column as ascending and upon those which items match, then 2nd
'              column as descending, etc.
'**************************************************************************************

Option Explicit
Option Compare Text

'**************************************************************************************
'Function:     SortOrderValidate
'Parameters:   intLB - lower bound of array's 2nd dimension to be checked against given
'                      sort order
'              intUB - upper bound of array's 2nd dimension to be checked against given
'                      sort order
'              arrSortOrder - the given sort order and order type
'Return:       0 - Valid input
'              1 - Invalid column reference
'              2 - column reference outside of array bounds
'              3 - invalid sort order type
'Description:  sifts through arrSortOrder checking for correct format (intgers and
'              strings that are within the correct bounds.  also, passes array bounds
'              instead of array to save on memory
'**************************************************************************************
Private Function SortOrderValidate(intLB As Integer, intUB As Integer, _
                                   arrSortOrder As Variant) As Integer
    On Error Resume Next
    Dim i As Integer
    Dim intCol As Integer
    Dim strOrder As String
    
    For i = LBound(arrSortOrder) To UBound(arrSortOrder) Step 2
        'convert column index to integer and check for error
        intCol = CInt(arrSortOrder(i))
        If Err.Number <> 0 Or intCol <> arrSortOrder(i) Then
            SortOrderValidate = 1
            Err.Clear
            Exit Function
        End If
        
        If intCol > intUB Or intCol < intLB Then
            SortOrderValidate = 2
            Err.Clear
            Exit Function
        End If
        
        'convert sort order type to string and check for error
        strOrder = CStr(arrSortOrder(i + 1))
        If (Err.Number <> 0 Or ((strOrder <> "asc") And (strOrder <> "dec"))) Then
            SortOrderValidate = 3
            Err.Clear
            Exit Function
        End If
    Next i
    
    SortOrderValidate = 0
End Function

'**************************************************************************************
'Function:     SortCompare
'Parameters:   intFirst - the "row" index of the first row to be compared
'              arrFirst - the array from which the first row is taken
'              intSecond - the "row" index of the second row to be compared
'              arrSecond - the array from which the second row is taken
'              arrSortOrder - the order of which to compare columns of arrays
'Return:       -1 - first row is less than second row
'              0 - first row is equal to second row
'              1 - first row is greater than second row
'**************************************************************************************
Private Function SortCompare(intFirst As Integer, arrFirst As Variant, _
                             intSecond As Integer, arrSecond As Variant, _
                             arrSortOrder As Variant) As Integer
    Dim i As Integer
    Dim intOrder As Integer
    intOrder = 0
    
    For i = LBound(arrSortOrder) To UBound(arrSortOrder) Step 2
        If arrFirst(intFirst, arrSortOrder(i)) < arrSecond(intSecond, arrSortOrder(i)) Then
            intOrder = -1
            Exit For
        
        ElseIf arrFirst(intFirst, arrSortOrder(i)) > arrSecond(intSecond, arrSortOrder(i)) Then
            intOrder = 1
            Exit For
        End If
    Next i
        
    'if sort type is "des" then reverse return value
    If intOrder <> 0 Then
        If CStr(arrSortOrder(i + 1)) = "dec" Then
            intOrder = -intOrder
        End If
    End If
    
    SortCompare = intOrder
End Function

'**************************************************************************************
'Sub:          RowCopy
'Parameters:   intTo - the row index to be copied to
'              arrTo - the array of the row to be copied to
'              intFrom - the row to be copied from
'              arrFrom - the array of the row to be copied from
'**************************************************************************************
Private Sub RowCopy(intTo As Integer, arrTo As Variant, intFrom As Integer, _
                    arrFrom As Variant)
    Dim i As Integer
    
    For i = LBound(arrTo, 2) To UBound(arrTo, 2)
        arrTo(intTo, i) = arrFrom(intFrom, i)
    Next i
End Sub

'**************************************************************************************
'Sub:          MultiSortRecur
'Parameters:   intBottom - the bottom of the sub array to be sorted
'              intTop - the top of the sub array to be sorted
'              arrArray - array to be sorted
'              arrSortOrder - the sort order to be used
'**************************************************************************************
Private Sub MultiSortRecur(ByVal intBottom As Integer, ByVal intTop As Integer, _
                           arrArray As Variant, arrSortOrder As Variant)
    Dim intLow As Integer
    Dim intHigh As Integer
    Dim arrTemp As Variant
    
    'base case of 0 or 1 rows left in array, so subarray is sorted
    If intBottom >= intTop Then
        Exit Sub
    End If
    
    'copy of middle value over which to create the two sublists
    ReDim arrTemp(1, LBound(arrArray, 2) To UBound(arrArray, 2)) As Variant
    
    intLow = intBottom
    intHigh = intTop
    RowCopy 1, arrTemp, intBottom, arrArray
    
    'move bigger rows above the median row and smaller ones bellow
    Do
        'sift down the array for a row less than the median row
        Do While SortCompare(intHigh, arrArray, 1, arrTemp, arrSortOrder) > -1
            intHigh = intHigh - 1
            If intHigh <= intLow Then
                Exit Do
            End If
        Loop
        
        If intHigh <= intLow Then
            RowCopy intLow, arrArray, 1, arrTemp
            Exit Do
        End If
        
        RowCopy intLow, arrArray, intHigh, arrArray
        
        'sift up the array for a row greater than the median row
        intLow = intLow + 1
        Do While SortCompare(intLow, arrArray, 1, arrTemp, arrSortOrder) = -1
            intLow = intLow + 1
            If intLow >= intHigh Then
                Exit Do
            End If
        Loop
        
        If intLow >= intHigh Then
            intLow = intHigh
            RowCopy intHigh, arrArray, 1, arrTemp
            Exit Do
        End If
        
        RowCopy intHigh, arrArray, intLow, arrArray
    Loop
    
    'now sort the two sub arrays
    MultiSortRecur intBottom, intLow - 1, arrArray, arrSortOrder
    MultiSortRecur intLow + 1, intTop, arrArray, arrSortOrder
End Sub

'**************************************************************************************
'Sub:          MultiSort
'Parameters:   arrArray - array to be sorted
'              SortOrder - array of columns and sort types to be sorted by
'              i.e.
'**************************************************************************************
Public Sub MultiSort(arrArray As Variant, ParamArray SortOrder() As Variant)
    Dim intValNum As Integer
    Dim intBottom As Integer
    Dim arrSortOrder As Variant
    Dim intTop As Integer
    
    'check if arrArray is a valid array
    If Not (IsArray(arrArray)) Then
        MsgBox "Error:  Not A Valid Array"
        Exit Sub
    End If
    
    'create a usable array out of the paramarray and validate it
    arrSortOrder = SortOrder()
    intValNum = 0
    intValNum = SortOrderValidate(LBound(arrArray, 2), UBound(arrArray, 2), arrSortOrder)
    If intValNum = 1 Then
        MsgBox "Error:  Invalid Column Reference"
        Exit Sub
    ElseIf intValNum = 2 Then
        MsgBox "Error:  Column Reference Outside Of Array Bounds"
        Exit Sub
    ElseIf intValNum = 3 Then
        MsgBox "Error:  Invalid Sort Order Type"
        Exit Sub
    End If
    
    'recursively sort the array
    intBottom = LBound(arrArray, 1)
    intTop = UBound(arrArray, 1)
    MultiSortRecur intBottom, intTop, arrArray, arrSortOrder
    
End Sub


