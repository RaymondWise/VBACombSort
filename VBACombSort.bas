Attribute VB_Name = "VBACombSort"
Option Explicit

Public Sub TestCombSort()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    Dim rawDataArray As Variant
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    rawDataArray = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, 1))
    Dim index As Long
    
    Dim isNumberArray As Boolean
    isNumberArray = isNumeric(rawDataArray(1, 1))
    
    If isNumberArray Then
        Dim numberArray() As Double
        ReDim numberArray(LBound(rawDataArray) To UBound(rawDataArray))
        numberArray() = IsNumericalArray(rawDataArray)
        CombSortNumbers numberArray, False
        targetSheet.Range(targetSheet.Cells(1, 2), targetSheet.Cells(lastRow, 2)) = Application.Transpose(numberArray)
    Else:
        Dim stringArray() As String
        ReDim stringArray(LBound(rawDataArray) To UBound(rawDataArray))
        stringArray() = IsStringArray(rawDataArray)
        CombSortStrings stringArray, False
        targetSheet.Range(targetSheet.Cells(1, 2), targetSheet.Cells(lastRow, 2)) = Application.Transpose(stringArray)
    End If
    
errHandler:
    Exit Sub
    
End Sub

Private Function IsNumericalArray(ByVal rawDataArray As Variant) As Double()
    Dim numberArray() As Double
    ReDim numberArray(LBound(rawDataArray) To UBound(rawDataArray))
    Dim index As Long
    For index = LBound(rawDataArray) To UBound(rawDataArray)
        If Not isNumeric(rawDataArray(index, 1)) Then GoTo errHandler
        numberArray(index) = CStr(rawDataArray(index, 1))
    Next
    IsNumericalArray = numberArray
    Exit Function
errHandler:
    MsgBox "not number"
End Function
Private Function IsStringArray(ByVal rawDataArray As Variant) As String()
    Dim stringArray() As String
    ReDim stringArray(LBound(rawDataArray) To UBound(rawDataArray))
    Dim index As Long
    For index = LBound(rawDataArray) To UBound(rawDataArray)
        If isNumeric(rawDataArray(index, 1)) Then GoTo errHandler
        stringArray(index) = CStr(rawDataArray(index, 1))
    Next
    IsStringArray = stringArray
    Exit Function
errHandler:
    MsgBox "not string"
End Function

Private Sub CombSortNumbers(ByRef numberArray() As Double, Optional ByVal sortAscending As Boolean = True)
    Const SHRINK As Double = 1.3
    Dim initialSize As Long
    initialSize = UBound(numberArray())
    Dim gap As Long
    gap = initialSize
    Dim index As Long
    Dim isSorted As Boolean
    
    Do While gap > 1 And Not isSorted
        gap = Int(gap / SHRINK)
        If gap > 1 Then
            isSorted = False
        Else
            gap = 1
            isSorted = True
        End If
        index = 1
        Do While index + gap <= initialSize
            If sortAscending Then
                If numberArray(index) > numberArray(index + gap) Then
                    SwapNumberElements numberArray, index, index + gap
                    isSorted = False
                End If
            Else
                If numberArray(index) < numberArray(index + gap) Then
                    SwapNumberElements numberArray, index, index + gap
                    isSorted = False
                End If
            End If
            index = index + 1
        Loop
    Loop
    
End Sub

Private Sub SwapNumberElements(ByRef numberArray() As Double, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As Double
    temporaryHolder = numberArray(i)
    numberArray(i) = numberArray(j)
    numberArray(j) = temporaryHolder
End Sub

Private Sub CombSortStrings(ByRef stringArray() As String, Optional ByVal sortAscending As Boolean = True)
    Const SHRINK As Double = 1.3
    Dim initialSize As Long
    initialSize = UBound(stringArray())
    Dim gap As Long
    gap = initialSize
    Dim index As Long
    Dim isSorted As Boolean
    
    Do While gap > 1 And Not isSorted
        gap = Int(gap / SHRINK)
        If gap > 1 Then
            isSorted = False
        Else
            gap = 1
            isSorted = True
        End If
        index = 1
        Do While index + gap <= initialSize
            If sortAscending Then
                If stringArray(index) > stringArray(index + gap) Then
                    SwapStringElements stringArray, index, index + gap
                    isSorted = False
                End If
            Else
                If stringArray(index) < stringArray(index + gap) Then
                    SwapStringElements stringArray, index, index + gap
                    isSorted = False
                End If
            End If
            index = index + 1
        Loop
    Loop
    
End Sub

Private Sub SwapStringElements(ByRef stringArray() As String, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As String
    temporaryHolder = stringArray(i)
    stringArray(i) = stringArray(j)
    stringArray(j) = temporaryHolder
End Sub
