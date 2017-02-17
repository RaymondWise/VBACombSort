Attribute VB_Name = "VBACombSort"
Option Explicit

Public Sub TestCombSort()
    Const SORT_KEY_COLUMN As Long = 1
    
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    Dim dataArray As Variant
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Dim lastColumn As Long
    lastColumn = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
    
    dataArray = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, lastColumn))
    Dim index As Long
    
    CombSortArray dataArray, lastColumn, SORT_KEY_COLUMN, True
    targetSheet.Range(targetSheet.Cells(1, lastColumn + 1), targetSheet.Cells(lastRow, (lastColumn * 2))) = dataArray
    
End Sub


Private Sub CombSortArray(ByRef dataArray As Variant, Optional ByVal numberOfColumns As Long = 1, Optional ByVal sortKeyColumn As Long = 1, Optional ByVal sortAscending As Boolean = True)
    Const SHRINK As Double = 1.3
    Dim initialSize As Long
    initialSize = UBound(dataArray, 1)
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
                If dataArray(index, sortKeyColumn) > dataArray(index + gap, sortKeyColumn) Then
                    SwapElements dataArray, numberOfColumns, index, index + gap
                    isSorted = False
                End If
            Else
                If dataArray(index) < dataArray(index + gap) Then
                    SwapElements dataArray, numberOfColumns, index, index + gap
                    isSorted = False
                End If
            End If
            index = index + 1
        Loop
    Loop
    
End Sub

Private Sub SwapElements(ByRef dataArray As Variant, ByVal numberOfColumns As Long, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As Variant
    Dim index As Long
    For index = 1 To numberOfColumns
        temporaryHolder = dataArray(i, index)
        dataArray(i, index) = dataArray(j, index)
        dataArray(j, index) = temporaryHolder
    Next
End Sub

