Attribute VB_Name = "MatrixMultiplication"
Option Explicit

Const p = 2
Const q = 2
Const r = 2

Sub MatrixMultiplication()
    Dim wsMm As Worksheet
    Dim m1#(p - 1, q - 1), m2#(q - 1, r - 1), i&, j&, k&, mOutput#(p - 1, r - 1)
    
    Set wsMm = ThisWorkbook.Worksheets("Matrix multiplication")
    
    ' Cleanup.
    wsMm.Range("G5:H6").ClearContents
    
    ' Initialization of arrays:
    For i = 0 To p - 1
        For j = 0 To q - 1
            m1(i, j) = wsMm.Cells(i + 1, j + 7)
        Next j
    Next i
    For i = 0 To q - 1
        For j = 0 To r - 1
            m2(i, j) = wsMm.Cells(i + 3, j + 7)
        Next j
    Next i
    
    For i = 0 To p - 1
        For j = 0 To r - 1
            ' Multiplication
            mOutput(i, j) = 0#
            For k = 0 To q - 1
                mOutput(i, j) = mOutput(i, j) + m1(i, k) * m2(k, j)
            Next k
            ' Output:
            wsMm.Cells(i + 5, j + 7) = mOutput(i, j)
        Next j
    Next i
End Sub
