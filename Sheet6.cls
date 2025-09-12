VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub labelcells()
Dim i As Integer, j As Integer, k As Integer
Dim s As String
Dim rowoff As Integer

With Worksheets("controlstates")
    For k = 0 To 9
        For i = 0 To 2
            For j = 0 To 29
                rowoff = (k * 5) + i
                s = "d" & k + 1 & "dd" & (i * 30) + j + 1
                .Range("A43").offset(rowoff, j).Name = s
            Next j
        Next i
    Next k
End With
End Sub
