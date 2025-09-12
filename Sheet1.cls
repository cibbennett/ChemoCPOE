VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'These were used to generate all the boxes that serves as controls

Sub initializenewboxes()
    Dim i As Integer, j As Integer, k As Integer
    Dim s As String
    
    For k = 1 To 10
        For i = 1 To 3
            For j = 1 To 30
                s = "bxd" & k & "dd" & ((i - 1) * 30) + j
                Worksheets("creator").Shapes(s).Select
            Next j
        Next i
    Next k
End Sub


Sub makeboxes()
    Dim i As Integer, j As Integer, k As Integer
    Dim stTop As Long, stLeft As Long
    Dim curtop As Long, curLeft As Long
    Dim curShape As Shape
    
    stTop = 234
    stLeft = 909
    
    For k = 1 To 10
        For i = 1 To 2
            For j = 1 To 12
                curtop = stTop + ((i - 1) * 17.25) + (k - 1) * 36
                curLeft = stLeft + ((j - 1) * 18)
                Set curShape = Worksheets("Creator").Shapes.AddShape(msoShapeRectangle, curLeft, curtop, 18.8, 17)
                curShape.Fill.ForeColor.RGB = RGB(60, 182, 206)
                curShape.Fill.BackColor.RGB = RGB(60, 182, 206)
                curShape.Line.ForeColor.RGB = RGB(60, 182, 206)
                curShape.Line.BackColor.RGB = RGB(60, 182, 206)
                curShape.Name = "bxd" & k & "w" & ((i - 1) * 12) + j
                curShape.TextFrame.Characters.Text = ((i - 1) * 12) + j
                curShape.TextFrame.Characters.Font.Name = "Gil Sans MT"
                curShape.TextFrame.Characters.Font.Size = 11
                curShape.TextFrame.Characters.Font.Color = RGB(0, 51, 89)
                curShape.TextFrame.HorizontalAlignment = xlHAlignCenter
                curShape.TextFrame.VerticalAlignment = xlVAlignCenter
                curShape.TextFrame.MarginLeft = 0
                curShape.TextFrame.MarginRight = 0
                curShape.TextFrame.MarginTop = 0
                curShape.TextFrame.MarginBottom = 0
                curShape.OnAction = "'" & ActiveWorkbook.Name & "'!controlChange"
            Next j
        Next i
        
        For i = 1 To 7
            curtop = stTop + (k - 1) * 36
            curLeft = 1138.5 + (i - 1) * 19.5
            Set curShape = Worksheets("Creator").Shapes.AddShape(msoShapeRectangle, curLeft, curtop, 19.5, 34)
            curShape.Fill.ForeColor.RGB = RGB(60, 182, 206)
            curShape.Fill.BackColor.RGB = RGB(60, 182, 206)
            curShape.Line.ForeColor.RGB = RGB(60, 182, 206)
            curShape.Line.BackColor.RGB = RGB(60, 182, 206)
            curShape.Name = "bxd" & k & "d" & i
            curShape.TextFrame.Characters.Text = i
            curShape.TextFrame.Characters.Font.Name = "Gil Sans MT"
            curShape.TextFrame.Characters.Font.Size = 11
            curShape.TextFrame.Characters.Font.Color = RGB(0, 51, 89)
            curShape.TextFrame.HorizontalAlignment = xlHAlignCenter
            curShape.TextFrame.VerticalAlignment = xlVAlignCenter
            curShape.TextFrame.MarginLeft = 0
            curShape.TextFrame.MarginRight = 0
            curShape.TextFrame.MarginTop = 0
            curShape.TextFrame.MarginBottom = 0
            curShape.OnAction = "'" & ActiveWorkbook.Name & "'!controlChange"
        Next i
    Next k
End Sub
Sub makemoreboxes()
    Dim i As Integer, j As Integer, k As Integer
    Dim stTop As Long, stLeft As Long
    Dim curtop As Long, curLeft As Long
    Dim curShape As Shape
    
    stTop = 234
    stLeft = 909
    
    For k = 1 To 10
        For i = 1 To 3
            For j = 1 To 30
                curtop = stTop + ((i - 1) * 17) + (k - 1) * 54
                curLeft = stLeft + ((j - 1) * 12.25)
                Set curShape = Worksheets("Creator").Shapes.AddShape(msoShapeRectangle, curLeft, curtop, 12.25, 17)
                curShape.Fill.ForeColor.RGB = RGB(60, 182, 206)
                curShape.Fill.BackColor.RGB = RGB(60, 182, 206)
                curShape.Line.ForeColor.RGB = RGB(60, 182, 206)
                curShape.Line.BackColor.RGB = RGB(60, 182, 206)
                curShape.Name = "bxd" & k & "dd" & ((i - 1) * 30) + j
                curShape.TextFrame.Characters.Text = ((i - 1) * 30) + j
                curShape.TextFrame.Characters.Font.Name = "Gil Sans MT"
                curShape.TextFrame.Characters.Font.Size = 9
                curShape.TextFrame.Characters.Font.Color = RGB(0, 51, 89)
                curShape.TextFrame.HorizontalAlignment = xlHAlignCenter
                curShape.TextFrame.VerticalAlignment = xlVAlignCenter
                curShape.TextFrame.MarginLeft = 0
                curShape.TextFrame.MarginRight = 0
                curShape.TextFrame.MarginTop = 0
                curShape.TextFrame.MarginBottom = 0
                curShape.OnAction = "'" & ActiveWorkbook.Name & "'!controlChange"
            Next j
        Next i
    Next k
End Sub
