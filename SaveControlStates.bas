Attribute VB_Name = "SaveControlStates"
Option Explicit
Sub controlChange()
    'A shared macro for all 310 of our day/week boxes, which are really just plain shapes
    'with a macro assigned to them.  It is assigned to the "on activate" function for every box,
    'so when a box is clicked on it will reverse the associated value in a true false table in a
    'hidden worksheet and reverse the color of the box
    
    'In earlier versions of this program, the day/week boxes were ActiveX listbox controls, but they
    'were extremely finicky with different screen resolutions and across different version of excel
    'This solution is slower, but more widely compatible.
    Dim cb As String
    Dim shps As Shapes
    Dim s As String
    
    Application.ScreenUpdating = False
    cb = Application.Caller
    Set shps = Worksheets("Creator").Shapes
    
    s = shps(cb).Name
    s = mid(s, 3)
    With Worksheets("controlstates").Range(s)
        If .Value Then                                                      'If our box was already set to true
            .Value = False                                                  'Make it false and unhighlight it
            shps(cb).Fill.BackColor.RGB = RGB(60, 182, 206)
            shps(cb).Fill.ForeColor.RGB = RGB(60, 182, 206)
            shps(cb).Line.ForeColor.RGB = RGB(60, 182, 206)
            shps(cb).Fill.BackColor.RGB = RGB(60, 182, 206)
            shps(cb).TextFrame.Characters.Font.Color = RGB(0, 51, 89)
            shps(cb).TextFrame.Characters.Font.Bold = False
        Else
            .Value = True                                                   'If it was false, make it true and
            shps(cb).Fill.BackColor.RGB = RGB(0, 51, 89)                    'highlight it
            shps(cb).Fill.ForeColor.RGB = RGB(0, 51, 89)
            shps(cb).Line.ForeColor.RGB = RGB(0, 51, 89)
            shps(cb).Fill.BackColor.RGB = RGB(0, 51, 89)
            shps(cb).TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            shps(cb).TextFrame.Characters.Font.Bold = True
        End If
    End With
    Application.ScreenUpdating = True
End Sub
