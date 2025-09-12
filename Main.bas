Attribute VB_Name = "Main"
Option Explicit
'Chemotherapy Order Maker makes use of calendar creator code written by
'Eric Bentzen, August 2017 and provided freely with no copyright.  Classes clDay and clMonth are
'entirely his work.  clCalendar is a modified version of his code that interfaces with my clAdminDays
'rather than his original Holidays class, which is not included in this project.  The rest of the code
'in this workbook was written by me, Blake Vander Wood, in February - March 2019.  Some common functions
'like bubble sort and the trick for seeing if an array is initialized were adapted from examples found
'on Stack Exchange.  I have tried to make the code as robust as possible but in the inevitable
'event that someone breaks something after I have left, I can be reached at Mppanthers18@gmail.com.
'The password for all worksheets is my first name, all lower case.


'Main serves to start and call other pieces of the program

Sub switchEntryType()
    'Swaps back and forth between a pure days entry (can accomodate up to 90 days) and a day/week entry
    'which can accomodate up to 6 months.  It does this by showing and hiding the input controls (which
    'are just shapes) and then reformatting some cells.  It also changes a flag from True (days only) to false (days and weeks)
    
    Dim i As Integer, j As Integer, k As Integer
    Dim s As String
    Dim OleObj As OLEObject
    
    'Unprotect the sheet and disable screen updating for speed
    Worksheets("creator").Unprotect "blake"
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'If E6 is true, then we are in Days only entry already, so we will swap to days/week entry
    If Worksheets("Backend").Range("E6") Then
        'Change the label on our button to toggle & change E6 to False (Days/Weeks Entry)
        Worksheets("creator").Shapes("btnToggleDaysEntry").TextFrame.Characters.Text = "Toggle to Days only Entry"
        Worksheets("Backend").Range("E6") = False
        
        'Change our column headings to Days/Weeks
        Worksheets("creator").Range("K13").MergeArea.UnMerge
        Worksheets("creator").Range("L13:M13").Merge
        Worksheets("creator").Range("K13") = "Delivered on Weeks"
        Worksheets("creator").Range("K13").Borders(xlEdgeRight).LineStyle = xlContinuous
        Worksheets("creator").Range("L13") = "Delivered on Days"
        
        'Hide all of our controls for days only entry
        For k = 1 To 10
            For i = 1 To 3
                For j = 1 To 30
                    s = "bxd" & k & "dd" & ((i - 1) * 30) + j
                    Worksheets("creator").Shapes(s).Visible = False
                Next j
            Next i
        Next k
        
        
        For k = 1 To 10
            'Show all of the controls for weeks
            For i = 1 To 2
                For j = 1 To 12
                    s = "bxd" & k & "w" & ((i - 1) * 12) + j
                    Worksheets("creator").Shapes(s).Visible = True
                Next j
            Next i
                
            'Now the days
            For j = 1 To 7
                s = "bxd" & k & "d" & j
                Worksheets("creator").Shapes(s).Visible = True
            Next j
        Next k
        
        'Resize the row heights and move the combo boxes within them
        For i = 0 To 9
            Worksheets("creator").Range("DrugStart").offset(i).RowHeight = 36
            Worksheets("creator").Range("DrugStart").offset(i, 10).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next i
        
        For Each OleObj In Worksheets("creator").OLEObjects
            If TypeName(OleObj.Object) = "ComboBox" And OleObj.Name <> "cmbProvider" Then
                OleObj.Top = OleObj.Top - 10
            End If
        Next
    'If E6 was false, then we were in days/weeks entry and need to switch to days only
    Else
        'Change the label on our button to toggle & change E6 to True (Days only entry)
        Worksheets("creator").Shapes("btnToggleDaysEntry").TextFrame.Characters.Text = "Toggle to Days/Weeks Entry"
        Worksheets("Backend").Range("E6") = True
        
        'Change the column heading
        Worksheets("creator").Range("L13").MergeArea.UnMerge
        Worksheets("creator").Range("K13:M13").Merge
        Worksheets("creator").Range("K13") = "Delivered on Days"
        
        'Show the controls for days only
        For k = 1 To 10
            For i = 1 To 3
                For j = 1 To 30
                    s = "bxd" & k & "dd" & ((i - 1) * 30) + j
                    Worksheets("creator").Shapes(s).Visible = True
                Next j
            Next i
        Next k
        
        For k = 1 To 10
            'Hide the controls for weeks
            For i = 1 To 2
                For j = 1 To 12
                    s = "bxd" & k & "w" & ((i - 1) * 12) + j
                    Worksheets("creator").Shapes(s).Visible = False
                Next j
            Next i
                
            'Hide the controls for days
            For j = 1 To 7
                s = "bxd" & k & "d" & j
                Worksheets("creator").Shapes(s).Visible = False
            Next j
        Next k
        
        'Resize the rows and move the comboboxes
        For i = 0 To 9
            Worksheets("creator").Range("DrugStart").offset(i).RowHeight = 54
            Worksheets("creator").Range("DrugStart").offset(i, 10).Borders(xlEdgeRight).LineStyle = xlNone
        Next i
        
        For Each OleObj In Worksheets("creator").OLEObjects
            If TypeName(OleObj.Object) = "ComboBox" And OleObj.Name <> "cmbProvider" Then
                OleObj.Top = OleObj.Top + 10
            End If
        Next
    End If
    Worksheets("creator").Protect "blake"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
Sub GenerateCalendar()
    initializeData
    CalBegin
End Sub
Sub GenerateOrders()
    'Check and make sure we have required inputs, then proceed with order generation
    With Worksheets("Creator")
        If Len(.Range("patheight").Value) = 0 Then
            MsgBox "Please enter a height."
            Exit Sub
        ElseIf Len(.Range("patweight").Value) = 0 Then
            MsgBox "Please enter a weight."
            Exit Sub
        ElseIf Len(.Range("patAllergies").Value) = 0 Then
            MsgBox "Please enter allergies."
            Exit Sub
        ElseIf Len(.Range("cycle").Value) = 0 Then
            MsgBox "Please enter starting cycle."
            Exit Sub
        ElseIf Len(.Range("clength").Value) = 0 Then
            MsgBox "Please enter cycle length."
            Exit Sub
        ElseIf Len(.Range("patdiagnosis").Value) = 0 Then
            MsgBox "Please enter diagnosis."
            Exit Sub
        ElseIf .OLEObjects("cmbProvider").Object.ListIndex = -1 Then
            MsgBox "Please enter a provider."
            Exit Sub
        ElseIf Len(.Range("F3").Value) = 0 Then
            MsgBox "Please enter a starting date."
            Exit Sub
        End If
    End With
    initializeData
    PreMedications.Show
End Sub
Sub ClearContents()
'Small subroutine to clear all data entry cells and reset the ListBoxes for days/weeks to be blank,
    'essentially clearing out the worksheet so that a set of data can be placed
    Dim mbResult As Integer
    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim s As String
    Dim OleObj As OLEObject
    
    
    'Generate confirmation dialogue box
    mbResult = MsgBox("Are you sure you'd like to clear all cells and reset the template?", vbYesNo, "Really clear?")
    
    Select Case mbResult
        Case vbYes
            'If the user is sure they want to clear the form, proceed
            With ActiveSheet
                'Get the patient and cycle data out of the way first
                .Range("I3").ClearContents
                .Range("F3").ClearContents
                .Range("I5").ClearContents
                .Range("L3").ClearContents
                .Range("L5").ClearContents
                .Range("F9").ClearContents
                .Range("F11").ClearContents
                .Range("I9").ClearContents
                .Range("I11").ClearContents
                .Range("L9").ClearContents
                .Range("C11").MergeArea.ClearContents
                
                'Now delete the drug data
                .Range("A14:B23").ClearContents
                .Range("D14:D23").ClearContents
                For i = 14 To 23
                    .Cells(i, 5).MergeArea.ClearContents
                Next
                
                'Rest the combo boxes
                For Each OleObj In .OLEObjects
                    If TypeName(OleObj.Object) = "ComboBox" Then
                        OleObj.Object.ListIndex = -1
                    End If
                Next
                
                For k = 1 To 10
                    For i = 1 To 3
                        For j = 1 To 30
                            s = "d" & k & "dd" & ((i - 1) * 30) + j
                            If Worksheets("controlstates").Range(s).Value Then
                                Worksheets("controlstates").Range(s).Value = False
                                s = "bx" & s
                                Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Fill.ForeColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Line.ForeColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Color = RGB(0, 51, 89)
                                Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Bold = False
                            End If
                        Next j
                    Next i
                Next k
                
                'Reset the weeks
                For k = 1 To 10
                    For i = 1 To 2
                        For j = 1 To 12
                            s = "d" & k & "w" & ((i - 1) * 12) + j
                            If Worksheets("controlstates").Range(s).Value Then
                                Worksheets("controlstates").Range(s).Value = False
                                s = "bx" & s
                                Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Fill.ForeColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Line.ForeColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                                Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Color = RGB(0, 51, 89)
                                Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Bold = False
                            End If
                        Next j
                    Next i
                        
                    'Now the days
                    For j = 1 To 7
                        s = "d" & k & "d" & j
                        If Worksheets("controlstates").Range(s).Value Then
                            Worksheets("controlstates").Range(s).Value = False
                            s = "bx" & s
                            Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                            Worksheets("creator").Shapes(s).Fill.ForeColor.RGB = RGB(60, 182, 206)
                            Worksheets("creator").Shapes(s).Line.ForeColor.RGB = RGB(60, 182, 206)
                            Worksheets("creator").Shapes(s).Fill.BackColor.RGB = RGB(60, 182, 206)
                            Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Color = RGB(0, 51, 89)
                            Worksheets("creator").Shapes(s).TextFrame.Characters.Font.Bold = False
                        End If
                    Next j
                Next k
                
            End With
        Case vbNo
            'If they do not want to clear the form, exit
            Exit Sub
    End Select
    
BeforeExit:
'Shouldn't be strictly necessary as this isn't C, but good practice nonetheless
Set OleObj = Nothing
Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure ClearContents"
Resume BeforeExit
End Sub
Sub initializeData()
    Dim i As Integer

    'Set three important global variables
    InitialWorkbook = ActiveWorkbook.Name
    startDate = ActiveWorkbook.Sheets("Creator").Range("F3").Value
    bDaysEntry = ActiveWorkbook.Sheets("Backend").Range("E6")
    
    'Set the master list of doses for drugs
    With ActiveWorkbook.Sheets("Backend")
        For i = 0 To Range(.Range("A6"), .Range("A6").End(xlDown)).Rows.Count - 1
            MasterDoseList(i) = .Range("A6").offset(i)
        Next
    End With
    
    'Create our major collections
    Set Information = Nothing
    Set Information = New clInformation
    Set Drugs = Nothing
    Set Drugs = New clDrugs
    Set AdminDays = Nothing
    Set AdminDays = New clAdminDays
    Set Labs = Nothing
    Set Labs = New clLabs
    Set PreMeds = Nothing
    Set PreMeds = New clPreMeds
End Sub
Sub CalBegin()
    Set Calendar = Nothing
    Set Calendar = New clCalendar
    InitialWorkbook = ActiveWorkbook.Name
    Calendar.StartCal
    If bAbort Then bAbort = False
End Sub
Sub cleanClose()
    Set Information = Nothing
    Set OrderSheets = Nothing
    Set Drugs = Nothing
    Set AdminDays = Nothing
    Set Labs = Nothing
    Set PreMeds = Nothing
    Set Calendar = Nothing
End Sub
