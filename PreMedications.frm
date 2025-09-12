VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreMedications 
   Caption         =   "Pre-Medications"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19905
   OleObjectBlob   =   "PreMedications.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PreMedications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************
' Initialize
'**********************
Private Sub UserForm_Initialize()
    
    On Error GoTo ErrorHandle
    
    'Populate our listbox with the drugs entered
    popListBox
    
    'Read from the Backend sheet to populate our collection of Premedications
    popPreMeds
    
Exit Sub
ErrorHandle:
MsgBox err.Description & " in Userform_Initialize, Premedications"
End Sub
'**********************
'Checkbox Exclusivity
'**********************
Private Sub chkBMP_Click()
    If Me.chkBMP.Value = True Then
        Me.chkCMP.Value = False
    End If
End Sub
Private Sub chkCMP_Click()
    If Me.chkCMP.Value = True Then
        Me.chkBMP.Value = False
    End If
End Sub
'************************
' Button Click Functions
'************************
Private Sub AddMore_Click()
    assignDrugs
    cleanList
    If Me.lstbxDrugsBox.ListCount = 0 Then  'If we've made orders for every drug, make the orders and unload the form
        Call MakeOrderBook
        Unload Me
    End If
End Sub
Private Sub AcceptPremeds_Click()
    assignDrugs
    Call MakeOrderBook
    Unload Me
End Sub
'************************
' Helper Functions
'************************
Private Sub cleanList()
    'Remove drugs that have a premed/lab order set associated with them from the list
    Dim i As Long
    
    On Error GoTo ErrorHandle:
    
    With Me.lstbxDrugsBox
        For i = 0 To .ListCount - 1
            If .ListCount <= i Then  'If all drugs have been evaluated, exit
                Exit Sub
            End If
            If .Selected(i) Then
                .RemoveItem i
                i = i - 1
            End If
        Next
    End With
    
ErrorHandle:
    MsgBox err.Description & " in cleanList, PreMedications form"
    Exit Sub
End Sub
Private Sub assignDrugs()
    Dim chkCur As Control
    Dim xPreMed As clPreMed
    Dim i As Long, j As Integer, k As Integer, l As Integer, ind As Integer
    Dim prmdKey As String
    Dim drgKey As String
    Dim withOff As Integer
    Dim withOffDate As Date
    Dim curWeek As Integer
    
    On Error GoTo ErrorHandle
    
    
    
    For Each chkCur In Me.Controls
        'Not exactly elegant but, we iterate through the form controls, check to see if they are a checkbox
        'and then do things based on what they correlate to
        If TypeName(chkCur) = "CheckBox" Then
            If chkCur.Value = True Then 'Has to be a separate line to keep it from being evaluated when the control is a label
                'For every checkbox that is checked, we're going to associate thing with every drug that is selected
                For i = 0 To Me.lstbxDrugsBox.ListCount - 1     'Go through listbox rows
                    If Me.lstbxDrugsBox.Selected(i) Then        'And find highlighted drugs
                        drgKey = Me.lstbxDrugsBox.List(i, 1)    'Get the key of the highlighted drug (stored in hidden column)
                        If chkCur.Name <> "chkStandbyMeds" And chkCur.Name <> "chkCBC" _
                            And chkCur.Name <> "chkCMP" And chkCur.Name <> "chkBMP" And chkCur.Name <> "chkUrinalysis" _
                             And chkCur.Name <> "chkMagPhos" And chkCur.Name <> "chkCSFGlucose" And chkCur.Name <> "chkCSFCells" _
                              And chkCur.Name <> "chkCSFProtein" Then
                             
                                    prmdKey = mid(chkCur.Name, 4)
                                   'Update the values for days, weeks, and dates in the PreMed
    
                                        PreMeds.Item(prmdKey).InsertDrug drgKey
                                        'Update the Days
                                        For j = 1 To 24
                                            If Drugs.Item(drgKey).bWeek(j) Then
                                                For k = 0 To Drugs.Item(drgKey).numDays(j) - 1
                                                    With PreMeds.Item(prmdKey)
                                                        For l = 0 To .numTiming - 1
                                                            If .numDays(j) < .iMaxFreq Then                                         'Make sure we haven't exceeded the maximum doses per week
                                                                withOff = Drugs.Item(drgKey).iDay(k, j) + .iTiming(l)               'Account for the timing offset
                                                                .InsertDay withOff, j                                               'Insert the appropriate day number
                                                                ind = Drugs.Item(drgKey).linIndex(j, Drugs.Item(drgKey).iDay(k, j)) 'Calculate the index of the correct date
                                                                withOffDate = Drugs.Item(drgKey).datDate(ind) + .iTiming(l)         'Calculate the date for the premed accounting for the timing offset
                                                                .InsertDate withOffDate                                             'insert the date
                                                            End If
                                                        Next l
                                                    End With
                                                    If Len(PreMeds.Item(prmdKey).sLinkedKey) > 0 Then                               'If a linked key exists, enter information for it as well
                                                        With PreMeds.Item(PreMeds.Item(prmdKey).sLinkedKey)
                                                            For l = 0 To .numTiming - 1
                                                                If .numDays(j) < .iMaxFreq Then                                     'Make sure we haven't exceeded the maximum doses per week
                                                                    withOff = Drugs.Item(drgKey).iDay(k, j) + .iTiming(l)           'Account for the timing offset
                                                                    .InsertDay withOff, j                                           'Insert the appropriate day number
                                                                    ind = Drugs.Item(drgKey).linIndex(j, Drugs.Item(drgKey).iDay(k, j)) 'Calculate the index of the correct date
                                                                    withOffDate = Drugs.Item(drgKey).datDate(ind) + .iTiming(l)     'Calculate the date for the premed accounting for the timing offset
                                                                    .InsertDate withOffDate                                         'insert the date
                                                                End If
                                                            Next l
                                                        End With
                                                    End If
                                                Next k
                                            End If
                                        Next j
    
                        'This else captures the labs
                        ElseIf chkCur.Name <> "chkStandbyMeds" Then
                            'Update the Labs collection, adding any selected checkbox to it if it isn't there already
                            'and then updating its days/weeks/dates
                            If Not Labs.ContainsLab(chkCur.Caption) Then
                                Labs.Add chkCur.Caption, chkCur.Caption
                                Labs.Item(chkCur.Caption).sName = chkCur.Caption    'Set the name
                                Labs.Item(chkCur.Caption).InsertDrug drgKey         'Associate the drug with the lab
                            End If
                            With Labs.Item(chkCur.Caption)
                                'Add days
                                For j = 1 To 24
                                    If Drugs.Item(drgKey).bWeek(j) Then
                                        For k = 0 To Drugs.Item(drgKey).numDays(j) - 1
                                            .InsertDay Drugs.Item(drgKey).iDay(k, j), j
                                        Next k
                                   End If
                                Next j
                                'Add dates
                                For j = 0 To Drugs.Item(drgKey).numDates - 1
                                    .InsertDate Drugs.Item(drgKey).datDate(j)
                                Next j
                            End With
    
                        'And finally the standbymeds
                        ElseIf chkCur.Name = "chkStandbyMeds" Then
                            Drugs.Item(drgKey).bStandby = True
                        End If
                    End If
                Next i
                'Reset the checkbox state, in case we're adding more premed orders
                chkCur.Value = False
                
            End If
        'Get any custom input from the text boxes
        'Custom inputs will be stored as members of their respective class
        'When we write a custom premed to the order sheet, we will treat it as a special case
        'and only attempt to write from special instructions
        ElseIf chkCur.Name = "txtBoxCustomPreMeds" Then
            If Len(chkCur.Text) > 0 Then
            For i = 0 To Me.lstbxDrugsBox.ListCount - 1
                If Me.lstbxDrugsBox.Selected(i) Then
                    drgKey = Me.lstbxDrugsBox.List(i, 1)
                    prmdKey = "cstm" & Left(chkCur.Text, 4)
                    If Not PreMeds.ContainsPreMed(prmdKey) Then
                        PreMeds.Add prmdKey, prmdKey
                        PreMeds.Item(prmdKey).sSpecInstruct = chkCur.Text
                        PreMeds.Item(prmdKey).InsertDrug drgKey
                    End If
                    With PreMeds.Item(prmdKey)
                        For j = 1 To 24
                            If Drugs.Item(drgKey).bWeek(j) Then
                                For k = 0 To Drugs.Item(drgKey).numDays(j) - 1
                                    .InsertDay Drugs.Item(drgKey).iDay(k, j), j
                                Next k
                            End If
                        Next j
                        For j = 0 To Drugs.Item(drgKey).numDates - 1
                            .InsertDate Drugs.Item(drgKey).datDate(j)
                        Next j
                    End With
                End If
            Next i
            chkCur.Text = ""
            End If
        ElseIf chkCur.Name = "txtboxCustomLabs" Then
            If Len(chkCur.Text) > 0 Then
            For i = 0 To Me.lstbxDrugsBox.ListCount - 1
                If Me.lstbxDrugsBox.Selected(i) Then
                    drgKey = Me.lstbxDrugsBox.List(i, 1)
                    If Not Labs.ContainsLab(chkCur.Text) Then
                        Labs.Add chkCur.Text, chkCur.Text
                        Labs.Item(chkCur.Text).sName = chkCur.Text
                        Labs.Item(chkCur.Text).InsertDrug drgKey
                    End If
                    With Labs.Item(chkCur.Text)
                        For j = 1 To 24
                            If Drugs.Item(drgKey).bWeek(j) Then
                                For k = 0 To Drugs.Item(drgKey).numDays(j) - 1
                                    .InsertDay Drugs.Item(drgKey).iDay(k, j), j
                                Next k
                            End If
                        Next j
                        For j = 0 To Drugs.Item(drgKey).numDates - 1
                            .InsertDate Drugs.Item(drgKey).datDate(j)
                        Next j
                    End With
                    
                End If
            Next i
            chkCur.Text = ""
            End If
        End If
     Next
    
BeforeExit:
    Set chkCur = Nothing
    Set xPreMed = Nothing
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in assignDrugs, PreMedications Form"
    Resume BeforeExit
End Sub
Private Sub popListBox()
'Small subroutine to populate the listbox from our drugs collection

Dim i As Integer
    
    On Error GoTo ErrorHandle
    
    'Populate our listbox with the drugs entered
    With Me.lstbxDrugsBox
        For i = 1 To Drugs.Count
            .AddItem
            .List(i - 1, 0) = Drugs.Item(i).sDrugName
            .List(i - 1, 1) = Drugs.Item(i).Key             'This column is hidden (its width is zero)
        Next
    End With
    
BeforeExit:
Exit Sub
ErrorHandle:
MsgBox err.Description & " in popListBox, PreMedications Form"
Resume BeforeExit
End Sub
Private Sub popPreMeds()
'Sub reads from excel sheet "Backend" to populate the PreMeds collection
'Each of the three sets of premed types are arranged in tables with the column order:
'(0,0)  (0,1)   (0,2)       (0,3)   (0,4)                 (0,5)         (0,6)   (0,7)   (0,8)
'Name   Dose    MaxDose     Route   SpecialInstructions   Timing        Units   Label   ShowBox
'String Double  Double      String  String                Array of Int  String  String  Boolean


readFillPreMeds "antiemetics", 60, 40
readFillPreMeds "GIProtection", 330, 40
readFillPreMeds "IVFluids", 60, 355


End Sub
Private Function readFillPreMeds(sAddress As String, startTop As Long, startLeft As Long)
'Sub reads from excel sheet "Backend" to populate the PreMeds collection
'Each of the three sets of premed types are arranged in tables with the column order:
'(0,0)  (0,1)   (0,2)       (0,3)   (0,4)                 (0,5)         (0,6)   (0,7)   (0,8)
'Name   Dose    MaxDose     Route   SpecialInstructions   Timing        Units   Label   ShowBox
'String Double  Double      String  String                Array of Int  String  String  Boolean
Dim i As Long
Dim j As Long
Dim chktemp As MSForms.CheckBox
Dim colorbck As OLE_COLOR
Dim colorfrnt As OLE_COLOR
Dim curtop As Long
Dim numRows As Long
Dim startCell As Range
Dim curCell As Range
Dim tempArr() As String
Dim Key As String
Dim chkName As String

On Error GoTo ErrorHandle

colorbck = &HCEB63C     'Set background color
colorfrnt = &H593300       'Set foreground color
curtop = startTop             'Set top of first checkbox

'Iterate through the first category - the anti-emetics
Set startCell = Worksheets.Item("Backend").Range(sAddress)
numRows = Worksheets.Item("Backend").Range(sAddress, Range(sAddress).End(xlDown)).Rows.Count

For i = 0 To numRows - 1
    Set curCell = startCell.offset(i, 0)                        'Advance to the current iterated row
    
    With curCell
        
        Key = .Value & .offset(0, 1).Value
        PreMeds.Add Key, Key

        With PreMeds.Item(Key)
            .sName = curCell.Value                               'Set the drug name, in the first column
            .sDoseUnits = curCell.offset(0, 6).Value             'Dose units are needed to create calculated dose
            .dubMaxDose = curCell.offset(0, 2).Value             'Set the maximum dose; also needed to create calculated dose
            .dubRoundTo = curCell.offset(0, 10).Value             'Set what to round to; also needed to create calculated dose
            .dubDose = curCell.offset(0, 1).Value                'Set the dose
            .sRoute = curCell.offset(0, 3).Value                 'Set the administration route
            .sSpecInstruct = curCell.offset(0, 4).Value          'Set the special instructions
            .iMaxFreq = curCell.offset(0, 9).Value               'Maximum dosing frequency per week
        End With
        
        'We only have the timing property because CERTAIN premeds (Emend) are given on multiple days at different doses
        'We can't get away with just including this in special instructions because it is dosed by weight
        'and we have to be able to calculate that dose.  We could probably explicitly code around Emend but this
        'is a more flexible approach that might make life easier in the future if there is another premed with similar properties
        tempArr = Split(.offset(0, 5), ",")
        For j = LBound(tempArr) To UBound(tempArr)
            PreMeds.Item(Key).appTiming (CInt(tempArr(j)))
        Next j
    
        'Now we're going to create Checkboxes, first checking our boolean to see if an entry has a checkbox
        'If an entry does NOT have a checkbox, then it is a "linked" entry, meaning it takes the state of the entry above it
        '(so, one checkbox controls two objects in essence)
        If .offset(0, 8).Value Then
            chkName = "chk" & Key                             'Generate the name of the checkbox and assign it
            PreMeds.Item(Key).sChkBox = chkName
            
            Set chktemp = Me.Controls.Add("Forms.CheckBox.1", chkName)
            'With statements are weird, but essentially placing one inside another hides the scope of the outer one
            With chktemp
                .Top = curtop
                .Left = startLeft
                .Font.Name = "Gil Sans MT"
                .Font.Size = 12
                .BackColor = colorbck
                .ForeColor = colorfrnt
                .AutoSize = True
                .WordWrap = True
                .Width = 290
                .Caption = curCell.offset(0, 7).Value
                .AutoSize = False
                .Width = 290
                curtop = curtop + .Height + 5
            End With
        Else
            'If we're not displaying a checkbox, we ASSUME that we are linked to closest entry above
            'that is displaying a checkbox
            If Len(chkName) = 0 Then
                MsgBox "You must display at least one checkbox before creating linked entries that do not" _
                & "display checkboxes.  Your first row of premeds in any category MUST display a checkbox."
                GoTo BeforeExit
            Else
                PreMeds.Item(Key).sChkBox = chkName
                PreMeds.Item(mid(chkName, 4)).sLinkedKey = Key
            End If
        End If
        
    
    End With
Next i

BeforeExit:
Set startCell = Nothing
Set curCell = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & " in readFillPreMeds, PreMedications Form"
Resume BeforeExit
End Function
