VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDrugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'A collection of drug classes.
Option Explicit

Public Key As String                    'Key
Public sKey As String
Private mcolDrugs As Collection         'Collection of clDrug Classes
Private Sub Class_Initialize()
'Runs when the class is initialized/create

On Error GoTo ErrorHandle

'Create a new class collection of drugs
Set mcolDrugs = New Collection

ThisWorkbook.Activate

'Add drugs as classes to the collection & sets their properties
FindDrugs


Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure Class_Initialize, clDrugs"
End Sub
Private Sub Class_Terminate()
    Set mcolDrugs = Nothing
End Sub
'*************************
' Add/Remove Utilities
'*************************
Public Sub Remove(vntIndexKey As Variant)
    'Can be used to remove a drug from the collection
    mcolDrugs.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clDrug
'Adds a clDrug class to our class collection
Dim objNewMember As clDrug
On Error GoTo ErrorHandle

'Create a new object, of type clDrug
Set objNewMember = New clDrug

objNewMember.Key = Key


If Len(sKey) = 0 Then
    mcolDrugs.Add objNewMember
Else
    mcolDrugs.Add objNewMember, sKey
End If
    
Set mcolDrugs.Item(sKey).Parent = Me

'Destroy the standalone object as we have added it to the collection
Set objNewMember = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clDrugs"
End Function
'********************
' Class Properties
'********************
Public Property Get Item(vntIndexKey As Variant) As clDrug
Set Item = mcolDrugs(vntIndexKey)
End Property
Public Property Get Count() As Long
Count = mcolDrugs.Count
End Property
'*******************
' Class Functions
'*******************
Private Sub FindDrugs()
'This will read from our GUI sheet to gather drug information and fill out the class properties
Dim r As Integer, i As Integer, j As Integer, k As Integer, l As Integer, cDay As Integer, cWeek As Integer
Dim ind As Integer          'index of the drug we're adding
Dim populated As Boolean    'Does the drug have any weeks assigned?
Dim c1Name As String
Dim c2Name As String
Dim nDate As Date
Dim startingCell As Range
Dim xDrug As clDrug

On Error GoTo ErrorHandle

Worksheets.Item("Creator").Activate                 'Activate the GUI sheet; it should already be active but better to be safe
Set startingCell = ActiveSheet.Range("DrugStart")   'Set the location of the first cell in the drug table

ind = 1

With Me
    'Iterate through the rows of drugs in the UI
    For r = 1 To 10
        If Len(startingCell.offset(r - 1, 0).Value) > 0 Then                    'If there's a drug name present, proceed
            populated = False                                                   'Reset our populated flag
            
            .Add "Drug" & ind, "Drug" & ind                                     'First Add a drug to the class
            
            'Enter the days first, if there are no days assigned we will exit
            If bDaysEntry Then
                For i = 1 To 3
                    For j = 1 To 30
                        c1Name = "d" & r & "dd" & ((i - 1) * 30) + j
                         If Worksheets("controlstates").Range(c1Name).Value Then
                            If ((((i - 1) * 30) + j) Mod 7) = 0 Then
                                cDay = 7
                                cWeek = (((i - 1) * 30) + j) \ 7
                            Else
                                cDay = (((i - 1) * 30) + j) Mod 7
                                cWeek = ((((i - 1) * 30) + j) \ 7) + 1
                            End If
                            .Item(ind).InsertDay cDay, cWeek
                         populated = True
                         End If
                    Next j
                Next i
            Else
                For i = 1 To 2                                                      'Which row of weeks
                    For j = 1 To 12                                                 'Which week
                        c1Name = "d" & r & "w" & ((i - 1) * 12) + j                 'Name of the week cell
                        'If there is chemo on this week, check the days
                        If Worksheets("controlstates").Range(c1Name).Value Then
                            For k = 1 To 7                                              'Which Day
                                c2Name = "d" & r & "d" & k                              'Name of the Day cell
                                'If the week and the day are true, then add the day
                                If Worksheets("controlstates").Range(c2Name).Value Then
                                    .Item(ind).InsertDay k, ((i - 1) * 12) + j
                                    populated = True
                                End If
                            Next k
                        End If
                    Next j
                Next i
            End If
    
            'If there were no days added to the weeks array, then delete the item and move on
            If Not populated Then
                mcolDrugs.Remove (ind)
                GoTo nextloop
            End If
            
            'If we did add at least one day, then fill out the rest of the drug information
            .Item(ind).sDrugName = startingCell.offset(r - 1, 0).Value
            .Item(ind).dubDose = startingCell.offset(r - 1, 1).Value
            .Item(ind).sRoute = startingCell.offset(r - 1, 3).Value
            .Item(ind).sInstructions = startingCell.offset(r - 1, 4).Value
            
            'Now we're going to read from our active x controls, which have been named to correlate with the row they are
            'Comboboxes are done first, the first row contains combobox 1 & 11, the second 2 & 12, etc
            c1Name = "ComboBox" & r
            c2Name = "ComboBox" & (r + 10)
            .Item(ind).iDoseUnits = ActiveSheet.OLEObjects(c1Name).Object.ListIndex 'Yikes
            .Item(ind).iLocation = ActiveSheet.OLEObjects(c2Name).Object.ListIndex
            
            'We have to calculate the dates based on the weeks, days, and our starting date, then write them into our class array
            'We check to make to sure the weeks array is nonzero first
            For i = 1 To 24
                If .Item(ind).bWeek(i) Then
                    For j = 0 To .Item(ind).numDays(i) - 1
                        nDate = startDate + ((i - 1) * 7) + .Item(ind).iDay(j, i) - 1
                        .Item(ind).InsertDate nDate
                    Next j
                End If
            Next i
            
            'Finally, we generated the calculated doses of the drugs based on the dosing units and the entered dose
            .Item(ind).CalcDoses
            
            'Increment our index
            ind = ind + 1
        End If
nextloop:
    Next
End With

BeforeExit:
Set startingCell = Nothing
Set xDrug = Nothing
Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure FindDrugs, clDrugs"
Debug.Print "r = " & r & ", ind = " & ind
GoTo BeforeExit
End Sub
