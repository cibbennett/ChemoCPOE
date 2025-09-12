VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clOrderSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolOrderSheets As Collection

'**********************
' Initialize/Terminate
'**********************
Private Sub Class_Initialize()
    Set mcolOrderSheets = New Collection
    
    'Add sheets for all weeks
    FindWeeks
    
    'Put them in order
    SortWeeks
    
    'Fill in each week with data for drugs, premeds, and labs
    FillWeeks
    
End Sub
Private Sub Class_Terminate()
    Set mcolOrderSheets = Nothing
End Sub
'**********************
' Add/Remove Utilities
'**********************
Public Sub Remove(vntIndexKey As Variant)
    'Can be used to remove a drug from the collection
    mcolOrderSheets.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clOrderSheet
'Adds a clOrderSheet class to our class collection
Dim objNewMember As clOrderSheet
On Error GoTo ErrorHandle

'Create a new object, of type clDrug
Set objNewMember = New clOrderSheet

objNewMember.Key = Key

If Len(sKey) = 0 Then
    mcolOrderSheets.Add objNewMember
Else
    mcolOrderSheets.Add objNewMember, sKey
End If
    
Set mcolOrderSheets.Item(sKey).Parent = Me

'Destroy the standalone object as we have added it to the collection
Set objNewMember = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clOrderSheets"
End Function
Public Function ContainsWeek(vntIndexKey As Variant) As Boolean
    ContainsWeek = Contains(mcolOrderSheets, vntIndexKey)
End Function
'**********************
' Class Properties
'**********************
Public Property Get Count() As Integer
    Count = mcolOrderSheets.Count
End Property
Public Property Get Item(vntIndexKey As Variant) As clOrderSheet
    Set Item = mcolOrderSheets(vntIndexKey)
End Property
'***********************
' Class Functions
'***********************
Private Sub FindWeeks()
    'Find all populated weeks and create new ordersheet entries for each one
    Dim i As Integer, j As Integer
    Dim week As String
    
    With Me
        For j = 1 To Drugs.Count
            For i = 1 To 24
                If Drugs.Item(j).bWeek(i) Then
                    week = "Week" & i
                    If Not .ContainsWeek(week) Then
                        .Add week, week
                        .Item(week).datWeekStart = startDate + (i - 1) * 7
                        .Item(week).iWeekNum = i
                    End If
                End If
            Next i
         Next j
    End With
End Sub
Private Sub SortWeeks()
    'Places weeks in order in the collection
    'Bubble sort has identical performance to the far more complex mergesort for 25 or fewer items
    
    Dim i As Integer, j As Integer
    Dim vTemp As clOrderSheet
    
    For i = 1 To mcolOrderSheets.Count - 1
        For j = i + 1 To mcolOrderSheets.Count
            If mcolOrderSheets.Item(i).iWeekNum > mcolOrderSheets.Item(j).iWeekNum Then
                Set vTemp = mcolOrderSheets.Item(j)
                mcolOrderSheets.Remove (j)
                mcolOrderSheets.Add vTemp, vTemp.Key, i
            End If
        Next j
    Next i

End Sub
Private Sub FillWeeks()
    'Fills in the properties for each order sheet (which can cover up to a week at a time)
    Dim i As Integer, j As Integer, k As Integer
    Dim cycleDay As Integer
    
    On Error GoTo ErrorHandle
    
    For i = 1 To Me.Count
        With Me.Item(i)
        
            'Add the drugs for this week first
            For j = 1 To Drugs.Count
                If Drugs.Item(j).bWeek(.iWeekNum) Then
                    .AddDrug Drugs.Item(j)
                    If Drugs.Item(j).bStandby Then
                        .InsertStandby Drugs.Item(j).sDrugName
                    End If
                    For k = 0 To Drugs.Item(j).numDays(.iWeekNum) - 1
                        cycleDay = ((.iWeekNum - 1) * 7) + Drugs.Item(j).iDay(k, .iWeekNum)
                        .InsertDay cycleDay
                    Next k
                End If
            Next j
            
            'Next add premeds
            For j = 1 To PreMeds.Count
                If PreMeds.Item(j).bWeek(.iWeekNum) Then
                    .AddPreMed PreMeds.Item(j)
                End If
            Next j
            
            'Now labs
            For j = 1 To Labs.Count
                If Labs.Item(j).bWeek(.iWeekNum) Then
                    .AddLab Labs.Item(j)
                End If
            Next j
            
        End With
    Next i
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in FillWeeks, clOrderSheets."
End Sub
