VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clAdminDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'A Collection of AdminDay classes
Option Explicit

Public Key As String                        'Key
Private mcolAdminDays As Collection         'Collection of clAdminDays Classes
Private Sub Class_Initialize()
On Error GoTo ErrorHandle

Set mcolAdminDays = New Collection

ThisWorkbook.Activate

'Add dates to the collection & set their properties
FindDates

'Put collection mcolAdminDays in date-order, will make life easier in the moveDates subroutine
sortDates

'Ensures that outpatient days do not fall on weekends
moveDates

'Now update the dates in the drug class objects
updateDrugs

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure Class_Initialize, clAdminDays"
End Sub
Private Sub Class_Terminate()
    Set mcolAdminDays = Nothing
End Sub
'*************************
' Add/Remove Utilities
'*************************
Public Sub Remove(vntIndexKey As Variant)
    'Can be used to remove a drug from the collection
    mcolAdminDays.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clAdminDay
'Adds a clAdminDay class to our class collection
Dim objNewMember As clAdminDay
On Error GoTo ErrorHandle

'Create a new object, of type clDrug
Set objNewMember = New clAdminDay

objNewMember.Key = Key

If Len(sKey) = 0 Then
    mcolAdminDays.Add objNewMember
Else
    mcolAdminDays.Add objNewMember, sKey
End If
    
Set mcolAdminDays.Item(sKey).Parent = Me

'Destroy the standalone object as we have added it to the collection
Set objNewMember = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clAdminDays"
End Function
'********************
' Class Properties
'********************
Public Property Get Item(vntIndexKey As Variant) As clAdminDay
Set Item = mcolAdminDays(vntIndexKey)
End Property
Public Property Get Count() As Long
Count = mcolAdminDays.Count
End Property
'*******************
' Class Functions
'*******************
Private Sub FindDates()
'To create this collection, we're going to need to iterate through the class collection clDrugs and identify all the dates
'that each drug is given on; for each date, we're going to add a new clAdminDay to our collection and set its property
Dim xDrug As clDrug
Dim i As Integer
Dim j As Integer
Dim curDate As String
Dim testobj As clAdminDay


On Error GoTo ErrorHandle:

With Me
    For i = 1 To Drugs.Count                                    'Iterate through the collection of drugs
    Set xDrug = Nothing
    Set xDrug = Drugs.Item(i)
        For j = 0 To xDrug.numDates - 1                       'Iterate through the array of dates for the drug
            curDate = CStr(xDrug.datDate(j))
            If Not Contains(mcolAdminDays, curDate) Then        'Check to see if we've already created an AdminDay for this date
                .Add curDate, curDate                                   'If not, create one and begin to fill it in
                .Item(curDate).datDate = xDrug.datDate(j)
                .Item(curDate).sDayName = xDrug.sDrugName
                .Item(curDate).iDayType = xDrug.iLocation
            Else
                'If the day already exists, then we will concatenate the name of the day (which is the
                'name of the drugs given on that day) with the current drug, with a carriage return (Chr(10))
                'serving as the delimiter
                .Item(curDate).sDayName = .Item(curDate).sDayName & Chr(10) & xDrug.sDrugName
                
                'Change the day type based on where the drug we're adding is given:
                'If the day was previously home and the new drug is given inpatient or outpatient,
                'set it to that; If the day was previous outpatient and the new drug is given inpatient
                'then change the day to inpatient
                If .Item(curDate).iDayType > xDrug.iLocation Then
                    .Item(curDate).iDayType = xDrug.iLocation
                End If
            End If
            'Add the current drug to array of drugs for that date, helpful later once we start moving dates around
            .Item(curDate).appDrugs xDrug
        Next j
    Next
End With

BeforeExit:
Set xDrug = Nothing
Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure FindDats, clAdminDays"
Resume BeforeExit
End Sub
Private Sub sortDates()
    Dim vItm As clAdminDay
    Dim i As Long, j As Long
    Dim vTemp As clAdminDay


    'Two loops to bubble sort
    For i = 1 To mcolAdminDays.Count - 1
        For j = i + 1 To mcolAdminDays.Count
            If mcolAdminDays(i).datDate > mcolAdminDays(j).datDate Then
                'store the lesser item
                Set vTemp = mcolAdminDays(j)
                'remove the lesser item
                mcolAdminDays.Remove j
                're-add the lesser item before the
                'greater Item
                mcolAdminDays.Add vTemp, vTemp.Key, i
            End If
        Next j
    Next i
End Sub
Private Sub moveDates()
    'This subroutine identifies blocks of contiguous days and moves them (or not) based on the properties of the days contained in the block
    'In brief: if any outpatient chemo day falls on a weekend and is not contiguous with an inpatient day, it moves the entire block of days
    'and all following days by an appropriate amount
    Dim i As Long
    Dim j As Long
    Dim contig As Boolean
    Dim bBegin As Long
    Dim bEnd As Long
    Dim containsIP As Boolean
    Dim bBeginDOW As Integer
    Dim BEndDOW As Integer
    Dim moveDistance As Integer
    Dim xDrug As clDrug
    
    For i = 1 To Me.Count
        With Me.Item(i)
            If .iDayType = 1 And (Weekday(.datDate, vbMonday) = 6 Or Weekday(.datDate, vbMonday) = 7) Then   'If the day type is outpatient chemo and fall on a weekend
                j = i
                contig = True
                
                'Search backwards first to find the first day of the contiguous block
                Do While contig = True
                    If i = 1 Then
                        bBegin = i
                        contig = False
                    ElseIf j = 1 Then
                        bBegin = j
                        contig = False
                    'Check to see if there is a day right before, and make sure it isn't a home day (they don't count as
                    'part of the block
                    ElseIf Me.Item(j).datDate - Me.Item(j - 1).datDate = 1 And Me.Item(j - 1).iDayType <> 2 Then
                            j = j - 1
                    Else
                        bBegin = j
                        contig = False
                    End If
                Loop
                
                'Search Forwards to find the last day of the contiguous block
                contig = True
                j = i
                Do While contig = True
                    If i = Me.Count Then
                        bEnd = i
                        contig = False
                    ElseIf j = Me.Count Then
                        bEnd = j
                        contig = False
                    ElseIf Me.Item(j + 1).datDate - Me.Item(j).datDate = 1 And Me.Item(j + 1).iDayType <> 2 Then
                        j = j + 1
                    Else
                        bEnd = j
                        contig = False
                    End If
                Loop
                
                
                For j = bBegin To bEnd
                    'If there is one inpatient day in the block, they all become inpatient
                    If Me.Item(j).iDayType = 0 Then
                        containsIP = True
                    End If
                Next
                
                'If block size is > 5 days or there is one inpatient day in the block, it becomes inpatient
                If containsIP Or (bEnd - bBegin > 5) Then
                    For j = bBegin To bEnd
                        Me.Item(j).iDayType = 0
                    Next
                    GoTo Proceed
                End If
                
                'Determine distance to move dates forward or back
                bBeginDOW = Weekday(Me.Item(bBegin).datDate, vbMonday)
                BEndDOW = Weekday(Me.Item(bEnd).datDate, vbMonday)
            
                Select Case bBeginDOW
                    Case Is = 1:
                        GoTo Proceed
                    Case Is = 2:
                        moveDistance = -1
                    Case Is = 3:
                        If BEndDOW = 6 Then
                            moveDistance = -1
                        Else
                            moveDistance = 5
                        End If
                    Case Is = 4:
                        If BEndDOW = 6 Then
                            moveDistance = -1
                        Else
                            moveDistance = 4
                        End If
                    Case Is = 5:
                        If BEndDOW = 6 Then
                            moveDistance = -1
                        Else
                            moveDistance = 3
                        End If
                    Case Is = 6:
                        If BEndDOW = 6 Then
                            moveDistance = -1
                        Else
                            moveDistance = 2
                        End If
                    Case Is = 7:
                        moveDistance = 1
                End Select
                
                'Move the dates
                For j = bBegin To Me.Count
                    Me.Item(j).datDate = Me.Item(j).datDate + moveDistance
                Next
                
                'Set the begin and end date for every block member
                For j = bBegin To bEnd
                    Me.Item(j).datBlockBegin = Me.Item(bBegin).datDate
                    Me.Item(j).datBlockEnd = Me.Item(bEnd).datDate
                Next
            End If
Proceed:
        End With
    Next
    
Exit Sub
ErrorHandle:
MsgBox err.Description & "Procedure moveDates, clAdminDays"
End Sub
Private Sub updateDrugs()
    'Once we have applied all the logic to move our days, we need to update
    'the date found in the original drug classes
    Dim i As Integer
    Dim j As Integer
    Dim xDrug As clDrug
    Dim xAdminDay As clAdminDay
    
    On Error GoTo ErrorHandle
    
    For Each xAdminDay In mcolAdminDays
        If Not xAdminDay.datDate = DateValue(xAdminDay.Key) Then                'The key still contains the original date
            For i = 0 To xAdminDay.DrugsLength - 1                              'So we can use it to see if the date changed
                Set xDrug = xAdminDay.objDrugs(i)                               'If it did, we can access the class object for the drugs
                For j = 0 To xDrug.numDates - 1                               'given on that day and update their dates
                    If xDrug.datDate(j) = DateValue(xAdminDay.Key) Then
                        xDrug.datDate(j) = xAdminDay.datDate
                    End If
                Next j
            Next i
        End If
    Next
    
BeforeExit:
    Set xDrug = Nothing
    Set xAdminDay = Nothing
    Exit Sub
ErrorHandle:
MsgBox err.Description & "Procedure updateDrugs, clAdminDays"
GoTo BeforeExit
End Sub
