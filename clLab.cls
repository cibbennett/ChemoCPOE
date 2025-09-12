VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class holding all labs to be done with associated chemotherapy

Public Key As String                'Key
Public sKey As String               'Secondary Key
Private mvarName As String          'Name of the lab
Private mvarWeeks() As Variant
Private mvarDates() As Date         'Dates on which its done
Private mvarDrugKeys() As String    'Keys for the associated drugs
Private mobjParent As clLabs        'Parent

'*************************
' Initialize & Terminate
'*************************
Private Sub Class_Initialize()
    'Initialize our arrays
    ReDim mvarDays(0)
    ReDim mvarWeeks(1 To 24) 'I really did not want to hardcode the weeks like this, but short of doing an array of arrays of arrays
    ReDim mvarDrugKeys(0)    'it was really the only option
    ReDim mvarDates(0)
End Sub
Private Sub Class_Terminate()
    Erase mvarWeeks, mvarDates, mvarDrugKeys
End Sub
'************************
' Class Functions
'************************
Public Sub InsertDay(ByVal vDay As Integer, ByVal vWeek As Integer)
    Dim tmpDay() As Variant
    
    On Error GoTo ErrorHandle
    
    If IsEmpty(mvarWeeks(vWeek)) Then
        ReDim tmpDay(0)
        tmpDay(0) = vDay
        mvarWeeks(vWeek) = tmpDay
    Else
        tmpDay = mvarWeeks(vWeek)
        InsertSort tmpDay, vDay
        mvarWeeks(vWeek) = tmpDay
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " InsertDayWeek, clLab"
End Sub
Public Sub InsertDrug(ByVal vData As String)
    'Appends a drug key to the array
     On Error GoTo ErrorHandle
     
    If mvarDrugKeys(0) = "" Then
        mvarDrugKeys(0) = vData
    Else
        ReDim Preserve mvarDrugKeys(UBound(mvarDrugKeys) + 1)
        mvarDrugKeys(UBound(mvarDrugKeys)) = vData
    End If
     
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in InsertDrug, clLab"
End Sub
Public Sub InsertDate(ByVal vData As Date)
    'Inserts a date into the array in order
    On Error GoTo ErrorHandle
    
    If mvarDates(0) = 0 Then
        mvarDates(0) = vData
    Else
        Call datInsertSort(mvarDates, vData)
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in InsertDate, clLab"
End Sub
'************************
' Class Properties
'************************
Public Property Get sName() As String
    sName = mvarName
End Property
Public Property Let sName(ByVal vData As String)
    mvarName = vData
End Property
Public Property Get iDay(index As Integer, week As Integer) As Integer
    iDay = mvarWeeks(week)(index)
End Property
Public Property Get numDays(week As Integer) As Integer
    numDays = ArrayLen(mvarWeeks(week))
End Property
Public Property Get sDrug(index As Integer) As String
    sDrug = mvarDrugKeys(index)
End Property
Public Property Get numDrugs() As Integer
    numDrugs = ArrayLen(mvarDrugKeys)
End Property
Public Property Get datDate(index As Integer) As Date
    datDate = mvarDates(index)
End Property
Public Property Get bWeek(week As Integer) As Boolean
    bWeek = Not IsEmpty(mvarWeeks(week))
End Property
Public Property Get numDates() As Integer
    numDates = ArrayLen(mvarDates)
End Property
Public Property Get linIndex(week As Integer, day As Integer) As Integer
    linIndex = LinearIndex(mvarWeeks, week, day)
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property
