VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clOrderSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Similar to the calendar, we need to compile all of our information before writing it to an order sheet
'Each object in this class will correspond to one order sheet, which is one week worth of treatment

Public Key As String                'Key
Private mvarWeekStart As Date       'Starting day of this week of treatment
Private mvarWeekNum As Integer      'Which week this is
Private mcolDrugs As Collection     'Collection of drugs given this week
Private mcolPreMeds As Collection   'Collection of PreMeds given this week
Private mcolLabs As Collection      'Collection of Labs given this week
Private mvarDays() As Integer       'Days in this week as related to the total treatment course
Private mvarStandby() As String     'Drugs that have standby medications in this week
Private mobjParent As clOrderSheets 'Parent

'**********************
' Initialize/Terminate
'**********************
Private Sub Class_Initialize()
    Set mcolDrugs = New Collection
    Set mcolPreMeds = New Collection
    Set mcolLabs = New Collection
    ReDim mvarStandby(0)
    ReDim mvarDays(0)
End Sub
Private Sub Class_Terminate()
    Set mcolDrugs = Nothing
    Set mcolPreMeds = Nothing
    Set mcolLabs = Nothing
End Sub
'**********************
' Class Functions
'**********************
Public Sub AddDrug(vData As clDrug)
    mcolDrugs.Add vData, vData.Key
End Sub
Public Sub AddPreMed(vData As clPreMed)
    mcolPreMeds.Add vData, vData.Key
End Sub
Public Sub AddLab(vData As clLab)
    mcolLabs.Add vData, vData.Key
End Sub
Public Sub InsertStandby(vData As String)
    If mvarStandby(0) = "" Then
        mvarStandby(0) = vData
    Else
        ReDim Preserve mvarStandby(UBound(mvarStandby) + 1)
        mvarStandby(UBound(mvarStandby)) = vData
    End If
End Sub
Public Sub InsertDay(vData As Integer)
    If mvarDays(0) = 0 Then
        mvarDays(0) = vData
    Else
        intInsertSort mvarDays, vData
    End If
End Sub
Public Sub RemoveDrug(vntIndexKey As Variant)
    mcolDrugs.Remove vntIndexKey
End Sub
Public Sub RemovePreMed(vntIndexKey As Variant)
    mcolPreMeds.Remove vntIndexKey
End Sub
Public Sub RemoveLab(vntIndexKey As Variant)
    mcolLabs.Remove vntIndexKey
End Sub
'**********************
' Class Properties
'**********************
Public Property Get CountDrugs() As Integer
    CountDrugs = mcolDrugs.Count
End Property
Public Property Get CountPreMeds() As Integer
    CountPreMeds = mcolPreMeds.Count
End Property
Public Property Get CountLabs() As Integer
    CountLabs = mcolLabs.Count
End Property
Public Property Get CountStandby() As Integer
    CountStandby = ArrayLen(mvarStandby)
End Property
Public Property Get CountDays() As Integer
    CountDays = ArrayLen(mvarDays)
End Property
Public Property Get itDrug(vntIndexKey As Variant) As clDrug
    Set itDrug = mcolDrugs(vntIndexKey)
End Property
Public Property Get itPreMed(vntIndexKey As Variant) As clPreMed
    Set itPreMed = mcolPreMeds(vntIndexKey)
End Property
Public Property Get itLab(vntIndexKey As Variant) As clLab
    Set itLab = mcolLabs(vntIndexKey)
End Property
Public Property Get sStandby(index As Integer) As String
    sStandby = mvarStandby(index)
End Property
Public Property Get iDay(index As Integer) As Integer
    iDay = mvarDays(index)
End Property
Public Property Get datWeekStart() As Date
    datWeekStart = mvarWeekStart
End Property
Public Property Let datWeekStart(ByVal vData As Date)
    mvarWeekStart = vData
End Property
Public Property Get iWeekNum() As Integer
    iWeekNum = mvarWeekNum
End Property
Public Property Let iWeekNum(ByVal vData As Integer)
    mvarWeekNum = vData
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property
