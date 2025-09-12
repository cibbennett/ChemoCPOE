VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'A class holding the information for each chemotherapy drug
Option Explicit

Public Key As String                    'Key
Public sKey As String
Private mvarDrugName As String          'Name of the drug
Private mvarLocation As Integer         'Location where drug is administered; 0 = Inpatient, 1 = Outpatient, 2 = Home
Private mvarDoseUnits As Integer        'Dosing Units
Private mvarDose As Double              'Dose of the medication
Private mvarCalcDose As Double          'Calculated dose using dosing units
Private mvarCalcDoseUnits As String     'Units for calculated dose
Private mvarRoute As String             'Route of Administration
Private mvarInstructions As String      'Administration instructions
Private mvarDoseUnitsName               'Holds the dosing units
Private mvarWeeks() As Variant          'Jagged array organized as mvarWeeks(weeks)(days)
Private mvarDates() As Date             'Array of dates on which the drug is given
Private mvarStandby As Boolean          'Whether or not the drug has standby meds associated with it
Private mobjParent As clDrugs           'Parent class collection
'******************
'Class Initialize
'******************
Private Sub Class_Initialize()
    'Initialize all of our dynamic arrays
    ReDim mvarWeeks(1 To 24)
    ReDim mvarDates(0)
End Sub
Private Sub Class_Terminate()
    Erase mvarWeeks, mvarDates
End Sub
'******************
'Class Functions
'******************
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
    MsgBox err.Description & " in InsertDate, clDrug"
End Sub
Public Sub RemoveDay(index As Integer, week As Integer)
    Dim tmpDay() As Variant
    
    On Error GoTo ErrorHandle
    tmpDay = mvarWeeks(week)
    
    Call RemoveArrIndex(index, tmpDay)
    
    mvarWeeks(week) = tmpDay
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in RemoveDay, clDrug"
End Sub
Public Sub CalcDoses()
    'Calcuates dose to administer from dose/BSA or dose/weight
    'Special dosing rule for vincristine is hardcoded in:
        ' If dose is 1.4 mg/m2 (Hodgkins protocol), max vincristine dose is 2.8 mg
        ' All other cases max vincristine dose is 2 mg
    Dim tmp() As String
    Dim inte As Integer
    Dim s As String
        
    Select Case mvarDoseUnits
        Case 5 To 9                 'Dosing in units/kg
            Me.dubCalcDose = Me.dubDose * Information.dubWeight
            tmp = Split(MasterDoseList(mvarDoseUnits), "/")
            Me.sCalcDoseUnits = tmp(0)
        Case 0 To 4                 'Dosing in units/m2
            Me.dubCalcDose = Me.dubDose * Information.dubBSA
            tmp = Split(MasterDoseList(mvarDoseUnits), "/")
            Me.sCalcDoseUnits = tmp(0)
        Case 10 To 14               'Dosing in fixed units
            Me.dubCalcDose = Me.dubDose
            Me.sCalcDoseUnits = MasterDoseList(mvarDoseUnits)
    End Select
    
    'IU and U are prohibited abbreviations
    If Me.sCalcDoseUnits = "IU" Or Me.sCalcDoseUnits = "U" Then
        Me.sCalcDoseUnits = "Units"
    End If
    
    With Application.WorksheetFunction
        'If greater than 10, round to 2 sig figs, else round to 3
        If Me.dubCalcDose > 10 Then
            Me.dubCalcDose = .Round(Me.dubCalcDose, 2 - (1 + Int(.Log10(Abs(Me.dubCalcDose)))))
        Else
            Me.dubCalcDose = .Round(Me.dubCalcDose, 3 - (1 + Int(.Log10(Abs(Me.dubCalcDose)))))
        End If
    End With
        
    'Test to see if drug is vincristine; has to be spelled right
    s = Me.sDrugName
    s = Trim(s)
    s = StrConv(s, vbLowerCase)
    If StrComp(s, "vincristine", vbTextCompare) = 0 Then
        If Me.dubDose = 1.4 And Me.dubCalcDose > 2.8 Then
            Me.dubCalcDose = 2.8
        ElseIf Me.dubDose <> 1.4 And Me.dubCalcDose > 2 Then
            Me.dubCalcDose = 2
        End If
    End If
    
    
End Sub
'******************
'Class Properties
'******************
Public Property Get sDrugName() As String
sDrugName = mvarDrugName
End Property
Public Property Let sDrugName(ByVal vData As String)
mvarDrugName = vData
End Property
Public Property Get iLocation() As Integer
iLocation = mvarLocation
End Property
Public Property Let iLocation(ByVal vData As Integer)
mvarLocation = vData
End Property
Public Property Get iDoseUnits() As Integer
iDoseUnits = mvarDoseUnits
End Property
Public Property Let iDoseUnits(ByVal vData As Integer)
mvarDoseUnits = vData
Me.sDoseUnits = MasterDoseList(mvarDoseUnits)
End Property
Public Property Get sDoseUnits() As String
sDoseUnits = mvarDoseUnitsName
End Property
Public Property Let sDoseUnits(ByVal vData As String)
mvarDoseUnitsName = vData
End Property
Public Property Get dubDose() As Single
dubDose = mvarDose
End Property
Public Property Let dubDose(ByVal vData As Single)
mvarDose = vData
End Property
Public Property Get dubCalcDose() As Double
    dubCalcDose = mvarCalcDose
End Property
Public Property Let dubCalcDose(ByVal vData As Double)
    mvarCalcDose = vData
End Property
Public Property Get sCalcDoseUnits() As String
    sCalcDoseUnits = mvarCalcDoseUnits
End Property
Public Property Let sCalcDoseUnits(ByVal vData As String)
    mvarCalcDoseUnits = vData
End Property
Public Property Get sRoute() As String
sRoute = mvarRoute
End Property
Public Property Let sRoute(ByVal vData As String)
mvarRoute = vData
End Property
Public Property Get sInstructions() As String
sInstructions = mvarInstructions
End Property
Public Property Let sInstructions(ByVal vData As String)
mvarInstructions = vData
End Property
Public Property Get bWeek(week As Integer) As Boolean
    bWeek = Not IsEmpty(mvarWeeks(week))
End Property
Public Property Get iDay(index As Integer, week As Integer) As Integer
    iDay = mvarWeeks(week)(index)
End Property
Public Property Get numDays(week As Integer) As Integer
    If Not isArrayAllocated(mvarWeeks(week)) Then        'check if array is initialized or not
        numDays = 0                                      'if it isn't initialized, return 0
    Else
        numDays = ArrayLen(mvarWeeks(week))
    End If
End Property
Public Property Get datDate(index As Integer) As Date
datDate = mvarDates(index)
End Property
Public Property Let datDate(index As Integer, ByVal vData As Date)
mvarDates(index) = vData
End Property
Public Property Get numDates() As Integer
    numDates = ArrayLen(mvarDates)
End Property
Public Property Get linIndex(week As Integer, day As Integer) As Integer
    linIndex = LinearIndex(mvarWeeks, week, day)
End Property
Public Property Get bStandby() As Boolean
    bStandby = mvarStandby
End Property
Public Property Let bStandby(ByVal vData As Boolean)
    mvarStandby = vData
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property




