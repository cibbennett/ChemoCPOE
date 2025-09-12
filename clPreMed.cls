VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clPreMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class for holding information about premedications that is used to write the actual order sheet
Public Key As String            'Key
Public sKey As String           'Secondary Key
Private mvarName As String      'Medication Name
Private mvarRoute As String     'Medication Route
Private mvarDose As Double      'Medication Dose
Private mvarMaxDose As Double   'Maximum Medication Dose
Private mvarMaxFreq As Integer  'Maximum doses per week
Private mvarRoundTo As Double   'What to round to
Private mvarDoseUnits As String 'Dose Units
Private mvarCalcDose As Double  'Calculated Dose
Private mvarCalcDoseUnits As String 'Units for calculated dose
Private mvarSpecInstruct As String  'Special Instructions
Private mvarChkBox As String    'Name of the associated checkbox
Private mvarLinkedKey As String  'Key for any linked premed that doesn't have its own checkbox
Private mvarTiming() As Integer  'An array holding the timing of premed doses as offsets from the admin day
Private mvarWeeks() As Variant   'Jagged array organized as mvarWeeks(weeks)(days)
Private mvarDates() As Date      'Dates on which its done
Private mvarDrugKeys() As String    'Keys for the associated drugs
Private mobjParent As clPreMeds    'Parent Object
'**************
' Initialize
'**************
Private Sub Class_Initialize()
    ReDim mvarWeeks(1 To 24)
    ReDim mvarDrugKeys(0)
    ReDim mvarDates(0)
End Sub
Private Sub Class_Terminate()
    Erase mvarTiming, mvarDrugKeys, mvarWeeks, mvarDates
End Sub
'**************
' Functions
'**************
Public Sub appTiming(ByVal vData As Integer)
    'Check to see if the array has been initialized...
    If (Not Not mvarTiming) = 0 Then
        ReDim Preserve mvarTiming(0 To 0)
        mvarTiming(0) = vData
    Else
        ReDim Preserve mvarTiming(0 To UBound(mvarTiming) + 1)
        mvarTiming(UBound(mvarTiming)) = vData
    End If
End Sub
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
    MsgBox err.Description & " in InsertDrug, clPreMed"
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
    MsgBox err.Description & " in InsertDate, clPreMed"
End Sub
Public Sub RemoveDay(index As Integer, week As Integer)
    Dim tmpDay() As Variant
    
    On Error GoTo ErrorHandle
    tmpDay = mvarWeeks(week)
    
    Call RemoveArrIndex(index, tmpDay)
    
    mvarWeeks(week) = tmpDay
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in RemoveDay, clPremed"
End Sub
Public Sub RemoveDate(index As Integer)

    On Error GoTo ErrorHandle
    
    Call datRemoveArrIndex(index, mvarDates)
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in RemoveDate, clPremed"
End Sub
'**************
' Properties
'**************
Public Property Get sName() As String
    sName = mvarName
End Property
Public Property Let sName(ByVal vData As String)
    mvarName = vData
End Property
Public Property Get sRoute() As String
    sRoute = mvarRoute
End Property
Public Property Let sRoute(ByVal vData As String)
    mvarRoute = vData
End Property
Public Property Get dubDose() As Double
    dubDose = mvarDose
End Property
Public Property Let dubDose(ByVal vData As Double)
    mvarDose = vData
    
    Select Case Me.sDoseUnits
        Case Is = "ml/m2/hr":
            Me.dubCalcDose = mvarDose * Information.dubBSA
            Me.sCalcDoseUnits = "ml/hr"
        Case Is = "mg/kg"
            Me.dubCalcDose = mvarDose * Information.dubWeight
            If Me.dubCalcDose > dubMaxDose Then
                Me.dubCalcDose = dubMaxDose
            End If
            Me.sCalcDoseUnits = "mg"
        Case Is = "g/kg"
            Me.dubCalcDose = mvarDose * Information.dubWeight
            If Me.dubCalcDose > dubMaxDose Then
                Me.dubCalcDose = dubMaxDose
            End If
            Me.sCalcDoseUnits = "g"
        Case Is = "ml/kg/hr"
            Me.dubCalcDose = mvarDose * Information.dubWeight
            Me.sCalcDoseUnits = "ml/hr"
        Case Is = "mg/m2"
            Me.dubCalcDose = mvarDose * Information.dubBSA
            If Me.dubCalcDose > dubMaxDose Then
                Me.dubCalcDose = dubMaxDose
            End If
            Me.sCalcDoseUnits = "mg"
        Case Else
            Me.dubCalcDose = mvarDose
            Me.sCalcDoseUnits = Me.sDoseUnits
    End Select
    Me.dubCalcDose = Round(Me.dubCalcDose / Me.dubRoundTo, 0) * Me.dubRoundTo 'round the dose to the unit specified
End Property
Public Property Get sDoseUnits() As String
    sDoseUnits = mvarDoseUnits
End Property
Public Property Let sDoseUnits(ByVal vData As String)
    mvarDoseUnits = vData
End Property
Public Property Get sCalcDoseUnits() As String
    sCalcDoseUnits = mvarCalcDoseUnits
End Property
Public Property Let sCalcDoseUnits(ByVal vData As String)
    mvarCalcDoseUnits = vData
End Property
Public Property Get dubCalcDose() As Double
    dubCalcDose = mvarCalcDose
End Property
Public Property Let dubCalcDose(ByVal vData As Double)
    mvarCalcDose = vData
End Property
Public Property Get iMaxFreq() As Integer
    iMaxFreq = mvarMaxFreq
End Property
Public Property Let iMaxFreq(ByVal vData As Integer)
    mvarMaxFreq = vData
End Property
Public Property Get dubRoundTo() As Double
    dubRoundTo = mvarRoundTo
End Property
Public Property Let dubRoundTo(ByVal vData As Double)
    mvarRoundTo = vData
End Property
Public Property Get sSpecInstruct() As String
    sSpecInstruct = mvarSpecInstruct
End Property
Public Property Let sSpecInstruct(ByVal vData As String)
    mvarSpecInstruct = vData
End Property
Public Property Get sChkBox() As String
    sChkBox = mvarChkBox
End Property
Public Property Let sChkBox(ByVal vData As String)
    mvarChkBox = vData
End Property
Public Property Get sLinkedKey() As String
    sLinkedKey = mvarLinkedKey
End Property
Public Property Let sLinkedKey(ByVal vData As String)
    mvarLinkedKey = vData
End Property
Public Property Get dubMaxDose() As Double
    dubMaxDose = mvarMaxDose
End Property
Public Property Let dubMaxDose(ByVal vData As Double)
    mvarMaxDose = vData
End Property
Public Property Get iTiming(index As Variant) As Integer
    iTiming = mvarTiming(index)
End Property
Public Property Get numTiming() As Integer
    numTiming = ArrayLen(mvarTiming)
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
Public Property Get bWeek(week As Integer) As Boolean
    bWeek = Not IsEmpty(mvarWeeks(week))
End Property
Public Property Get sDrug(index As Integer) As String
    sDrug = mvarDrugKeys(index)
End Property
Public Property Get linIndex(week As Integer, day As Integer) As Integer
    linIndex = LinearIndex(mvarWeeks, week, day)
End Property
Public Property Get numDrugs() As Integer
    numDrugs = ArrayLen(mvarDrugKeys)
End Property
Public Property Get containsDay(ByVal day As Integer, ByVal week As Integer) As Integer
    'Returns the index of a day if it is in the specified week, else returns -1
    Dim i As Integer
    Dim match As Integer
    
    match = -1
    
    For i = 0 To Me.numDays(week) - 1
        If Me.iDay(i, week) = day Then
            match = i
        End If
    Next i
    
    containsDay = match
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property

