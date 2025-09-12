VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clAdminDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'A class that represents each day any drug is given on
Option Explicit

Public Key As String
Public sKey As String
Private mvarDayName As String            'Concatenated name of all drugs given on this day
Private mvarDayType As Integer           'Inpatient, outpatient, or home
Private mvarDay As Long                  'Day in the month the date falls on
Private mvarMonth As Long                'Month in the year the date falls on
Private mvarDate As Date                 'The date of the day
Private mvarBlockBegin As Date           'First day in a contiguous block that contains this date
Private mvarBlockEnd As Date             'Last day in a contiguous block that contains this date
Private mobjDrugs() As clDrug            'An array of the drugs that point
Private mobjParent As clAdminDays        'A pointer to the parent collection class
'******************
'Class Initialize
'******************
Private Sub Class_Initialize()
    'Initialize our dynamic array
    ReDim mobjDrugs(0)
End Sub
'******************
'Class Functions
'******************
Public Sub appDrugs(vData As Object)
    'Append function to add data to the end of an array
    'If arr(0) == 0, we're going to assume its a freshly initialized array because 0 is not valid data for this array
    If mobjDrugs(0) Is Nothing Then
        Set mobjDrugs(0) = vData
    Else
        ReDim Preserve mobjDrugs(0 To UBound(mobjDrugs) + 1)
        Set mobjDrugs(UBound(mobjDrugs)) = vData
    End If
End Sub
Public Function DrugsLength() As Integer
    DrugsLength = ArrayLen(mobjDrugs)
End Function
'******************
'Class Properties
'******************
Public Property Get sDayName() As String
    sDayName = mvarDayName
End Property
Public Property Let sDayName(ByVal vData As String)
    mvarDayName = vData
End Property
Public Property Get iDayType() As Integer
    iDayType = mvarDayType
End Property
Public Property Let iDayType(ByVal vData As Integer)
    mvarDayType = vData
End Property
Public Property Get datDate() As Date
    datDate = mvarDate
End Property
Public Property Let datDate(ByVal vData As Date)
    mvarDate = vData
    Me.lDay = day(datDate)
    Me.lMonth = Month(datDate)
End Property
Public Property Get lDay() As Long
    lDay = mvarDay
End Property
Public Property Let lDay(ByVal vData As Long)
    mvarDay = vData
End Property
Public Property Get lMonth() As Long
    lMonth = mvarMonth
End Property
Public Property Let lMonth(ByVal vData As Long)
    mvarMonth = vData
End Property
Public Property Get datBlockBegin() As Date
    datBlockBegin = mvarBlockBegin
End Property
Public Property Let datBlockBegin(ByVal vData As Date)
    mvarBlockBegin = vData
End Property
Public Property Get datBlockEnd() As Date
    datBlockEnd = mvarBlockEnd
End Property
Public Property Let datBlockEnd(ByVal vData As Date)
    mvarBlockEnd = vData
End Property
Public Property Get objDrugs(index As Integer) As Object
    Set objDrugs = mobjDrugs(index)
End Property
Public Property Let objDrugs(index As Integer, vData As Object)
    mobjDrugs(index) = vData
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property

