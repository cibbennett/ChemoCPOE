VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Key As String               'Key
Public sKey As String              'Secondary key
Private mvarFill As Boolean        'Background like weekends or not
Private mvarChemoDay As Boolean     'Chemotherapy day True/False
Private mvarWriteName As Boolean   'Write the holiday name or not
Private mvarWeekEnd As Boolean     'True if Saturday or Sunday
Private mvarDayInWeek As Long      'The days number in week. If week starts with Mon., Mondays are = 1
Private mvarDayName As String      'Day's name (Monday, Tuesday etc.)
Private mvarSpecialName As String  'A special title for the day, eg the drugs administered on that day or a holiday name
Private mvarDayType As Integer     'Flag for inpatient, outpatient, or home day (0 = inpatiet, 1 = outpatient, 2 = home)
Private mvarFullDate As Date       'The day's full date
Private mobjParent As clMonth      'Parent class
Private Sub Class_Initialize()
'Runs when the class is initialized/created.
'There is no code here, but it can be
'useful elsewhere.
End Sub

'***************************
'Properties
'***************************
Property Get bFill() As Boolean
bFill = mvarFill
End Property
Property Let bFill(ByVal vData As Boolean)
mvarFill = vData
End Property
Public Property Get bChemoDay() As Boolean
bChemoDay = mvarChemoDay
End Property
Public Property Let bChemoDay(ByVal vData As Boolean)
mvarChemoDay = vData
End Property
Property Get bWeekEnd() As Boolean
bWeekEnd = mvarWeekEnd
End Property
Property Let bWeekEnd(ByVal vData As Boolean)
mvarWeekEnd = vData
End Property
Property Get bWriteName() As Boolean
bWriteName = mvarWriteName
End Property
Property Let bWriteName(ByVal vData As Boolean)
mvarWriteName = vData
End Property
Property Get lDayInWeek() As Long
lDayInWeek = mvarDayInWeek
End Property
Property Let lDayInWeek(ByVal vData As Long)
mvarDayInWeek = vData
End Property
Property Get sDayName() As String
sDayName = mvarDayName
End Property
Property Let sDayName(ByVal vData As String)
mvarDayName = vData
End Property
Property Get sSpecialName() As String
sSpecialName = mvarSpecialName
End Property
Property Let sSpecialName(ByVal vData As String)
mvarSpecialName = vData
End Property
Property Get datFullDate() As Date
datFullDate = mvarFullDate
End Property
Property Let datFullDate(ByVal vData As Date)
mvarFullDate = vData
End Property
Property Get iDayType() As Integer
iDayType = mvarDayType
End Property
Property Let iDayType(ByVal vData As Integer)
mvarDayType = vData
End Property
Property Get Parent() As Object
    Set Parent = mobjParent
End Property
Property Set Parent(objParent As Object)
    If mobjParent Is Nothing Then
        Set mobjParent = objParent
    End If
End Property
