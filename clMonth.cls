VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The month class. There are 12 of these classes
'in the calendar´s class collection.
'Each month class is a class collection of
'clDay classes.

Option Explicit

Public Key As String             'Key
Public sKey As String            'Secondary key
Private mvarMonth As Long        'Month number
Private mvarDaysInMonth As Long  'Number of days in month
Private mvarYear As Long         'Year that the month falls in
Private mcolDays As Collection   'Collection of days (clDay classes) in month
Private Sub Class_Initialize()
Dim lDays As Long
'Runs when the class is initialized/created.
'Initialize the collection of day classes:
Set mcolDays = New Collection
End Sub
Private Sub Class_Terminate()
    Set mcolDays = Nothing
End Sub
Public Sub Remove(vntIndexKey As Variant)
'Removes a clDay class from the collection.
'Not used in this program, but can be useful if
'the code is used elsewhere.
mcolDays.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clDay
'This function adds a clDay class to our class collection.
Dim objNewMember As clDay

On Error GoTo ErrorHandle
    
'Create a new object of the clDay class
Set objNewMember = New clDay

'Assigns a key to the new object
objNewMember.Key = Key

'Sets the day's parent property = this month class
Set objNewMember.Parent = Me

'Add the class to the collection.
If Len(sKey) = 0 Then
    mcolDays.Add objNewMember
Else
    mcolDays.Add objNewMember, sKey
End If

'The class now "lives" as a member of the collection,
'so we can destroy the old "stand alone" object.
Set objNewMember = Nothing

Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clMonth"
bAbort = True
End Function

'***************************
'Properties
'***************************

Public Property Get Item(vntIndexKey As Variant) As clDay
'Finds/gets a certain class in the collection
Set Item = mcolDays(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mcolDays.Count
End Property
Property Get lDaysInMonth() As Long
lDaysInMonth = mvarDaysInMonth
End Property
Property Let lDaysInMonth(ByVal vData As Long)
mvarDaysInMonth = vData
End Property
Property Get lMonth() As Long
lMonth = mvarMonth
End Property
Property Let lMonth(ByVal vData As Long)
mvarMonth = vData
End Property
Property Get lYear() As Long
lYear = mvarYear
End Property
Property Let lYear(ByVal vData As Long)
mvarYear = vData
End Property
