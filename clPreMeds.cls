VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clPreMeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class collection for PreMeds
'This collection is filled out via Add calls in the PreMedications
'userform at the time that it is initialized; this allows
'the checkboxes that correlate to the premeds to be made
'simultaneously

Private mcolPreMeds As Collection         'Collection of clPreMed Classes
'*************************
' Initialize/Terminate
'*************************
Private Sub Class_Initialize()
'Runs when the class is initialized/create

On Error GoTo ErrorHandle

'Create a new class collection of drugs
Set mcolPreMeds = New Collection

ThisWorkbook.Activate

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure Class_Initialize, clPreMeds"
End Sub
Private Sub Class_Terminate()
    Set mcolPreMeds = Nothing
End Sub
'*************************
' Add/Remove Utilities
'*************************
Public Sub Remove(vntIndexKey As Variant)
    'Can be used to remove a drug from the collection
    mcolPreMeds.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clPreMed
'Adds a clDrug class to our class collection
Dim objNewMember As clPreMed
On Error GoTo ErrorHandle

'Create a new object, of type clDrug
Set objNewMember = New clPreMed

objNewMember.Key = Key


If Len(sKey) = 0 Then
    mcolPreMeds.Add objNewMember
Else
    mcolPreMeds.Add objNewMember, sKey
End If
    
Set mcolPreMeds.Item(sKey).Parent = Me

'Destroy the standalone object as we have added it to the collection
Set objNewMember = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clDrugs"
End Function
'********************
' Class Properties
'********************
Public Property Get Item(vntIndexKey As Variant) As clPreMed
Set Item = mcolPreMeds(vntIndexKey)
End Property
Public Property Get Count() As Long
Count = mcolPreMeds.Count
End Property
'*******************
' Class Functions
'*******************
Public Function ContainsPreMed(vntIndexKey As Variant) As Boolean
    ContainsPreMed = Contains(mcolPreMeds, vntIndexKey)
End Function
