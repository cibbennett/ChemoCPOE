VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clLabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolLabs As Collection         'Collection of clLab Classes
'*************************
' Initialize/Terminate
'*************************
Private Sub Class_Initialize()
'Runs when the class is initialized/create

On Error GoTo ErrorHandle

'Create a new class collection of drugs
Set mcolLabs = New Collection

ThisWorkbook.Activate

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure Class_Initialize, clLabs"
End Sub
Private Sub Class_Terminate()
    Set mcolLabs = Nothing
End Sub
'*************************
' Add/Remove Utilities
'*************************
Public Sub Remove(vntIndexKey As Variant)
    'Can be used to remove a drug from the collection
    mcolLabs.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clLab
'Adds a clDrug class to our class collection
Dim objNewMember As clLab
On Error GoTo ErrorHandle

'Create a new object, of type clDrug
Set objNewMember = New clLab

objNewMember.Key = Key


If Len(sKey) = 0 Then
    mcolLabs.Add objNewMember
Else
    mcolLabs.Add objNewMember, sKey
End If
    
Set mcolLabs.Item(sKey).Parent = Me

'Destroy the standalone object as we have added it to the collection
Set objNewMember = Nothing
Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clDrugs"
End Function
'********************
' Class Properties
'********************
Public Property Get Item(vntIndexKey As Variant) As clLab
Set Item = mcolLabs(vntIndexKey)
End Property
Public Property Get Count() As Long
Count = mcolLabs.Count
End Property
'*******************
' Class Functions
'*******************
Public Function ContainsLab(vntIndexKey As Variant) As Boolean
    ContainsLab = Contains(mcolLabs, vntIndexKey)
End Function
