VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'A module that holds the patient and cycle information entered by the user
'You could store all of this in global variables (as we do the starting date
'because it is widely used in other functions), but this seems cleaner for most of the data

Private mvarCycle As Long
Private mvarCycleName As String
Private mvarProtocol As String
Private mvarCycleLength As Long
Private mvarMRN As String
Private mvarPName As String
Private mvarDOB As Date
Private mvarHeight As Double
Private mvarWeight As Double
Private mvarBSA As Double
Private mvarAllergies As String
Private mvarDiagnosis As String
Private mvarProvName As String
Private mvarProvPage As String
Private mvarProvNPI As String
Private mvarProvPPG As String
Private mvarPatCOGID As String


'******************
'Class Initialize
'******************
Private Sub Class_Initialize()
    'When we initialize the class we will pull data from the GUI to fill in our class properties
    With ThisWorkbook.Worksheets("Creator")
        If IsNumeric(.Range("cycle").Value) Then
            Me.lCycle = .Range("cycle").Value
        Else
            Me.sCycleName = .Range("cycle")
        End If
        Me.sProtocol = .Range("protocol").Value
        Me.lCycleLength = .Range("clength").Value
        Me.sMRN = .Range("patMRN").Value
        Me.sPName = .Range("patname").Value
        Me.datDOB = .Range("patDOB").Value
        Me.dubHeight = .Range("patHeight").Value
        Me.dubWeight = .Range("patWeight").Value
        Me.dubBSA = .Range("patBSA").Value
        Me.sAllergies = .Range("patAllergies").Value
        Me.sProvName = .OLEObjects("cmbProvider").Object.Column(0)
        Me.sProvPage = .OLEObjects("cmbProvider").Object.Column(1)
        Me.sProvNPI = .OLEObjects("cmbProvider").Object.Column(2)
        Me.sProvPPG = .OLEObjects("cmbProvider").Object.Column(3)
        Me.sDiagnosis = .Range("patDiagnosis").Value
        Me.sPatCOGID = .Range("patCOGID").Value
    End With
End Sub
'******************
'Class Properties
'******************
Public Property Get lCycle() As Long
    lCycle = mvarCycle
End Property
Public Property Let lCycle(ByVal vData As Long)
    mvarCycle = vData
End Property
Public Property Get sCycleName() As String
    sCycleName = mvarCycleName
End Property
Public Property Let sCycleName(ByVal vData As String)
    mvarCycleName = vData
End Property
Public Property Get sProtocol() As String
    sProtocol = mvarProtocol
End Property
Public Property Let sProtocol(ByVal vData As String)
    mvarProtocol = vData
End Property
Public Property Get lCycleLength() As Long
    lCycleLength = mvarCycleLength
End Property
Public Property Let lCycleLength(ByVal vData As Long)
    mvarCycleLength = vData
End Property
Public Property Get sMRN() As String
    sMRN = mvarMRN
End Property
Public Property Let sMRN(ByVal vData As String)
    mvarMRN = vData
End Property
Public Property Get sPName() As String
    sPName = mvarPName
End Property
Public Property Let sPName(ByVal vData As String)
    mvarPName = vData
End Property
Public Property Get sDiagnosis() As String
    sDiagnosis = mvarDiagnosis
End Property
Public Property Let sDiagnosis(ByVal vData As String)
    mvarDiagnosis = vData
End Property
Public Property Get datDOB() As Date
    datDOB = mvarDOB
End Property
Public Property Let datDOB(ByVal vData As Date)
    mvarDOB = vData
End Property
Public Property Get dubHeight() As Double
    dubHeight = mvarHeight
End Property
Public Property Let dubHeight(ByVal vData As Double)
    mvarHeight = vData
End Property
Public Property Get dubWeight() As Double
    dubWeight = mvarWeight
End Property
Public Property Let dubWeight(ByVal vData As Double)
    mvarWeight = vData
End Property
Public Property Get dubBSA() As Double
    dubBSA = mvarBSA
End Property
Public Property Let dubBSA(ByVal vData As Double)
    mvarBSA = vData
    With Application.WorksheetFunction
        mvarBSA = .Round(mvarBSA, 3 - (1 + Int(.Log10(Abs(mvarBSA)))))  'round to 3 sig figs
    End With
End Property
Public Property Get sAllergies() As String
    sAllergies = mvarAllergies
End Property
Public Property Let sAllergies(ByVal vData As String)
    mvarAllergies = vData
End Property
Public Property Get sProvName() As String
    sProvName = mvarProvName
End Property
Public Property Let sProvName(ByVal vData As String)
    mvarProvName = vData
End Property
Public Property Get sProvPage() As String
    sProvPage = mvarProvPage
End Property
Public Property Let sProvPage(ByVal vData As String)
    mvarProvPage = vData
End Property
Public Property Get sProvNPI() As String
    sProvNPI = mvarProvNPI
End Property
Public Property Let sProvNPI(ByVal vData As String)
    mvarProvNPI = vData
End Property
Public Property Get sProvPPG() As String
    sProvPPG = mvarProvPPG
End Property
Public Property Let sProvPPG(ByVal vData As String)
    mvarProvPPG = vData
End Property
Public Property Get sPatCOGID() As String
    sPatCOGID = mvarPatCOGID
End Property
Public Property Let sPatCOGID(ByVal vData As String)
    mvarPatCOGID = vData
End Property
