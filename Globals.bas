Attribute VB_Name = "Globals"
Option Explicit
'This module holds global variables and public functions used across the rest of the program

Public startDate As Date                    'The starting day of treatment
Public InitialWorkbook As String            'The name of the initial workbook
Public Drugs As clDrugs                     'The collection of chemotherapy drugs
Public AdminDays As clAdminDays             'The abstracted collection of days on which medication is given
Public Calendar As clCalendar               'The controller class for making a calendar
Public Information As clInformation         'The class holding demographic information
Public PreMeds As clPreMeds                 'The collection of premedictions
Public Labs As clLabs                       'The collection of labs
Public OrderSheets As clOrderSheets         'The collection of order sheets
Public bDaysEntry As Boolean                'Flag telling us how data entry for days and weeks is occuring (true = days, false = day/week)
Public arMonth(1 To 31, 1 To 4)
Public bAbort As Boolean
Public MasterDoseList(0 To 14) As String    'The list of all dosing options
Public lFillColour As Long                  'User defined fill colour for headers
Public iFillColour As Long                  'Fill colour for inpatient Days
Public oFillColour As Long                  'Fill colour for outpatient days
Public hFillColour As Long                  'Fill colour for home days

Public Function isArrayAllocated(arr As Variant) As Boolean
    'Returns whether or not an array has been initialized
    On Error Resume Next
    isArrayAllocated = IsArray(arr) And Not IsError(LBound(arr, 1)) And _
                        LBound(arr, 1) <= UBound(arr, 1)
End Function
Public Function ArrayLen(arr As Variant) As Integer
    'Returns the length of an array
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Public Function Contains(col As Collection, Key As Variant) As Boolean
    'Returns whether or not a collection contains a certain key
    Dim obj As Variant
    On Error GoTo err
        Contains = True
        IsObject (col(Key))
        Exit Function
err:
    Contains = False
End Function
Public Function WorksheetExists(shtname As String, Optional wb As Workbook) As Boolean
    'Returns whether or not a worksheet exists
    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Set sht = wb.Sheets(shtname)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
Public Function IndexOf(ByVal coll As Collection, ByVal Item As Variant) As Long
    'Returns the index of a member of a collection
    Dim i As Long
    For i = 1 To coll.Count
        If coll(i) = Item Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function
Public Function ArraySort(arr As Variant) As Variant()
    'Sorts an array, uses bubble sort
    Dim first As Integer, last As Integer
    Dim i As Long, j As Long
    Dim temp As Variant
    
    first = LBound(arr)
    last = UBound(arr)
    
    'I DO LOVE MY BUBBLE SORTS PERFORMANCE BE DAMNED
    For i = first To last - 1
        For j = i + 1 To last
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
    
    ArraySort = arr
End Function
Public Sub insertRowBelow(rng As Range)
    'Inserts a row below a given row with identical formatting to the row above
    rng.offset(1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    rng.EntireRow.Copy
    rng.offset(1).EntireRow.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub
Public Sub intInsertSort(ByRef arr() As Integer, ByVal vData As Integer)
    'Inserts a value into an array, keeping it sorted from least to most
    'Does not allow duplicate entries
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ErrorHandle
    
    For i = 0 To UBound(arr)
        If arr(i) = vData Then          'No duplicates allowed, if the value is already in the array, exit
            Exit Sub
        ElseIf arr(i) > vData Then     'Search until we find a value greater than our input value
            ReDim Preserve arr(UBound(arr) + 1)
            For j = UBound(arr) To (i + 1) Step -1        'Loop backwards from the end, moving the values
                arr(j) = arr(j - 1)
            Next j
            arr(i) = vData
            Exit Sub
        End If
    Next i
    ReDim Preserve arr(UBound(arr) + 1)   'if vData is now the biggest value, append it to the end
    arr(UBound(arr)) = vData

    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in intInsertSort, Globals"
End Sub
Public Sub datInsertSort(ByRef arr() As Date, ByVal vData As Date)
    'Inserts a value into an array, keeping it sorted from least to most
    'Does NOT allow duplicate entries
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ErrorHandle
    
    For i = 0 To UBound(arr)
        If arr(i) = vData Then          'No duplicates allowed, if the value is already in the array, exit
            Exit Sub
        ElseIf arr(i) > vData Then      'Search until we find a value greater than our input value
            ReDim Preserve arr(UBound(arr) + 1)
            For j = UBound(arr) To (i + 1) Step -1        'Loop backwards from the end, moving the values
                arr(j) = arr(j - 1)
            Next j
            arr(i) = vData
            Exit Sub
        End If
    Next i
    ReDim Preserve arr(UBound(arr) + 1)   'if vData is now the biggest value, append it to the end
    arr(UBound(arr)) = vData

    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in datInsertSort, Globals"
End Sub
Public Sub InsertSort(ByRef arr() As Variant, ByVal vData As Variant)
    'Inserts a value into an array, keeping it sorted from least to most
    'Does NOT allow duplicate entries
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ErrorHandle
    
    For i = 0 To UBound(arr)
        If arr(i) = vData Then          'No duplicates allowed, if the value is already in the array, exit
            Exit Sub
        ElseIf arr(i) > vData Then      'Search until we find a value greater than our input value
            ReDim Preserve arr(UBound(arr) + 1)
            For j = UBound(arr) To (i + 1) Step -1        'Loop backwards from the end, moving the values
                arr(j) = arr(j - 1)
            Next j
            arr(i) = vData
            Exit Sub
        End If
    Next i
    ReDim Preserve arr(UBound(arr) + 1)   'if vData is now the biggest value, append it to the end
    arr(UBound(arr)) = vData

    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in InsertSort, Globals"
End Sub
Public Function LinearIndex(ByRef arr As Variant, ByVal week As Integer, ByVal day As Integer)
    'returns the "linear index" of a given day and week in our 2d Week arrays
    'A linear index is the index if we converted our jagged array into a vector (row major order)
    
    Dim i As Integer, j As Integer, index As Integer
    
    On Error GoTo ErrorHandle
    For i = 1 To week
        If Not IsEmpty(arr(i)) Then
            For j = 0 To ArrayLen(arr(i)) - 1
                    If i = week And arr(i)(j) = day Then
                        GoTo Exitloop
                    End If
                  index = index + 1
            Next j
        End If
    Next i
    MsgBox "Error: day/week combination not found"
    Debug.Print "Tried to find day " & day & " and week " & week
    'We don't want this to error silently so
    LinearIndex = -1
    Exit Function
Exitloop:
    LinearIndex = index
    Exit Function
ErrorHandle:
    MsgBox err.Description & " in Linearindex, Globals"
End Function
Public Sub RemoveArrIndex(ByVal index As Integer, ByRef arr() As Variant)
    'Removes a specific index from an array; the indices of all entries after
    'the removed entry change by -1
    Dim i As Integer
    Dim lwr() As Variant
    Dim uppr() As Variant
    
    On Error GoTo ErrorHandle

    If (Not Not arr) = 0 Then                           'If the array is not initialized, exit
        Exit Sub
    ElseIf ArrayLen(arr) = 1 Then                       'If the array only has one entry, erase it
        Erase arr
        Exit Sub
    ElseIf index = UBound(arr) Then                     'If the index is the last entry in the array,
        ReDim Preserve arr(UBound(arr) - 1)             'simply redimension it to be one shorter
    ElseIf index = LBound(arr) Then                     'If the index is the first entry in the array,
        ReDim uppr(LBound(arr) + 1 To UBound(arr))      'Copy the rest of the array into a new array and then resize
        For i = LBound(arr) + 1 To UBound(arr)
            uppr(i) = arr(i)
        Next i
        ReDim arr(LBound(arr) To UBound(arr) - 1)
        For i = LBound(uppr) To UBound(uppr)
            arr(i - 1) = uppr(i)
        Next i
    Else                                                'Otherwise, slice the array arround the index we want to
        ReDim lwr(LBound(arr) To index - 1)             'remove into two temporary arrays, resize the array
        ReDim uppr(index + 1 To UBound(arr))            'and copy values back from the two temp arrays
        For i = LBound(arr) To index - 1
            lwr(i) = arr(i)
        Next i
        For i = index + 1 To UBound(arr)
            uppr(i) = arr(i)
        Next i
        ReDim arr(LBound(arr) To UBound(arr) - 1)
        For i = LBound(lwr) To UBound(lwr)
            arr(i) = lwr(i)
        Next i
        For i = LBound(uppr) To UBound(uppr)
            arr(i - 1) = uppr(i)
        Next i
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in RemoveArrIndex, Globals"
End Sub
Public Sub datRemoveArrIndex(ByVal index As Integer, ByRef arr() As Date)
    Dim i As Integer
    Dim lwr() As Date
    Dim uppr() As Date
    
    On Error GoTo ErrorHandle

    If (Not Not arr) = 0 Then
        Exit Sub
    ElseIf ArrayLen(arr) = 1 Then
        Erase arr
        Exit Sub
    ElseIf index = UBound(arr) Then
        ReDim Preserve arr(UBound(arr) - 1)
    ElseIf index = LBound(arr) Then
        ReDim uppr(LBound(arr) + 1 To UBound(arr))
        For i = LBound(arr) + 1 To UBound(arr)
            uppr(i) = arr(i)
        Next i
        ReDim arr(LBound(arr) To UBound(arr) - 1)
        For i = LBound(uppr) To UBound(uppr)
            arr(i - 1) = uppr(i)
        Next i
    Else
        ReDim lwr(LBound(arr) To index - 1)
        ReDim uppr(index + 1 To UBound(arr))
        For i = LBound(arr) To index - 1
            lwr(i) = arr(i)
        Next i
        For i = index + 1 To UBound(arr)
            uppr(i) = arr(i)
        Next i
        ReDim arr(LBound(arr) To UBound(arr) - 1)
        For i = LBound(lwr) To UBound(lwr)
            arr(i) = lwr(i)
        Next i
        For i = LBound(uppr) To UBound(uppr)
            arr(i - 1) = uppr(i)
        Next i
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description & " in RemoveArrIndex, Globals"
End Sub
