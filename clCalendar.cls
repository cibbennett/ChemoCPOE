VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The Calendar class is the top class keeping the pieces
'together. It is a class collection of Month classes,
'and each Month class is a class collection of Day classes.
'All months and all days have the same properties, and
'that is a huge advantage, when you are dealing with 12
'months and 360 something days.


'The clCalendar class has the following procedures and functions:

'Class_Initialize - not used except for reading the fill colour for weekends and holidays
'StartCal - the top, controlling procedure
'MakeCalendarCols - makes a calendar with columns for days and a sheet for each month
'MakeCalendarRows - makes a calendar with rows for days and a sheet for each month
'MakeCalendarSq - makes a calendar with "boxes" for days, 7 in x rows and a sheet for each month
'MakeCalendarPage - makes a calendar with 2 tables, eachcontaining 6 months
'InsertMonth - inserts a month table in the calendar with 2 tables, each containing 6 months
'MonthLayout - formats the layout of the 2 tables calendar
'TransferProperties - transfers user defined holiday properties to the class collection of days
'MakeCellBorders - makes cell borders in the calendars with columns or rows for days
'SquareBorders - makes borders around a passed range
'WithHours - start procedure for calendar with columns for days and 24 rows (inserts hours)
'MakeWorkBook - makes a new workbook and (for most calendars) a sheet for each month
'DayPosition - finds the day and month position in a date depending on the user's system settings
'Remove - not used, but can remove a month class from the class collection of months
'Class_Terminate - executes when the Calendar class is set to nothing
'Function Add - adds a clMonth to the Calendar's class collection
'Function DaysInMonth - returns the number of days in a given month (checks for leap year)
'Function ReturnDate - returns a date with day, month and year in the positions defined by the system settings
'There are also some properties. Included is an array. There are some restrictions to defining arrays as
'variables in a class module, but you can do it if the array is a property.

Option Explicit
Public KeyMonth As String           'Key
Public sKey As String               'Secondary key if any
Private mvarDayPos As Long          'Position of day number in a date
Private mvarMonthPos As Long        'Position of month number in a date
Private mvarWeekStart As Long       'Week start day number
Private mvarStartCell As Range      'Start cell for holidays settings
Private mcolMonths As Collection    'The collection of clMonth classes
Private Sub Class_Initialize()
'The code in the procedure below, StartCal,
'could be executed here, but when running
'it the first time in a session, I got an
'error when writing to some properties,
'because the class wasn't fully initialized.
'Therefore the class is first initialized, and
'THEN we execute the StartCal procedure.

'Sets the fill colour for weekends and holidays
oFillColour = Range("ofillcolour").Interior.Color
iFillColour = Range("ifillcolour").Interior.Color
hFillColour = Range("hfillcolour").Interior.Color
lFillColour = Range("fillcolour").Interior.Color
End Sub
Private Sub Class_Terminate()
    Set mcolMonths = Nothing
End Sub
Sub StartCal()
'Creates sub classes (all the initial work)
'and creates a calendar

Dim lCount As Long      'Counter
Dim lDayCount As Long   'Counter
Dim sMonth As String    'Month name
Dim lstartMonth As Long  'Starting Month

On Error GoTo ErrorHandle

'Turn off screen updating for speed
Application.ScreenUpdating = False

'Initialize the month class collection:
Set mcolMonths = New Collection

'Get starting month
lstartMonth = Month(startDate)

'Populate the collection with months
If lstartMonth = 1 Then
    For lCount = 1 To 12
        sMonth = MonthName(lCount)
        Me.Add sMonth, sMonth
        Me.Item(sMonth).lYear = Year(startDate)
        If bAbort Then GoTo BeforeExit
    Next
Else
    For lCount = lstartMonth To 12
        sMonth = MonthName(lCount)
        Me.Add sMonth, sMonth
        Me.Item(sMonth).lYear = Year(startDate)
        If bAbort Then GoTo BeforeExit
    Next
    For lCount = 1 To (lstartMonth - 1)
        sMonth = MonthName(lCount)
        Me.Add sMonth, sMonth
        Me.Item(sMonth).lYear = Year(startDate) + 1
        If bAbort Then GoTo BeforeExit
    Next
End If

'Find the user's date format, i.e. the
'placement of day and month in a date.
DayPosition

'Loop through the months and create their
'collections of days.
For lCount = 1 To 12
   'The number of days in the month is
   'stored in the month's lDays-property.
   With Me.Item(MonthName(lCount))
      'Number of days in month
      .lDaysInMonth = DaysInMonth(lCount)
      
      'The months' days collections were set by
      'the clMonth's Initialize procedure, and
      'now we populate them with days and set
      'some properties.
      For lDayCount = 1 To .lDaysInMonth
         'Add the day to the days collection
         .Add Str(lDayCount), Str(lDayCount)
         
         'The day's full date
         .Item(lDayCount).datFullDate = ReturnDate(lDayCount, lCount, .lYear)
         
         'The day's name. By using the function StrConv and vbPropercase
         'the first letter will be in upper case.
         .Item(lDayCount).sDayName = StrConv(WeekdayName _
         (Weekday(.Item(lDayCount).datFullDate, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek), vbProperCase)
         
         'Saturday or Sunday?
         If Weekday(.Item(lDayCount).datFullDate, vbMonday) = 6 Or _
            Weekday(.Item(lDayCount).datFullDate, vbMonday) = 7 Then
            .Item(lDayCount).bWeekEnd = True
         End If
         
         'Day position in week using system settings
         'for first day in week.
         .Item(lDayCount).lDayInWeek = Weekday(.Item(lDayCount).datFullDate, vbUseSystemDayOfWeek)
      Next
   End With
Next


TransferProperties


Worksheets.Item(1).Activate


'Make the new workbook that
'will contain the calendar
MakeWorkBook

If bAbort Then GoTo BeforeExit

'Now for the calendar
MakeCalendarSq

   
BeforeExit:
   
Application.ScreenUpdating = True
   
Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure StartCal, clCalendar"
Resume BeforeExit
End Sub

Sub MakeCalendarSq()
'Makes a calendar with "boxes" for days
'arranged with 7 horizontal boxes
'representing a week and as many as
'required. Depending on which day is the
'first in the month, the leading and
'trailing boxes can be days in the
'previous and next month.
Dim bBg As Boolean
Dim lWeeks As Long      'Number of weeks row
Dim lB As Long          'Variable
Dim lC As Long          'Variable
Dim lD As Long          'Day in month counter
Dim lM As Long          'Month counter 1-12
Dim lCount As Long      'Counter
Dim lFirstPos As Long   'Position in week of Jan. 1st
Dim lLastPos As Long    'Position in week of month's last day
Dim lCurPos As Long     'Current position in month table
Dim rMonth As Range     'Range for month table
Dim rSq As Range        'Range variable for 2 cells

On Error GoTo ErrorHandle

'Loop the months (outer loop)
For lM = 1 To 12
   'Each week row starts with the first day in the week.
   'In most countries that is Monday, but in e.g. the
   'USA it is Sunday, so we use the system settings.
   'That means that if the first day in a month isn't
   'the first day in the week, we find the position
   'and there will be some days from the previous
   'month in front.
   'If for instance the first day in the month is
   'a Wednesday, and the week starts with a Monday,
   'the first week row in February will be:
   'MO-30, TU-31, WE-1, TH-2, FR-3, SA-4, SU-5.
   'The same goes for the last week row with days
   'from next month, if the last day of the month
   'isn't the last day of the week.
   
   'Activate the month's sheet
   Worksheets.Item(lM).Activate
   
   'Position of first and last day in the month.
   'This depends on which day starts a week.
   With Me.Item(lM)
      lFirstPos = Weekday(.Item(1).datFullDate, vbUseSystemDayOfWeek)
      lLastPos = Weekday(.Item(.Count).datFullDate, vbUseSystemDayOfWeek)
   End With
   
   'We now calculate how many week rows we need in the table
   lWeeks = Me.Item(lM).lDaysInMonth - (8 - lFirstPos)
   lB = Int(lWeeks / 7)
   lC = lB Mod 7
   If lC > 0 Then lC = 1
   If lLastPos < 7 Then
      lWeeks = (lB + lC + 1)
   Else
      lWeeks = (lB + lC)
   End If
   
   Set rMonth = Range("B2")
   'We need two rows for each day/week,
   'hence the lWeeks * 2 in the next line.
   Set rMonth = rMonth.Resize(lWeeks * 2, 7)
   
   lCurPos = lFirstPos - 1
   
   'Loop the days (inner loop)
   For lD = 1 To Me.Item(lM).Count
      'If we have finished the last date in a
      'row, we skip one row.
      If lCurPos > 0 And lCurPos Mod 7 = 0 Then
         lCurPos = lCurPos + 8
      Else
         lCurPos = lCurPos + 1
      End If
      
      'Days before 1st day in month
      If lD = 1 And lFirstPos > 1 Then
         'If January
         If lM = 1 Then
            lB = 32
         Else 'Not January
            lB = Me.Item(lM - 1).Count + 1
         End If
         For lCount = lFirstPos - 1 To 1 Step -1
            lB = lB - 1
            With rMonth
               'Write the date (only day part)
               .Item(lCount).Value = lB
               'Dates not in this month are faded - use the
               'ColorIndex you prefer.
               Range(.Item(lCount), .Item(lCount).offset(1, 0)).Font.ColorIndex = 48
            End With
         Next
      End If
      
      'Days after last day in the last week row of the month
      If lD = Me.Item(lM).Count And lLastPos < 7 Then
         lB = 0
         For lCount = rMonth.Count - 7 - (6 - lLastPos) To rMonth.Count - 7
            lB = lB + 1
            With rMonth
               .Item(lCount).Value = lB
               'Dates not in this month are faded - use the
               'ColorIndex you prefer.
               Range(.Item(lCount), .Item(lCount).offset(1, 0)).Font.ColorIndex = 48
            End With
         Next
      End If
      
      'Dates in this month: Check special day settings and do
      'what they say. Word wrap text.
      With Me.Item(lM).Item(lD)
         rMonth.Item(lCurPos).Value = lD
         'Word wrap and align text to top
         With rMonth.Item(lCurPos).offset(1, 0)
            .WrapText = True
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .Font.Size = 10
         End With
         If .bChemoDay Then
               rMonth.Item(lCurPos).offset(1, 0).Value = .sSpecialName
               bBg = True
         End If
      End With
      
      'Color inpatient vs outpatient days
      If bBg Then
         With Range(rMonth.Item(lCurPos), rMonth.Item(lCurPos).offset(1, 0)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            If Me.Item(lM).Item(lD).iDayType = 0 Then
                .Color = iFillColour
            ElseIf Me.Item(lM).Item(lD).iDayType = 1 Then
                .Color = oFillColour
            Else
                .Color = hFillColour
            End If
         End With
         bBg = False
      End If
      
      'Week numbers in column A - use system settings.
      If lCurPos Mod 7 = 0 Or lD = Me.Item(lM).Count Then
         lC = DatePart("ww", Me.Item(lM).Item(lD).datFullDate, vbUseSystemDayOfWeek, vbUseSystem)
         With rMonth
            'If the current position is in column 3 or higher, we use
            'the .End(xlToLeft) and offset to locate the cell.
            If .Item(lCurPos).Column > 2 Then
               rMonth.Item(lCurPos).End(xlToLeft).offset(0, -1).Value = "Week " & " " & lC
            Else
               'If column 2 (B) it is just 1 cell to the left
               .Item(lCurPos).offset(0, -1) = "Week " & " " & lC
            End If
         End With
      End If
   Next lD

   'Insert short day names in row A using the user's language settings.
   'By using the StrConv function we can get the first letter in
   'upper case ("vbProperCase").
   For lCount = 1 To 7
      With rMonth.Item(lCount).offset(-1, 0)
         .Value = StrConv(WeekdayName(lCount, True, vbUseSystemDayOfWeek), vbProperCase)
         .HorizontalAlignment = xlCenter
         .Interior.Pattern = xlSolid
         .Interior.PatternColorIndex = xlAutomatic
         .Interior.Color = lFillColour   'Use what fill colour you want
         .Font.FontStyle = "bold"
      End With
   Next
   
   'Resize the tables's cells
   With rMonth.Rows
      For lCount = 1 To .Rows.Count Step 2
         .Item(lCount).HorizontalAlignment = xlRight
      Next
      For lCount = 2 To .Rows.Count Step 2
         .Item(lCount).RowHeight = 70
      Next
      rMonth.Columns.ColumnWidth = 20
      Columns("A:A").ColumnWidth = 3
   End With

   'Turn off gridlines
   ActiveWindow.DisplayGridlines = False
   
   'Format the week numbers
   For lCount = 1 To rMonth.Count - 13 Step 14
      Set rSq = rMonth.Item(lCount).offset(0, -1)
      Set rSq = rSq.Resize(2)
      With rSq
          .VerticalAlignment = xlCenter
          .Orientation = 90
          .MergeCells = True
          .Font.Size = 8
      End With
      'Call the procedure that makes borders
      SquareBorders rSq
   Next lCount
   
   'Now borders for the dates & squares
   For lCount = 1 To rMonth.Count - 7
      Set rSq = rMonth.Item(lCount)
      Set rSq = rSq.Resize(2)
      'Procedure call for borders
      SquareBorders rSq
      If lCount Mod 7 = 0 Then
         If lCount + 1 < rMonth.Count - 6 Then
            lCount = lCount + 7
         End If
      End If
   Next lCount
   
   'Change to Landscape orientation & scale the columns
   ActiveSheet.PageSetup.Orientation = xlLandscape
   Application.PrintCommunication = False
   ActiveSheet.PageSetup.FitToPagesWide = 1
   ActiveSheet.PageSetup.FitToPagesTall = False
   Application.PrintCommunication = True
   
   'Write the name of the month at the bottom of the calendar
   ActiveSheet.Rows("15:16").RowHeight = 22
   ActiveSheet.Range("B15:B16").Merge
   With ActiveSheet.Range("B15")
    .Value = Me.Item(lM).Key
    .Font.Size = 18
    .Font.Bold = True
    .Font.Name = "Gill Sans MT"
    .VerticalAlignment = xlVAlignCenter
    .HorizontalAlignment = xlHAlignLeft
   End With
   
    With ActiveSheet.Range("B17")
    .Value = Me.Item(lM).lYear
    .Font.Size = 14
    .Font.Bold = True
    .Font.Name = "Gill Sans MT"
    .VerticalAlignment = xlVAlignCenter
    .HorizontalAlignment = xlHAlignLeft
   End With
   
   'Create the color legends
   With ActiveSheet.Range("H15")
   .Value = "Inpatient"
   .Font.Size = 12
   .Font.Bold = True
   .Font.Name = "Gill Sans MT"
    .VerticalAlignment = xlVAlignCenter
    .HorizontalAlignment = xlHAlignRight
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = iFillColour
   End With
   
    With ActiveSheet.Range("H16")
   .Value = "Outpatient"
   .Font.Size = 12
   .Font.Bold = True
   .Font.Name = "Gill Sans MT"
    .VerticalAlignment = xlVAlignCenter
    .HorizontalAlignment = xlHAlignRight
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = oFillColour
   End With
   
    With ActiveSheet.Range("H17")
   .Value = "Home"
   .Font.Size = 12
   .Font.Bold = True
   .Font.Name = "Gill Sans MT"
    .VerticalAlignment = xlVAlignCenter
    .HorizontalAlignment = xlHAlignRight
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = hFillColour
   End With
   
   'Copy Pictures into each page
   Workbooks(InitialWorkbook).Sheets("Logos").Pictures("Picture 1").Copy
   ActiveSheet.Range("C15").PasteSpecial
   Workbooks(InitialWorkbook).Sheets("Logos").Pictures("Picture 2").Copy
   ActiveSheet.Range("E15").PasteSpecial
   
Next lM

'If we are through, activate January sheet
If lM = 13 Then Worksheets.Item(1).Activate

BeforeExit:
Set rMonth = Nothing
Set rSq = Nothing

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure MakeCalendarSq, clCalendar"
bAbort = True
Resume BeforeExit
End Sub
Private Sub TransferProperties()
'Transfers AdminDay properties from the
'AdminDays collection to the matching
'dates in the Calendar's class collection.
Dim i As Long   'Counter

On Error GoTo ErrorHandle

'Iterate through all admin days, transfering properties as we go
For i = 1 To AdminDays.Count
   With AdminDays.Item(i)
         'Set the flag for chemo day to true
         Me.Item(MonthName(.lMonth)).Item(.lDay).bChemoDay = True
         'Set the flag to fill the background of the cell to true
         Me.Item(MonthName(.lMonth)).Item(.lDay).bFill = True
         'Set the flag to write the SpecialName (the chemo given that day) of the day
         Me.Item(MonthName(.lMonth)).Item(.lDay).bWriteName = True
         'User defined holiday name
         Me.Item(MonthName(.lMonth)).Item(.lDay).sSpecialName = .sDayName
         'Inpatient, outpatient, or home?
         Me.Item(MonthName(.lMonth)).Item(.lDay).iDayType = .iDayType
   End With
Next

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure Transfer Properties, clCalendar"
bAbort = True
End Sub
Sub SquareBorders(ByRef rRange As Range)
'Makes borders around the passed range

On Error GoTo ErrorHandle

With rRange
  .Borders(xlDiagonalDown).LineStyle = xlNone
  .Borders(xlDiagonalUp).LineStyle = xlNone
End With
With rRange.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With rRange.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With rRange.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
With rRange.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure SquareBorders, clCalendar"
bAbort = True
End Sub
Sub MakeWorkBook()
Dim lSheet As Long
Dim startMonth As Long

On Error GoTo ErrorHandle

startMonth = Month(startDate)

'Add a new workbook to the workbooks collection
Workbooks.Add

For lSheet = 1 To 12 - Worksheets.Count
    Worksheets.Add
Next

For lSheet = 1 To 12
    'Name the sheet
    If (lSheet + startMonth - 1) <= 12 Then
        Worksheets.Item(lSheet).Name = UCase(Left(MonthName(lSheet + startMonth - 1), 3))
    Else
        Worksheets.Item(lSheet).Name = UCase(Left(MonthName(lSheet + startMonth - 13), 3))
    End If
Next

Exit Sub
ErrorHandle:
MsgBox err.Description & " Procedure MakeWorkBook, clCalendar"
bAbort = True
End Sub
Private Sub DayPosition()
'Find the day and month position in a date
'depending on the user's system settings
'There are other ways to do it, but this
'is a very simple approach.

Me.lDayPos = day("01-02-03")
Me.lMonthPos = Month("01-02-03")

End Sub
Public Sub Remove(vntIndexKey As Variant)
'Removes a clMonth class from the collection.
'Not used in this program, but can be useful if
'the code is used elsewhere.
mcolMonths.Remove vntIndexKey
End Sub
Public Function Add(Key As String, Optional sKey As String) As clMonth
'This function adds a clMonth class to the clCalendar's class collection.
Dim objNewMember As clMonth

On Error GoTo ErrorHandle
    
'Create a new object of the clMonth class
Set objNewMember = New clMonth

'Assigns a key to the new object
objNewMember.Key = Key

'Add the class to the collection.
If Len(sKey) = 0 Then
    mcolMonths.Add objNewMember
Else
    mcolMonths.Add objNewMember, sKey
End If

'The class now "lives" as a member of the collection,
'so we can destroy the old "stand alone" object.
Set objNewMember = Nothing

Exit Function
ErrorHandle:
MsgBox err.Description & ", Function Add, clCalendar"
bAbort = True
End Function
Private Function DaysInMonth(ByVal lMonth As Long) As Long
'Returns the number of days in a given month

On Error GoTo ErrorHandle

Select Case lMonth
   Case 1, 3, 5, 7, 8, 10, 12
      DaysInMonth = 31
   Case 2
      'You could check for leap year
      'using Year mod 4 = 0, but there
      'are a few exceptions, so this is
      'the safe way.
      If IsDate(ReturnDate(29, 2, Me.Item(2).lYear)) Then
         DaysInMonth = 29
      Else
         DaysInMonth = 28
      End If
   Case Else
      DaysInMonth = 30
End Select

Exit Function
ErrorHandle:
MsgBox err.Description & " Function DaysInMonth, clCalendar"
bAbort = True
End Function
Public Function ReturnDate(ByVal day As Long, ByVal Month As Long, ByVal Year As Long)
'Returns a date in the format (sequence of day, month and year)
'defined by the user's system settings.

Select Case lDayPos
   Case 1
      If Me.lMonthPos = 2 Then
         ReturnDate = day & "/" & Month & "/" & Year
      Else
         ReturnDate = day & "/" & Year & "/" & Month
      End If
   Case 2
      If Me.lMonthPos = 1 Then
         ReturnDate = Month & "/" & day & "/" & Year
      Else
         ReturnDate = Year & "/" & day & "/" & Month
      End If
   Case 3
      If Me.lMonthPos = 2 Then
         ReturnDate = Year & "/" & Month & "/" & day
      Else
         ReturnDate = Month & "/" & Year & "/" & day
      End If
End Select

End Function
'***************
'Properties
'***************
Public Property Get Item(vntIndexKey As Variant) As clMonth
'Finds/gets a certain month in the collection
Set Item = mcolMonths(vntIndexKey)
End Property
Public Property Get Count() As Long
Count = mcolMonths.Count
End Property
Public Property Get lDayPos() As Long
lDayPos = mvarDayPos
End Property
Property Let lDayPos(ByVal vData As Long)
mvarDayPos = vData
End Property
Public Property Get lMonthPos() As Long
lMonthPos = mvarMonthPos
End Property
Property Let lMonthPos(ByVal vData As Long)
mvarMonthPos = vData
End Property
Property Get rStartcell() As Range
    Set rStartcell = mvarStartCell
End Property
Public Property Set rStartcell(ByVal vData As Range)
   Set mvarStartCell = vData
End Property

