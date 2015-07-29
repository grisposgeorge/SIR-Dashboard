'Constant variable declaration
Dim RowCount As Integer

Type Events
EventName As String
EventCount As Integer
End Type
Type Incidents
IncidentName As String
IncidentCount As Integer
End Type

Dim SheetNumber As Integer
Dim LastModifiedDate As Date

Dim CurrentDate As Date

'Sets menubar on the top of Excel sheet to provide interface for user
Sub Auto_Open()
 
 For Each mb In MenuBars
   mb.Reset
   mb.Menus.Add "Dashboard"
      
   mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Load Data", _
        OnAction:="Load_Data"
   mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
   mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Build Dashboard", _
        OnAction:="BuildDashboard"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Take SnapShot of Database", _
        OnAction:="SnapShotDb"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Close Incomplete Records", _
        OnAction:="CloseIncompleteRecords"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Close Open Records", _
        OnAction:="CloseOpenRecords"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="Close Data Extract", _
        OnAction:="CloseDataExtract"
  mb.Menus("Dashboard").MenuItems.Add _
        Caption:="-"
  Next
  
Call AddandDeleteSheets
MsgBox "The dashboard contains old data, select OK to reload the database and take a new snapshot", vbOKOnly + vbCritical, "Old Data Present"
Call Load_Data
    
End Sub

Sub AddandDeleteSheets()

'Delete all other sheets in spreadsheet apart from "Raw Data"
Dim ws As Worksheet
Application.DisplayAlerts = False
For Each ws In Worksheets
If ws.Name <> "Raw Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Sheets("Raw Data").Cells(1, 1).Value = "Subject"
Sheets("Raw Data").Cells(1, 2).Value = "Category"
Sheets("Raw Data").Cells(1, 3).Value = "Date Reported"
Sheets("Raw Data").Cells(1, 4).Value = "Time Reported"
Sheets("Raw Data").Cells(1, 5).Value = "Date Discovered"
Sheets("Raw Data").Cells(1, 6).Value = "Status"
Sheets("Raw Data").Cells(1, 7).Value = "Date Opened"
Sheets("Raw Data").Cells(1, 8).Value = "Time Opened"
Sheets("Raw Data").Cells(1, 9).Value = "Date Closed"
Sheets("Raw Data").Cells(1, 10).Value = "Time Closed"
Sheets("Raw Data").Cells(1, 11).Value = "Working Hours"
Sheets("Raw Data").Cells(1, 12).Value = "Incident Handler Name"
Sheets("Raw Data").Cells(1, 13).Value = "Summary"
Sheets("Raw Data").UsedRange.Columns.AutoFit
Sheets("Raw Data").Range("A1:M1").HorizontalAlignment = xlCenter
Sheets("Raw Data").Range("A1:M1").Font.Bold = True

CurrentDate = Date
CurrentMonth = Format(CurrentDate, "mmm")
CurrentYear = Format(CurrentDate, "yyyy")

End Sub
 
Sub Load_Data() 'Load data has been specifically written to access a Lotus Notes Server and fetch data.

Dim Stringbody As String

Set notesSess = CreateObject("Notes.NotesSession")
Set notesDB = notesSess.GETDATABASE("SERVER_NAME", "LOCATION OF DATABASE") 'Replace SERVER_NAME and LOCATION OF DATABASE with actual information

Set notesVW = notesDB.GetView("All Documents")
Set notesDoc = notesVW.GETFIRSTDOCUMENT '

Set notesDoc = notesVW.GETFIRSTDOCUMENT '

done1 = "N"
MaxDocCount = 0

If notesDoc Is Nothing Then
    done1 = "Y"
    End If

Do While done1 = "N"
    MaxDocCount = MaxDocCount + 1
    Set notesDoc = notesVW.GETNEXTDOCUMENT(notesDoc)
    If notesDoc Is Nothing Then
      done1 = "Y"
    End If
Loop

    Application.ScreenUpdating = False
    done = "NO"
    CurrentDocCount = 0
    Set notesDoc = notesVW.GETFIRSTDOCUMENT '
    If notesDoc Is Nothing Then
    done = "YES"
    End If

    RowCount = 2

    Do While done = "NO"

    CurrentDocCount = CurrentDocCount + 1

    Application.StatusBar = "Fetching record " & Trim(Str(CurrentDocCount)) & "/" & Trim(Str(MaxDocCount))

    Subject = notesDoc.getitemvalue("Subject")
    Category = notesDoc.getitemvalue("Categories")
    Body = notesDoc.getitemvalue("Body")

    Stringbody = Body(0)
    HandlerName = notesDoc.Authors
    LastModifiedDate = notesDoc.LastModified

    NewCheck = 0
    NewCheck = InStr(1, Stringbody, "Reporting and Contact Information", vbTextCompare)

    If NewCheck = 22 Then
    NewFormCheck = True
    Else
    NewFormCheck = False
    End If
         
    'If incident record is not actual record jump to the next record
    If Stringbody = "" Or Subject(0) = "" Or Subject(0) = 0 Or Category(0) = "Forms & Templates" Or _
    Subject(0) = "Lessons Learned" Or Subject(0) = "Security Incident Reporting Help" Or Subject(0) _
    = "Incident Reporting Checklist" Or Subject(0) = "Blank Incident Report" Then
    
    GoTo JumpLoop

    'If the Subject is correct carry on as normal
    Else

    'Display incident title
    RangeString = "A" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).Value = Subject(0)

    'Display incident handler's name
    RangeString = "L" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).Value = HandlerName

    'Call function to determine and display event/incident type
     Call Categories(Stringbody, Category(0), NewFormCheck)
     Call DateandTimeReported(Stringbody, NewFormCheck)
     Call NewDateDiscovered(Stringbody, NewFormCheck)
     Call OldDateDiscovered(Stringbody, NewFormCheck)
     Call RecordStatus(Subject)
     Call DateandTimeOpened(Stringbody, NewFormCheck)
     Call DateTimeClosed(Stringbody, LastModifiedDate, NewFormCheck, Subject(0))
     Call WorkingHours(Stringbody, NewFormCheck, Subject(0))
     Call SummaryCollect(Stringbody)
         
    End If
    RowCount = RowCount + 1

JumpLoop:
    Set notesDoc = notesVW.GETNEXTDOCUMENT(notesDoc)
    If notesDoc Is Nothing Then
      done = "YES"
    End If
Loop
    Worksheets("Raw Data").Select
    Sheets("Raw Data").UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = False

Call SnapShotDb

End Sub

'Determine & assign event and incident classification
Sub Categories(IncidentBody, NewCategory, NewRecordCheck)

    If NewRecordCheck = True Then

    Category_Clean = Application.WorksheetFunction.Clean(NewCategory)
    Checked_Category = Assign_Category(Category_Clean)
    
    RangeString = "B" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = Checked_Category

    Else
    Old_Category_1 = InStr(1, IncidentBody, "Incident type:", vbTextCompare)
    Old_Category_2 = InStr(1, IncidentBody, "Incident location:", vbTextCompare)
    Old_Category_3 = Old_Category_2 - Old_Category_1

    Defined_Category = Mid(IncidentBody, Old_Category_1 + 14, Old_Category_3 - 15)
    Category_Clean = Application.WorksheetFunction.Clean(Defined_Category)
    Checked_Category = Assign_Category(Category_Clean)

    RangeString = "B" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = Checked_Category
    End If

 End Sub

'Determines and assigns the date and time the event/incident was reported, works with the ChangeReportedDate and
'ChangeReportedTime functions described below
Sub DateandTimeReported(DateReportedBody, NewRecordCheck)

    'Calculate date and time reported for the new forms
    If NewRecordCheck = True Then

    NewDateReported1 = InStr(1, DateReportedBody, "Date Reported:", vbTextCompare)
    NewDateReported2 = InStr(1, DateReportedBody, "Time Reported:", vbTextCompare)

    NewTimeReported1 = InStr(1, DateReportedBody, "Time Reported:", vbTextCompare)
    NewTimeReported2 = InStr(1, DateReportedBody, "Reported To:", vbTextCompare)

    NewDateReportedDiff = NewDateReported2 - NewDateReported1
    NewTimeReportedDiff = NewTimeReported2 - NewTimeReported1

    TheDate2 = Mid(DateReportedBody, NewDateReported1 + 16, NewDateReportedDiff - 17)
    TheTime2 = Mid(DateReportedBody, NewTimeReported1 + 16, NewTimeReportedDiff - 17)

    TheDate2 = Application.WorksheetFunction.Clean(TheDate2)
    TheTime2 = Application.WorksheetFunction.Clean(TheTime2)

    TheDate3 = ChangeReportedDate(TheDate2)
    TheTime3 = ChangeReportedTime(TheTime2)

    Else

    'Calculate the date and time for the old forms
    OldDateCheck = InStr(1, DateReportedBody, "Date:", vbTextCompare)

    If OldDateCheck <> 339 Then

    TheDate3 = "Unknown"
    TheTime3 = "Unknown"

    Else

    OldDateReported1 = InStr(1, DateReportedBody, "Date:", vbTextCompare)
    oldDateReported2 = InStr(1, DateReportedBody, "Time:", vbTextCompare)

    OldTimeReported1 = InStr(1, DateReportedBody, "Time:", vbTextCompare)
    oldTimeReported2 = InStr(1, DateReportedBody, "Duration:", vbTextCompare)

    OldDateReportedDiff = oldDateReported2 - OldDateReported1
    OldTimeReportedDiff = oldTimeReported2 - OldTimeReported1

    TheDate = Mid(DateReportedBody, OldDateReported1 + 7, OldDateReportedDiff - 8)
    TheTime = Mid(DateReportedBody, OldTimeReported1 + 7, OldTimeReportedDiff - 8)

    TheDate2 = Application.WorksheetFunction.Clean(TheDate)
    TheTime2 = Application.WorksheetFunction.Clean(TheTime)

    TheDate3 = ChangeReportedDate(TheDate2)
    TheTime3 = ChangeReportedTime(TheTime2)

    End If
    End If

    RangeString = "C" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = TheDate3
    Sheets("Raw Data").Columns("C").HorizontalAlignment = xlCenter

    RangeString = "D" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = TheTime3
    Sheets("Raw Data").Columns("D").HorizontalAlignment = xlCenter

 End Sub

'Determines and assigns the date the event/incident was discovered - only works with the new records because this
'information was only being captured from the new forms going forward
Sub NewDateDiscovered(DiscoverReportedBody, Check)

    If Check = True Then

    NewDateDiscover1 = InStr(1, DiscoverReportedBody, "Date of Discovery / Detection", vbTextCompare)
    NewDateDiscover2 = InStr(1, DiscoverReportedBody, "Contact Name:", vbTextCompare)
    NewDateDiscoverDiff = NewDateDiscover2 - NewDateDiscover1
    DiscoverDate = Mid(DiscoverReportedBody, NewDateDiscover1 + 29, NewDateDiscoverDiff - 35)

    Else
    DiscoverDate = "Unknown"

    End If
    DiscoverDate2 = Application.WorksheetFunction.Clean(DiscoverDate)
    DiscoverDate3 = ChangeNewDiscoveredData(DiscoverDate2)

    RangeString = "E" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = DiscoverDate3
    Sheets("Raw Data").Columns("E").HorizontalAlignment = xlCenter

End Sub

Sub OldDateDiscovered(DiscoverReportedBody, Check)

   OldFormCheck = InStr(1, DiscoverReportedBody, "This is a Strictly Confidential Incident Report", vbTextCompare)

    If OldFormCheck >= 1 Then

    DiscoverDate = "Not recorded"
    DiscoverDate2 = Application.WorksheetFunction.Clean(DiscoverDate)
    DiscoverDate3 = ChangeOldDiscoveredData(DiscoverDate2)

    Else: GoTo EndofSub
    End If

    RangeString = "E" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = DiscoverDate3
    Sheets("Raw Data").Columns("E").HorizontalAlignment = xlCenter

EndofSub:

End Sub

'Determines and assigns the status of the record based on the "Closed" value assigned to the records
Sub RecordStatus(StatusSubject)

    IncidentClosedCheck = InStr(1, StatusSubject(0), "<CLOSED>")

    If IncidentClosedCheck <> 0 Then
    Status = "Closed"

    Else
    Status = "Open"

    End If
    RangeString = "F" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).Value = Status
End Sub

'Determines and assigns the date and time the record was opened, uses ChangeOpenDate and ChangeOpenTime to ensure
'that the date and time entered is consistent
Sub DateandTimeOpened(OpenBody, NewCheck)

    'Determines the date and time for the new form
    If NewCheck = True Then

    NewDate1 = InStr(1, OpenBody, "Date Opened:", vbTextCompare)
    NewDate2 = InStr(1, OpenBody, "Time Opened:", vbTextCompare)
    NewTime1 = InStr(1, OpenBody, "Time Opened:", vbTextCompare)
    NewTime2 = InStr(1, OpenBody, "Location:", vbTextCompare)

    NewOpenDateDiff = NewDate2 - NewDate1
    NewOpenTimeDiff = NewTime2 - NewTime1

    NewDateOpen = Mid(OpenBody, NewDate1 + 14, NewOpenDateDiff - 15)
    NewTimeOpen = Mid(OpenBody, NewTime1 + 14, NewOpenTimeDiff - 15)

    NewDateOpen2 = Application.WorksheetFunction.Clean(NewDateOpen)
    NewTimeOpen2 = Application.WorksheetFunction.Clean(NewTimeOpen)

    NewDateOpen3 = ChangeOpenDate(NewDateOpen2)
    NewTimeOpen3 = NewChangeOpenTime(NewTimeOpen2)

    RangeString = "G" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = NewDateOpen3

    RangeString = "H" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = NewTimeOpen3

    'Determines the date and time for the old form
    Else
    'If a date and time value are not present in the old form
    OldDateCheck = InStr(1, OpenBody, "Date:", vbTextCompare)

    If OldDateCheck <> 339 Then

    OldDateOpened = "Unknown"
    OldTimeOpened = "Unknown"

    RangeString = "G" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).Value = OldDateOpened

    RangeString = "H" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).Value = OldTimeOpened

    Else

    OldDateOpenTemp = InStr(1, OpenBody, "Details about the incident itself.", vbTextCompare)

    OldDateOpen1 = InStr(OldDateOpenTemp, OpenBody, "Date:", vbTextCompare)
    OldDateOpen2 = InStr(OldDateOpenTemp, OpenBody, "Time:", vbTextCompare)

    OldTimeOpen1 = InStr(OldDateOpenTemp, OpenBody, "Time:", vbTextCompare)
    OldTimeOpen2 = InStr(OldDateOpenTemp, OpenBody, "Incident Type:", vbTextCompare)

    OldDiff = OldDateOpen2 - OldDateOpen1
    OldTimeDiff = OldTimeOpen2 - OldTimeOpen1

    If OldDiff = 0 Or OldDiff < 0 Then

    OldDateOpened = "Unknown"

    Else

    OldDateOpened = Mid(OpenBody, OldDateOpen1 + 6, OldDiff - 7)
    OldTimeOpened = Mid(OpenBody, OldTimeOpen1 + 6, OldTimeDiff - 7)

    End If

    OldDateOpened2 = Application.WorksheetFunction.Clean(OldDateOpened)
    OldTimeOpened2 = Application.WorksheetFunction.Clean(OldTimeOpened)

    OldDateOpened3 = ChangeOpenDate(OldDateOpened2)
    OldTimeOpened3 = OldChangeOpenTime(OldTimeOpened2)

    RangeString = "G" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = OldDateOpened3
    Sheets("Raw Data").Columns("G").HorizontalAlignment = xlCenter

    RangeString = "H" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = OldTimeOpened3
    Sheets("Raw Data").Columns("H").HorizontalAlignment = xlCenter

    End If
    End If

End Sub

Sub DateTimeClosed(ClosedBody, Modified, Check, Subject)

    If Check = True Then

    NewDateClose1 = InStr(1, ClosedBody, "Date Closed:", vbTextCompare)
    NewDateClose2 = InStr(1, ClosedBody, "Time Closed:", vbTextCompare)
    NewTimeClose1 = InStr(1, ClosedBody, "Time Closed:", vbTextCompare)
    NewTimeClose2 = InStr(1, ClosedBody, "Actions To Be Taken:", vbTextCompare)

    NewDateCloseDiff = NewDateClose2 - NewDateClose1
    NewTimeCloseDiff = NewTimeClose2 - NewTimeClose1

    CloseDate = Mid(ClosedBody, NewDateClose1 + 14, NewDateCloseDiff - 14)
    CloseTime = Mid(ClosedBody, NewTimeClose1 + 14, NewTimeCloseDiff - 14)

    CloseDate2 = Application.WorksheetFunction.Clean(CloseDate)
    CloseTime2 = Application.WorksheetFunction.Clean(CloseTime)
    CloseDate3 = ChangeDateClosed(CloseDate2)
    CloseTime3 = ChangeTimeClosed(CloseTime2)

    Else
    CloseDate3 = "Not recorded"
    CloseTime3 = "Not recorded"

    End If
    
    RangeString = "I" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = CloseDate3
    RangeString = "J" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
    Sheets("Raw Data").Range(RangeString).Value = CloseTime3
    Sheets("Raw Data").Columns("I").HorizontalAlignment = xlCenter
    Sheets("Raw Data").Columns("J").HorizontalAlignment = xlCenter
End Sub

Sub WorkingHours(WorkingBody, WorkingCheck, Subject)

    If WorkingCheck = True Then

    WorkCheck1 = InStr(1, WorkingBody, "How many working hours spend on investigation?", vbTextCompare)
    WorkCheck2 = InStr(1, WorkingBody, "Appendix for communications", vbTextCompare)
    WorkCheck3 = InStr(1, WorkingBody, "Escalated Date:", vbTextCompare)
    
    If WorkCheck3 = 0 Then
    
    WorkDiff = WorkCheck2 - WorkCheck1
    WorkAns = Mid(WorkingBody, WorkCheck1 + 51, WorkDiff - 54)
    
    WorkAns2 = Application.WorksheetFunction.Clean(WorkAns)
    WorkHours1 = ChangeWorking(WorkAns2)

    RangeString = "K" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "0.0"

    If WorkHours1 <> 0 Then
    Sheets("Raw Data").Range(RangeString).Value = WorkHours1

    Else

    Sheets("Raw Data").Range(RangeString).Value = "Unknown"

    End If
    End If
    End If
    
    If WorkingCheck = False Then

    WorkAns = "Not recorded"

    WorkAns2 = Application.WorksheetFunction.Clean(WorkAns)

    WorkHours1 = ChangeWorking(WorkAns2)

    RangeString = "K" & Trim(Str(RowCount))
    Sheets("Raw Data").Range(RangeString).NumberFormat = "0.0"
    
    If WorkHours1 <> 0 Then
    Sheets("Raw Data").Range(RangeString).Value = WorkHours1

    Else

    Sheets("Raw Data").Range(RangeString).Value = "Not recorded"

    End If
    End If

    Sheets("Raw Data").Columns("K").HorizontalAlignment = xlCenter
    
    If WorkCheck3 <> 0 Then
    
    
    WorkDiff = WorkCheck3 - WorkCheck1
    WorkAns = Mid(WorkingBody, WorkCheck1 + 51, WorkDiff - 51)
    WorkAns2 = Application.WorksheetFunction.Clean(WorkAns)
    WorkHours1 = ChangeWorking(WorkAns2)
    
    RangeString = "K" & Trim(Str(RowCount))

    Sheets("Raw Data").Range(RangeString).NumberFormat = "0.0"
    
    If WorkHours1 <> 0 Then
    Sheets("Raw Data").Range(RangeString).Value = WorkHours1

    Else

    Sheets("Raw Data").Range(RangeString).Value = "Unknown"
    
    End If
    
    End If
    
    Sheets("Raw Data").Columns("K").HorizontalAlignment = xlCenter
    
    End If


End Sub


Sub CreateDashboard()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "Dashboard" Then ws.Delete
Next

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "Dashboard"
Worksheets("Dashboard").Visible = True

Sheets("Dashboard").Range("I1").Value = "Security Incident Response Dashboard"
Sheets("Dashboard").Range("I1").Select
Selection.Font.Bold = True
Selection.Font.Size = 18

Sheets("Dashboard").Range("C3").Value = "30 Most Recent Events/Incidents"
Sheets("Dashboard").Range("C3").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("T3").Value = "Incident Types - Previous 30 days"
Sheets("Dashboard").Range("T3").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("J3").Value = "Event Types - Previous 30 days"
Sheets("Dashboard").Range("J3").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("K22").Value = "12 Month Event Trend"
Sheets("Dashboard").Range("K22").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("U22").Value = "12 Month Incident Trend"
Sheets("Dashboard").Range("U22").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("B5").Value = "Record Number"
Sheets("Dashboard").Range("C5").Value = "Type"
Sheets("Dashboard").Range("D5").Value = "Date Reported"
Sheets("Dashboard").Range("E5").Value = "Status"

Range("B5:E5").Select
    Selection.Font.Bold = True
    Columns("B:B").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    
Sheets("Dashboard").Range("B5:E5").HorizontalAlignment = xlCenter

'Yellow line on top of dashboard
Range("A1:Z1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

  Range("H41:Y56").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("P41").Select
    ActiveCell.FormulaR1C1 = "METRICS"
    
    Range("I42").Select
    ActiveCell.FormulaR1C1 = "Mean Time to Resolution"
       
    Range("O42").Select
    ActiveCell.FormulaR1C1 = "Mean Time to Reduce Risk"
        
    Range("T42").Select
    ActiveCell.FormulaR1C1 = "Mean Time Working on Occurrence"
    
    Range("I48").Select
     ActiveCell.FormulaR1C1 = "This metric is the average time in working"
     
    Range("I49").Select
     ActiveCell.FormulaR1C1 = "hours between when the occurrence was"
     
    Range("I50").Select
     ActiveCell.FormulaR1C1 = "reported to when the occurrence record was"
    
    Range("I51").Select
     ActiveCell.FormulaR1C1 = "closed. The metric provided is for occurences"
     
    Range("I52").Select
     ActiveCell.FormulaR1C1 = "in the previous three months from today's date."
        
    Range("O48").Select
     ActiveCell.FormulaR1C1 = "This metric is the average time in calendar"
     
    Range("O49").Select
     ActiveCell.FormulaR1C1 = "hours between when the occurrence was"
     
    Range("O50").Select
     ActiveCell.FormulaR1C1 = "reported to when the occurrence record was"
    
    Range("O51").Select
     ActiveCell.FormulaR1C1 = "closed (including weekends). The metric"
     
    Range("O52").Select
     ActiveCell.FormulaR1C1 = "provided is for occurences in the previous"
     
    Range("O53").Select
     ActiveCell.FormulaR1C1 = "three months from today's date."
        
    Range("U48").Select
     ActiveCell.FormulaR1C1 = "This metric provides in an indication of the"
     
    Range("U49").Select
     ActiveCell.FormulaR1C1 = "average working hours Information Security have"
     
    Range("U50").Select
     ActiveCell.FormulaR1C1 = "taken to manage and handle a given occurence."
    
    Range("U51").Select
     ActiveCell.FormulaR1C1 = "The metric provided is for occurences in the"
     
    Range("U52").Select
     ActiveCell.FormulaR1C1 = "previous three months from today's date."

  Range("O48:U53").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
        
    Range("I48:M52").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    Range("I48:M52").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
          
    Range("I42:M42").Select
        With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    Selection.Font.Bold = True
    
    Range("P41").Select
    Selection.Font.Bold = True
    
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
   
   Range("O42").Select
    Selection.Font.Bold = True
    
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
Range("T42").Select
    Selection.Font.Bold = True
    
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
Set dd_object = ActiveSheet.DropDowns.Add(Left:=480, Top:=670, Width:=144, Height:=32)
    With dd_object
     .Name = "DropDownBox1"
    End With
ActiveSheet.Shapes.Range(Array("DropDownBox1")).Select
    Selection.OnAction = "MTERSelect"

Set dd_object_3 = ActiveSheet.DropDowns.Add(Left:=770, Top:=670, Width:=144, Height:=32)
    With dd_object_3
     .Name = "DropDownBox3"
    End With
ActiveSheet.Shapes.Range(Array("DropDownBox3")).Select
    Selection.OnAction = "MTRERSelect"

Set dd_object_5 = ActiveSheet.DropDowns.Add(Left:=1050, Top:=670, Width:=144, Height:=32)
    With dd_object_5
     .Name = "DropDownBox5"
    End With

ActiveSheet.Shapes.Range(Array("DropDownBox5")).Select
    Selection.OnAction = "MTWESelect"
 
Set dd_object_7 = ActiveSheet.DropDowns.Add(Left:=180, Top:=555, Width:=144, Height:=32)
    With dd_object_7
     .Name = "DropDownBox7"
    End With

ActiveSheet.Shapes.Range(Array("DropDownBox7")).Select
    Selection.OnAction = "CallRecordStatus"
    
Call PopulateDropDowns

Columns("A:A").Select
    Selection.ColumnWidth = 2.57

Columns("Y:Y").ColumnWidth = 10

Rows("42:42").RowHeight = 39
    
Range("A1:Z57").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

Worksheets("Dashboard").Visible = False
Application.DisplayAlerts = True
Application.ScreenUpdating = True
    
End Sub

'This will collect all the date from Raw Data sheet between the two dates in the Config sheet, this is only data which have complete records - no missing information
Sub Collect_Attack_Type_Data()

Dim DateFrom As Date
Dim DateTo As Date
Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "Attack_Type_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "Attack_Type_Data"
Worksheets("Attack_Type_Data").Visible = False

DateTo = Date
DateFrom = DateDiff("d", 30, DateTo)

Set ws2 = Sheets("Raw Data")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

Sheets("Attack_Type_Data").Cells(1, 1).Value = "Subject"
Sheets("Attack_Type_Data").Cells(1, 2).Value = "Category"
Sheets("Attack_Type_Data").Cells(1, 3).Value = "Date Reported"
Sheets("Attack_Type_Data").Cells(1, 4).Value = "Time Reported"
Sheets("Attack_Type_Data").Cells(1, 5).Value = "Date Discovered"
Sheets("Attack_Type_Data").Cells(1, 6).Value = "Status"
Sheets("Attack_Type_Data").Cells(1, 7).Value = "Date Opened"
Sheets("Attack_Type_Data").Cells(1, 8).Value = "Time Opened"
Sheets("Attack_Type_Data").Cells(1, 9).Value = "Date Closed"
Sheets("Attack_Type_Data").Cells(1, 10).Value = "Time Closed"
Sheets("Attack_Type_Data").Cells(1, 11).Value = "Working Hours"

Sheets("Attack_Type_Data").Range("A1:K1").HorizontalAlignment = xlCenter
Sheets("Attack_Type_Data").Range("A1:K1").Font.Bold = True
Sheets("Attack_Type_Data").UsedRange.Columns.AutoFit

RangeString = "C2:C" & Trim(Str(maxRows))

SubString = "A2:A" & Trim(Str(maxRows))
CatString = "B2:B" & Trim(Str(maxRows))
DateRepString = "C2:C" & Trim(Str(maxRows))
TimeRepString = "D2:D" & Trim(Str(maxRows))
DateDiscString = "E2:E" & Trim(Str(maxRows))
DateOpenString = "G2:G" & Trim(Str(maxRows))
TimeOpenString = "H2:H" & Trim(Str(maxRows))
DateClString = "I2:I" & Trim(Str(maxRows))
TimeClString = "J2:J" & Trim(Str(maxRows))
WorkingString = "K2:K" & Trim(Str(maxRows))

Set r = Sheets("Raw Data").Range(RangeString)

mynumber = 2
rawdatanumber = 2

         For Each cell In r

            SubString = "A" & Trim(Str(rawdatanumber))
            CatString = "B" & Trim(Str(rawdatanumber))
            DateRepString = "C" & Trim(Str(rawdatanumber))
            TimeRepString = "D" & Trim(Str(rawdatanumber))
            DateDiscString = "E" & Trim(Str(rawdatanumber))
            DateOpenString = "G" & Trim(Str(rawdatanumber))
            TimeOpenString = "H" & Trim(Str(rawdatanumber))
            DateClString = "I" & Trim(Str(rawdatanumber))
            TimeClString = "J" & Trim(Str(rawdatanumber))
            WorkingString = "K" & Trim(Str(rawdatanumber))

            SubCheck = Sheets("Raw Data").Range(SubString).Value
            CatCheck = Sheets("Raw Data").Range(CatString).Value
            DateRepCheck = Sheets("Raw Data").Range(DateRepString).Value
            TimeRepCheck = Sheets("Raw Data").Range(TimeRepString).Value
            DateDiscCheck = Sheets("Raw Data").Range(DateDiscString).Value
            DateOpenCheck = Sheets("Raw Data").Range(DateOpenString).Value
            TimeOpenCheck = Sheets("Raw Data").Range(TimeOpenString).Value
            DateClCheck = Sheets("Raw Data").Range(DateClString).Value
            TimeClCheck = Sheets("Raw Data").Range(TimeClString).Value
            WorkingCheck = Sheets("Raw Data").Range(WorkingString).Value

            If IsDate(DateRepCheck) Then

            LowerDateDiff = DateDiff("d", DateFrom, cell.Value)
            UpperDateDiff = DateDiff("d", DateTo, cell.Value)

            If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then

            CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":K" & Trim(Str(rawdatanumber))
            CopyRangeTo = "A" & Trim(Str(mynumber)) & ":K" & Trim(Str(mynumber))
            Sheets("Attack_Type_Data").Range(CopyRangeTo).NumberFormat = "@"
            Sheets("Attack_Type_Data").Range(CopyRangeTo).Value = Sheets("Raw Data").Range(CopyRangeFrom).Value

            mynumber = mynumber + 1

            End If
            End If
              rawdatanumber = rawdatanumber + 1
         Next

Sheets("Attack_Type_Data").UsedRange.Columns.AutoFit

End Sub

Sub BreakDownEvents()

Dim EventCounter As Long
Dim counter As Integer
Dim ArrEvents(500) As Events

EventCounter = 0

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "EventsBreakDown" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "EventsBreakDown"
Worksheets("EventsBreakDown").Visible = False

Sheets("EventsBreakDown").Cells(1, 1).Value = "Event Type"
Sheets("EventsBreakDown").Cells(1, 2).Value = "Count"

Sheets("EventsBreakDown").UsedRange.Columns.AutoFit
Sheets("EventsBreakDown").Range("A1:E1").HorizontalAlignment = xlCenter
Sheets("EventsBreakDown").Range("A1:E1").Font.Bold = True

Set ws2 = Sheets("Attack_Type_Data")
  With ws2
    maxRows = .Range("B" & .Rows.count).End(xlUp).Row
  End With

For counter = 2 To maxRows

    RangeString = "B" & Trim(Str(counter))
    TypeTemp = Sheets("Attack_Type_Data").Range(RangeString).Value
    EventTest = InStr(1, TypeTemp, "Event -", vbTextCompare)

If EventTest = 1 Then

    EventCounter = EventCounter + 1

    found = "N"

    For Arr_Event_Counter = 0 To maxRows

            If ArrEvents(Arr_Event_Counter).EventName = TypeTemp Then

                found = "Y"

                ArrEvents(Arr_Event_Counter).EventCount = ArrEvents(Arr_Event_Counter).EventCount + 1

            End If
     Next

     If found = "N" Then

        For Arr_Event_Counter = 0 To maxRows

            If ArrEvents(Arr_Event_Counter).EventName = "" Then
                ArrEvents(Arr_Event_Counter).EventName = TypeTemp
                ArrEvents(Arr_Event_Counter).EventCount = 1
                GoTo Temp
            End If
        Next
      End If
End If

Temp:
Next

   CurrentPosition = 2

    For DisplayCounter1 = 0 To EventCounter

    RangeStringEvents1 = "A" & Trim(Str(CurrentPosition))
    RangeStringEvents2 = "B" & Trim(Str(CurrentPosition))

    If ArrEvents(DisplayCounter1).EventName <> "" Then

          Sheets("EventsBreakDown").Range(RangeStringEvents1) = ArrEvents(DisplayCounter1).EventName
          Sheets("EventsBreakDown").Range(RangeStringEvents2) = ArrEvents(DisplayCounter1).EventCount
    End If

    CurrentPosition = CurrentPosition + 1
   Next

Sheets("EventsBreakDown").UsedRange.Columns.AutoFit

Set ws3 = Sheets("EventsBreakDown")
  With ws3
    maxRows = .Range("A" & .Rows.count).End(xlUp).Row
  End With

For i = 2 To maxRows

RangeString = "A" & Trim(Str(i))

OldEventType = Sheets("EventsBreakDown").Range(RangeString).Value

NewEventType = Mid(OldEventType, 9)

If NewEventType = "Policies" Then

NewEventType = "Policies, Process or Procedure Event"

End If

Sheets("EventsBreakDown").Range(RangeString).Value = NewEventType

Next

Columns("B:B").Select
    ActiveWorkbook.Worksheets("EventsBreakDown").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("EventsBreakDown").Sort.SortFields.Add Key:=Range( _
        "B1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("EventsBreakDown").Sort
        .SetRange Range("A2:B6")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

''This function gives you a breakdown of the incidents which have occured during the period given in the config
Sub BreakDownIncidents()

Dim IncidentCounter As Long
Dim counter As Integer
Dim ArrIncidents(500) As Incidents

IncidentCounter = 0

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "IncidentsBreakDown" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "IncidentsBreakDown"
Worksheets("IncidentsBreakDown").Visible = False

Sheets("IncidentsBreakDown").Cells(1, 1).Value = "Incident Type"
Sheets("IncidentsBreakDown").Cells(1, 2).Value = "Count"
Sheets("IncidentsBreakDown").UsedRange.Columns.AutoFit
Sheets("IncidentsBreakDown").Range("A1:E1").HorizontalAlignment = xlCenter
Sheets("IncidentsBreakDown").Range("A1:E1").Font.Bold = True

Set ws2 = Sheets("Attack_Type_Data")
  With ws2
    maxRows = .Range("B" & .Rows.count).End(xlUp).Row
  End With

For counter = 2 To maxRows

    RangeString = "B" & Trim(Str(counter))
    TypeTemp = Sheets("Attack_Type_Data").Range(RangeString).Value
    IncidentTest = InStr(1, TypeTemp, "Incident -", vbTextCompare)

If IncidentTest = 1 Then

    IncidentCounter = IncidentCounter + 1
    found = "N"

    For Arr_Incident_Counter = 0 To maxRows

            If ArrIncidents(Arr_Incident_Counter).IncidentName = TypeTemp Then

                found = "Y"
                ArrIncidents(Arr_Incident_Counter).IncidentCount = ArrIncidents(Arr_Incident_Counter).IncidentCount + 1
            End If
     Next

     If found = "N" Then

        For Arr_Incident_Counter = 0 To maxRows

            If ArrIncidents(Arr_Incident_Counter).IncidentName = "" Then

                ArrIncidents(Arr_Incident_Counter).IncidentName = TypeTemp
                ArrIncidents(Arr_Incident_Counter).IncidentCount = 1
                GoTo Temp
            End If
        Next
      End If
End If

Temp:
Next

   CurrentPosition = 2

   'Loop to display events results in sheet
   For DisplayCounter1 = 0 To IncidentCounter

    RangeStringEvents1 = "A" & Trim(Str(CurrentPosition))
    RangeStringEvents2 = "B" & Trim(Str(CurrentPosition))

    If ArrIncidents(DisplayCounter1).IncidentName <> "" Then

          Sheets("IncidentsBreakDown").Range(RangeStringEvents1) = ArrIncidents(DisplayCounter1).IncidentName
          Sheets("IncidentsBreakDown").Range(RangeStringEvents2) = ArrIncidents(DisplayCounter1).IncidentCount
    End If

    CurrentPosition = CurrentPosition + 1

   Next

Sheets("IncidentsBreakDown").UsedRange.Columns.AutoFit

Set ws3 = Sheets("IncidentsBreakDown")
  With ws3
    maxRows = .Range("B" & .Rows.count).End(xlUp).Row
  End With

For i = 2 To maxRows

RangeString = "A" & Trim(Str(i))

OldIncidentType = Sheets("IncidentsBreakDown").Range(RangeString).Value

NewIncidentType = Mid(OldIncidentType, 12)

If NewIncidentType = "Policies" Then

NewIncidentType = "Policies, Process or Procedure Violation"

End If

Sheets("IncidentsBreakDown").Range(RangeString).Value = NewIncidentType

Next

Columns("B:B").Select
    ActiveWorkbook.Worksheets("IncidentsBreakDown").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("IncidentsBreakDown").Sort.SortFields.Add Key:= _
        Range("B1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("IncidentsBreakDown").Sort
        .SetRange Range("A2:B3")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub IncidentEvents_Chart()

Application.ScreenUpdating = False

Worksheets("EventsBreakDown").Activate

Set ws = Sheets("EventsBreakDown")
  With ws
    maxRows = .Range("B" & .Rows.count).End(xlUp).Row
  End With

RangeString = "'EventsBreakDown'!$A$1:$B$" & Trim(Str(maxRows))
Range(RangeString).Select 'actively select the pie chart source cells

Worksheets("Dashboard").Activate

ActiveSheet.Shapes.AddChart.Select 'add the chart and keep it selected as the focus
ActiveChart.ChartType = xlPie ' set its type
ActiveChart.SetSourceData Source:=Range(RangeString) 'set source data
ActiveChart.PlotBy = xlColumns

ActiveChart.ChartArea.Height = 250
ActiveChart.ChartArea.Width = 415
ActiveChart.ChartArea.Left = 345
ActiveChart.ChartArea.Top = 58

ActiveChart.Rotation = 40
ActiveChart.AutoScaling = True

ActiveChart.ChartTitle.Select
Selection.Delete

ActiveChart.SeriesCollection(1).Select
ActiveChart.SeriesCollection(1).ApplyDataLabels

ActiveChart.Legend.Select
Selection.Position = xlLeft
Selection.Format.TextFrame2.TextRange.Font.Size = 11
Selection.Height = 237.139
Selection.Top = 3.43
Selection.Height = 247.139

End Sub

Sub Attack_Incidents_Chart()

Application.ScreenUpdating = False

Worksheets("IncidentsBreakDown").Activate

Set ws = Sheets("IncidentsBreakDown")
  With ws
    maxRows = .Range("B" & .Rows.count).End(xlUp).Row
  End With

RangeString = "'IncidentsBreakDown'!$A$1:$B$" & Trim(Str(maxRows))
Range(RangeString).Select 'actively select the pie chart source cells

Worksheets("Dashboard").Activate

ActiveSheet.Shapes.AddChart.Select 'add the chart and keep it selected as the focus
ActiveChart.ChartType = xlPie  ' set its type
'ActiveChart.ChartType = xlBarClustered
ActiveChart.SetSourceData Source:=Range(RangeString) 'set source data

ActiveChart.ChartArea.Height = 250
ActiveChart.ChartArea.Width = 415
ActiveChart.ChartArea.Left = 825
ActiveChart.ChartArea.Top = 58

ActiveChart.Rotation = 40
ActiveChart.AutoScaling = True

ActiveChart.ChartTitle.Select
Selection.Delete

ActiveChart.SeriesCollection(1).Select
ActiveChart.SeriesCollection(1).ApplyDataLabels

ActiveChart.Legend.Select
Selection.Position = xlLeft
Selection.Format.TextFrame2.TextRange.Font.Size = 11
Selection.Height = 237.139
Selection.Top = 3.43
Selection.Height = 247.139

If maxRows = 2 Then
ActiveChart.ChartArea.Select
    ActiveChart.PlotBy = xlColumns
    ActiveChart.Legend.Select
    Selection.Width = 94.894

End If

End Sub

Sub MostRecentChart()

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "RecentTemp" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "RecentTemp"
Worksheets("RecentTemp").Visible = False

copynumber = 1

For i = 2 To 31

CopyRangeFrom = "A" & Trim(Str(i)) & ":K" & Trim(Str(i))
CopyRangeTo = "A" & Trim(Str(copynumber)) & ":K" & Trim(Str(copynumber))
Sheets("RecentTemp").Range(CopyRangeTo).NumberFormat = "@"

Sheets("RecentTemp").Range(CopyRangeTo).Value = Sheets("Raw Data").Range(CopyRangeFrom).Value

copynumber = copynumber + 1
Next

Set ws3 = Sheets("RecentTemp")
  With ws3
    maxRows = .Range("A" & .Rows.count).End(xlUp).Row
  End With

For i = 1 To maxRows

RangeString = "A" & Trim(Str(i))
Temp = Sheets("RecentTemp").Range(RangeString).Value
ClosedTest = InStr(1, Temp, "<CLOSED>", vbTextCompare)

If ClosedTest = 1 Then

OldSub = Sheets("RecentTemp").Range(RangeString).Value
NewSub = Mid(OldSub, 1, 17)
Else

OldSub = Sheets("RecentTemp").Range(RangeString).Value
NewSub = Mid(OldSub, 1, 8)
End If

Sheets("RecentTemp").Range(RangeString).Value = NewSub

RangeString2 = "B" & Trim(Str(i))
Temp2 = Sheets("RecentTemp").Range(RangeString2).Value

If Temp2 = "Event - Data Subject Access Request (DSAR)" Then

Sheets("RecentTemp").Range(RangeString2).Value = "Event - DSAR"

End If

Next

copynumber2 = 6

For i = 1 To 30

CopyRangeFrom2 = "A" & Trim(Str(i)) & ":C" & Trim(Str(i))
StatusFrom = "F" & Trim(Str(i))
CopyRangeTo2 = "B" & Trim(Str(copynumber2)) & ":D" & Trim(Str(copynumber2))
StatusTo = "E" & Trim(Str(copynumber2))

CopyValue = "E" & Trim(Str(copynumber2))

Sheets("Dashboard").Range(CopyRangeTo2).NumberFormat = "@"
Sheets("Dashboard").Range(StatusTo).NumberFormat = "@"

Sheets("Dashboard").Range(CopyRangeTo2).Value = Sheets("RecentTemp").Range(CopyRangeFrom2).Value
Sheets("Dashboard").Range(StatusTo).Value = Sheets("RecentTemp").Range(StatusFrom).Value

CheckValue = Sheets("Dashboard").Range(CopyValue).Value

CatString = "B" & Trim(Str(i))
DateRepString = "C" & Trim(Str(i))
TimeRepString = "D" & Trim(Str(i))
DateDiscString = "E" & Trim(Str(i))
DateOpenString = "G" & Trim(Str(i))
TimeOpenString = "H" & Trim(Str(i))
DateClString = "I" & Trim(Str(i))
TimeClString = "J" & Trim(Str(i))
WorkingString = "K" & Trim(Str(i))

CatCheck = Sheets("RecentTemp").Range(CatString).Value
DateRepCheck = Sheets("RecentTemp").Range(DateRepString).Value
TimeRepCheck = Sheets("RecentTemp").Range(TimeRepString).Value
DateDiscCheck = Sheets("RecentTemp").Range(DateDiscString).Value
DateOpenCheck = Sheets("RecentTemp").Range(DateOpenString).Value
TimeOpenCheck = Sheets("RecentTemp").Range(TimeOpenString).Value
DateClCheck = Sheets("RecentTemp").Range(DateClString).Value
TimeClCheck = Sheets("RecentTemp").Range(TimeClString).Value
WorkingCheck = Sheets("RecentTemp").Range(WorkingString).Value

Worksheets("Dashboard").Activate

If CheckValue = "Open" Then

    Sheets("Dashboard").Range(CopyRangeTo2).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Sheets("Dashboard").Range(StatusTo).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End If

If CheckValue = "Closed" Then
    
    If CatCheck = "Unknown" Or DateRepCheck = "Unknown" Or TimeRepCheck = "Unknown" Or DateDiscCheck = "Unknown" Or DateOpenCheck = "Unknown" Or TimeOpenCheck = "Unknown" Or DateClCheck = "Unknown" Or TimeClCheck = "Unknown" Or WorkingCheck = "Unknown" Then
        
    Sheets("Dashboard").Range(CopyRangeTo2).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Sheets("Dashboard").Range(StatusTo).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Else
    
    Sheets("Dashboard").Range(CopyRangeTo2).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 7667457
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Sheets("Dashboard").Range(StatusTo).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 7667457
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If
End If

copynumber2 = copynumber2 + 1

Next

Range("B6:E35").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit

    Range("B6:E35").Select
        With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Range("B6:E35").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Range("B6:E6").Select
Selection.Font.Bold = False

End Sub

Sub DataForPastTwelveMonths()

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "BreakDownLastSixMonths" Then ws.Delete
Next

Dim ws5 As Worksheet
For Each ws5 In Worksheets
If ws5.Name = "TempBreakDown" Then ws5.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "BreakDownLastSixMonths"
Worksheets("BreakDownLastSixMonths").Visible = False
Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "TempBreakDown"
Worksheets("TempBreakDown").Visible = False

Sheets("BreakDownLastSixMonths").Cells(1, 1).Value = "Period"
Sheets("BreakDownLastSixMonths").Cells(1, 2).Value = "Number of Events"
Sheets("BreakDownLastSixMonths").Cells(1, 3).Value = "Average"
Sheets("BreakDownLastSixMonths").Cells(1, 4).Value = "Period"
Sheets("BreakDownLastSixMonths").Cells(1, 5).Value = "Number of Incidents"
Sheets("BreakDownLastSixMonths").Cells(1, 6).Value = "Average"

Sheets("BreakDownLastSixMonths").UsedRange.Columns.AutoFit
Sheets("BreakDownLastSixMonths").Range("A1:F10").HorizontalAlignment = xlCenter
Sheets("BreakDownLastSixMonths").Range("A1:F1").Font.Bold = True

Set ws2 = Sheets("Raw Data")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("Raw Data").Range(RangeString)

DateTemp = Date

DayValue = Format(DateTemp, "dd")

If DayValue <> 1 Then

Date1 = DateAdd("d", -DayValue + 1, DateTemp)

Else

Date1 = DateAdd("d", -1, DateTemp)

End If

Date2 = DateAdd("m", -12, Date1)

rawdatanumber = 2
mynumber = 2

   For Each cell In r

    If IsDate(cell.Value) Then

        TempVar = cell.Value
        LowerDateDiff = DateDiff("d", Date2, cell.Value)
        UpperDateDiff = DateDiff("d", Date1, cell.Value)

        If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then

        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":C" & Trim(Str(rawdatanumber))
        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":C" & Trim(Str(mynumber))
        Sheets("TempBreakDown").Range(CopyRangeTo).NumberFormat = "@"
        Sheets("TempBreakDown").Range(CopyRangeTo).Value = Sheets("Raw Data").Range(CopyRangeFrom).Value

        mynumber = mynumber + 1
        End If

        End If
        rawdatanumber = rawdatanumber + 1
Next

Set ws3 = Sheets("TempBreakDown")
  With ws3
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))

Set r = Sheets("TempBreakDown").Range(RangeString)

OutputNo1 = 2

            For i = -12 To -1

            TempDate = DateAdd("m", i, Date1)
            Date2 = DateAdd("m", 1, TempDate)

            OutputString = "A" & Trim(Str(OutputNo1))
            DateOneMonth = Format(TempDate, "mmm")
            DateOneyear = Format(TempDate, "yy")

            MonthDisplay = DateOneMonth & "-" & DateOneyear

            Sheets("BreakDownLastSixMonths").Range(OutputString).NumberFormat = "@"
            Sheets("BreakDownLastSixMonths").Range(OutputString).Value = MonthDisplay

            OutputString = "D" & Trim(Str(OutputNo1))
            Sheets("BreakDownLastSixMonths").Range(OutputString).NumberFormat = "@"
            Sheets("BreakDownLastSixMonths").Range(OutputString).Value = MonthDisplay

            rawdatanumber = 2
            PeriodEventcounter = 0
            PeriodIncidentCounter = 0

                For Each cell In r

                    If IsDate(cell.Value) Then
                        TempVar = cell.Value
                        LowerDateDiff = DateDiff("d", TempDate, cell.Value)
                        UpperDateDiff = DateDiff("d", Date2, cell.Value)

                        RangeStringCat = "B" & Trim(Str(rawdatanumber))
                        TypeTemp = Sheets("TempBreakDown").Range(RangeStringCat).Value
                        EventTest = InStr(1, TypeTemp, "Event -", vbTextCompare)
                        IncidentTest = InStr(1, TypeTemp, "Incident -", vbTextCompare)

                            If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventTest = 1 Then

                            mynumber = mynumber + 1
                            PeriodEventcounter = PeriodEventcounter + 1

                            ElseIf LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentTest = 1 Then

                            mynumber = mynumber + 1
                            PeriodIncidentCounter = PeriodIncidentCounter + 1

                            End If
                    End If
                rawdatanumber = rawdatanumber + 1

                Next
                OutputString = "B" & Trim(Str(OutputNo1))
                Sheets("BreakDownLastSixMonths").Range(OutputString).Value = PeriodEventcounter

                OutputString = "E" & Trim(Str(OutputNo1))
                Sheets("BreakDownLastSixMonths").Range(OutputString).Value = PeriodIncidentCounter

                OutputNo1 = OutputNo1 + 1
               Next

Sheets("TempBreakDown").UsedRange.Columns.AutoFit
Sheets("BreakDownLastSixMonths").UsedRange.Columns.AutoFit
Sheets("BreakDownLastSixMonths").Activate

Columns("A:F").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Range("C2").Select
ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R13C2)"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C13"), Type:=xlFillDefault
Range("C2:C10").Select

Range("F2").Select
ActiveCell.FormulaR1C1 = "=AVERAGE(R2C5:R13C5)"
Range("F2").Select
Selection.AutoFill Destination:=Range("F2:F13"), Type:=xlFillDefault
Range("F2:F10").Select

End Sub

Sub CreateAverageCharts()

Sheets("BreakDownLastSixMonths").Activate

Set ws2 = Sheets("BreakDownLastSixMonths")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "'BreakDownLastSixMonths'!$A$1:$C$" & Trim(Str(maxRows))
Range(RangeString).Select

Worksheets("Dashboard").Activate

ActiveSheet.Shapes.AddChart.Select
ActiveChart.ChartType = xlColumnClustered
ActiveChart.SetSourceData Source:=Range(RangeString)

ActiveChart.ChartArea.Height = 250
ActiveChart.ChartArea.Width = 415
ActiveChart.ChartArea.Left = 475
ActiveChart.ChartArea.Top = 348

ActiveChart.Legend.Select
ActiveChart.SeriesCollection(2).Select
ActiveChart.SeriesCollection(2).ChartType = xlLine
ActiveChart.ApplyLayout (3)
ActiveChart.ChartTitle.Select
Selection.Delete

ActiveChart.SeriesCollection(1).Select
ActiveChart.SeriesCollection(1).ApplyDataLabels

Sheets("BreakDownLastSixMonths").Activate

RangeString2 = "'BreakDownLastSixMonths'!$D$1:$F$" & Trim(Str(maxRows))
Range(RangeString2).Select

Worksheets("Dashboard").Activate

ActiveSheet.Shapes.AddChart.Select
ActiveChart.ChartType = xlColumnClustered
ActiveChart.SetSourceData Source:=Range(RangeString2)

ActiveChart.ChartArea.Height = 250
ActiveChart.ChartArea.Width = 415
ActiveChart.ChartArea.Left = 955
ActiveChart.ChartArea.Top = 348

ActiveChart.Legend.Select
ActiveChart.SeriesCollection(2).Select
ActiveChart.SeriesCollection(2).ChartType = xlLine
ActiveChart.ApplyLayout (3)
ActiveChart.ChartTitle.Select
Selection.Delete

ActiveChart.SeriesCollection(1).Select
ActiveChart.SeriesCollection(1).ApplyDataLabels

End Sub

Sub CollectCompleteRecords()

Dim DateFrom As Date
Dim DateTo As Date
Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "CompleteRecords" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "CompleteRecords"
Worksheets("CompleteRecords").Visible = False

Date1 = Date
Date2 = "23/02/2014"

Set ws2 = Sheets("Raw Data")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

Sheets("CompleteRecords").Cells(1, 1).Value = "Subject"
Sheets("CompleteRecords").Cells(1, 2).Value = "Category"
Sheets("CompleteRecords").Cells(1, 3).Value = "Date Reported"
Sheets("CompleteRecords").Cells(1, 4).Value = "Time Reported"
Sheets("CompleteRecords").Cells(1, 5).Value = "Date Discovered"
Sheets("CompleteRecords").Cells(1, 6).Value = "Status"
Sheets("CompleteRecords").Cells(1, 7).Value = "Date Opened"
Sheets("CompleteRecords").Cells(1, 8).Value = "Time Opened"
Sheets("CompleteRecords").Cells(1, 9).Value = "Date Closed"
Sheets("CompleteRecords").Cells(1, 10).Value = "Time Closed"
Sheets("CompleteRecords").Cells(1, 11).Value = "Working Hours"

Sheets("CompleteRecords").Range("A1:K1").HorizontalAlignment = xlCenter
Sheets("CompleteRecords").Range("A1:K1").Font.Bold = True
Sheets("CompleteRecords").UsedRange.Columns.AutoFit

RangeString = "C2:C" & Trim(Str(maxRows))

SubString = "A2:A" & Trim(Str(maxRows))
CatString = "B2:B" & Trim(Str(maxRows))
DateRepString = "C2:C" & Trim(Str(maxRows))
TimeRepString = "D2:D" & Trim(Str(maxRows))
DateDiscString = "E2:E" & Trim(Str(maxRows))
DateOpenString = "G2:G" & Trim(Str(maxRows))
TimeOpenString = "H2:H" & Trim(Str(maxRows))
DateClString = "I2:I" & Trim(Str(maxRows))
TimeClString = "J2:J" & Trim(Str(maxRows))
WorkingString = "K2:K" & Trim(Str(maxRows))

Set r = Sheets("Raw Data").Range(RangeString)

mynumber = 2
rawdatanumber = 2

         For Each cell In r

            SubString = "A" & Trim(Str(rawdatanumber))
            CatString = "B" & Trim(Str(rawdatanumber))
            DateRepString = "C" & Trim(Str(rawdatanumber))
            TimeRepString = "D" & Trim(Str(rawdatanumber))
            DateDiscString = "E" & Trim(Str(rawdatanumber))
            DateOpenString = "G" & Trim(Str(rawdatanumber))
            TimeOpenString = "H" & Trim(Str(rawdatanumber))
            DateClString = "I" & Trim(Str(rawdatanumber))
            TimeClString = "J" & Trim(Str(rawdatanumber))
            WorkingString = "K" & Trim(Str(rawdatanumber))

            SubCheck = Sheets("Raw Data").Range(SubString).Value
            CatCheck = Sheets("Raw Data").Range(CatString).Value
            DateRepCheck = Sheets("Raw Data").Range(DateRepString).Value
            TimeRepCheck = Sheets("Raw Data").Range(TimeRepString).Value
            DateDiscCheck = Sheets("Raw Data").Range(DateDiscString).Value
            DateOpenCheck = Sheets("Raw Data").Range(DateOpenString).Value
            TimeOpenCheck = Sheets("Raw Data").Range(TimeOpenString).Value
            DateClCheck = Sheets("Raw Data").Range(DateClString).Value
            TimeClCheck = Sheets("Raw Data").Range(TimeClString).Value
            WorkingCheck = Sheets("Raw Data").Range(WorkingString).Value

            If SubCheck <> "Unknown" And CatCheck <> "Unknown" And DateRepCheck <> "Unknown" And TimeRepCheck <> "Unknown" And DateDiscCheck <> "Unknown" And DateOpenCheck <> "Unknown" And TimeOpenCheck <> "Unknown" And DateClCheck <> "Unknown" And TimeClCheck <> "Unknown" And WorkingCheck <> "Unknown" Then

            LowerDateDiff = DateDiff("d", Date2, cell.Value)
            UpperDateDiff = DateDiff("d", Date1, cell.Value)

            If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then

            CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":K" & Trim(Str(rawdatanumber))
            CopyRangeTo = "A" & Trim(Str(mynumber)) & ":K" & Trim(Str(mynumber))

            Sheets("CompleteRecords").Range(CopyRangeTo).NumberFormat = "@"
            Sheets("CompleteRecords").Range(CopyRangeTo).Value = Sheets("Raw Data").Range(CopyRangeFrom).Value

            mynumber = mynumber + 1

            End If

            End If
              rawdatanumber = rawdatanumber + 1
         Next

Sheets("CompleteRecords").UsedRange.Columns.AutoFit
Worksheets("Dashboard").Activate

End Sub

Sub LoadNewTrendData()

Dim directory As String, fileName As String, sheet As Worksheet, total As Integer
Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")

Application.ScreenUpdating = False
Application.DisplayAlerts = False

fso.copyfolder "C:\test", "C:\Dashboard_Temp"

directory = "C:\Dashboard_Temp\"
fileName = Dir(directory & "*.????")

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "Trend_Chart_Data" Then ws.Delete
Next

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "Trend_Chart_Data"
Worksheets("Trend_Chart_Data").Visible = True

Sheets("Trend_Chart_Data").Cells(1, 1).Value = "Date"
Sheets("Trend_Chart_Data").Cells(1, 2).Value = "Number of Open Records"
Sheets("Trend_Chart_Data").Cells(1, 3).Value = "Number of Closed and Incomplete Records"

mynumber = 2

Do While fileName <> ""
    Workbooks.Open (directory & fileName)

        maxRows = 0

        Set ws2 = Workbooks(fileName).Sheets(1)
        With ws2
        maxRows = .Range("F" & .Rows.count).End(xlUp).Row
        End With

        RangeString = "F2:F" & Trim(Str(maxRows))
        Set r = Workbooks(fileName).Sheets(1).Range(RangeString)

        SubString = "A2:A" & Trim(Str(maxRows))
        CatString = "B2:B" & Trim(Str(maxRows))
        DateRepString = "C2:C" & Trim(Str(maxRows))
        TimeRepString = "D2:D" & Trim(Str(maxRows))
        DateDiscString = "E2:E" & Trim(Str(maxRows))
        DateOpenString = "G2:G" & Trim(Str(maxRows))
        TimeOpenString = "H2:H" & Trim(Str(maxRows))
        DateClString = "I2:I" & Trim(Str(maxRows))
        TimeClString = "J2:J" & Trim(Str(maxRows))
        WorkingString = "K2:K" & Trim(Str(maxRows))

        OpenRecords = 0
        Badrecords = 0
        rawdatanumber = 2

            For Each cell In r

                SubString = "A" & Trim(Str(rawdatanumber))
                CatString = "B" & Trim(Str(rawdatanumber))
                DateRepString = "C" & Trim(Str(rawdatanumber))
                TimeRepString = "D" & Trim(Str(rawdatanumber))
                DateDiscString = "E" & Trim(Str(rawdatanumber))
                DateOpenString = "G" & Trim(Str(rawdatanumber))
                TimeOpenString = "H" & Trim(Str(rawdatanumber))
                DateClString = "I" & Trim(Str(rawdatanumber))
                TimeClString = "J" & Trim(Str(rawdatanumber))
                WorkingString = "K" & Trim(Str(rawdatanumber))

                SubCheck = Workbooks(fileName).Sheets(1).Range(SubString).Value
                CatCheck = Workbooks(fileName).Sheets(1).Range(CatString).Value
                DateRepCheck = Workbooks(fileName).Sheets(1).Range(DateRepString).Value
                TimeRepCheck = Workbooks(fileName).Sheets(1).Range(TimeRepString).Value
                DateDiscCheck = Workbooks(fileName).Sheets(1).Range(DateDiscString).Value
                DateOpenCheck = Workbooks(fileName).Sheets(1).Range(DateOpenString).Value
                TimeOpenCheck = Workbooks(fileName).Sheets(1).Range(TimeOpenString).Value
                DateClCheck = Workbooks(fileName).Sheets(1).Range(DateClString).Value
                TimeClCheck = Workbooks(fileName).Sheets(1).Range(TimeClString).Value
                WorkingCheck = Workbooks(fileName).Sheets(1).Range(WorkingString).Value

                myDate1 = Format(DateRepCheck, "dd/mm/20yy")
                myDate2 = Format(DateDiscCheck, "dd/mm/20yy")
                myDate3 = Format(DateOpenCheck, "dd/mm/20yy")
                myDate4 = Format(DateClCheck, "dd/mm/20yy")

                myTime1 = Format(TimeRepCheck, "##:##")
                myTime2 = Format(TimeOpenCheck, "##:##")
                myTime3 = Format(TimeClCheck, "##:##")

                If cell.Value = "Open" Then

                    OpenRecords = OpenRecords + 1

                    ElseIf cell.Value = "Closed" And SubCheck = "Unknown" Or CatCheck = "Unknown" Or DateRepCheck = "Unknown" Or TimeRepCheck = "Unknown" Or DateDiscCheck = "Unknown" Or DateOpenCheck = "Unknown" Or TimeOpenCheck = "Unknown" Or DateClCheck = "Unknown" Or TimeClCheck = "Unknown" Or WorkingCheck = "Unknown" Or SubCheck = "" Or CatCheck = "" Or DateRepCheck = "" Or TimeRepCheck = "" Or DateDiscCheck = "" Or DateOpenCheck = "" Or TimeOpenCheck = "" Or DateClCheck = "" Or TimeClCheck = "" Or WorkingCheck = "" Or Not IsDate(DateRepCheck) Or Not IsDate(DateOpenCheck) Or Not IsDate(DateClCheck) Then

                    Badrecords = Badrecords + 1
                    
                End If
                
                rawdatanumber = rawdatanumber + 1

                Next

    Workbooks("Incident Response Dashboard").Worksheets("Trend_Chart_Data").Activate

    WriteDate = "A" & Trim(Str(mynumber))
    WriteOpen = "B" & Trim(Str(mynumber))
    WriteBad = "C" & Trim(Str(mynumber))

    Temp = fileName
    
    If Right$(Temp, 1) = "s" Then
    
    Temp = Left$(Temp, Len(Temp) - 4)
    
    ElseIf Right$(Temp, 1) = "x" Then

    Temp = Left$(Temp, Len(Temp) - 5)
    
    End If
    
    Temp = Mid(Temp, 5)
    
    Sheets("Trend_Chart_Data").Range(WriteDate).Value = Temp
    Sheets("Trend_Chart_Data").Range(WriteOpen).Value = OpenRecords
    Sheets("Trend_Chart_Data").Range(WriteBad).Value = Badrecords

    mynumber = mynumber + 1
    
    LastFileName = fileName
    
    Workbooks(fileName).Close
    fileName = Dir()
    
Loop

OutputRangeString = "B2:C" & Trim(Str(maxRows))

Sheets("Trend_Chart_Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("Trend_Chart_Data").Range(OutputRangeString).HorizontalAlignment = xlCenter
Sheets("Trend_Chart_Data").Range("A1:C1").Font.Bold = True
Sheets("Trend_Chart_Data").UsedRange.Columns.AutoFit

Application.ScreenUpdating = True
Application.DisplayAlerts = True

Set ws2 = Sheets("Trend_Chart_Data")
  With ws2
    maxRows3 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

LastRecordOpen = "B" & Trim(Str(maxRows3))
LastRecordIncomp = "C" & Trim(Str(maxRows3))

NumberOpen = Sheets("Trend_Chart_Data").Range(LastRecordOpen).Value
NumberIncomp = Sheets("Trend_Chart_Data").Range(LastRecordIncomp).Value

Worksheets("Dashboard").Activate

Sheets("Dashboard").Range("B37").Value = "Total Open Records: " & NumberOpen
Sheets("Dashboard").Range("B37").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("B38").Value = "Total Closed and Incomplete Records: " & NumberIncomp
Sheets("Dashboard").Range("B38").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Range("B37:C37").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range("B38:C38").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
LastRecordData (LastFileName)

MyPath = "C:\Dashboard_Temp\"
    
    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If
    If fso.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If
    fso.deletefolder MyPath

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Sub LastRecordData(PassedFileName)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim directory As String, fileName As String, sheet As Worksheet, total As Integer

directory = "C:\test\"
fileName = Dir(directory & PassedFileName)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "Incomplete Records" Then ws2.Delete
Next

Dim ws3 As Worksheet
For Each ws3 In Worksheets
If ws3.Name = "Open Records" Then ws3.Delete
Next

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "Incomplete Records"
Worksheets("Incomplete Records").Visible = False

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "Open Records"
Worksheets("Open Records").Visible = False

Sheets("Incomplete Records").Cells(1, 1).Value = "Subject"
Sheets("Incomplete Records").Cells(1, 2).Value = "Category"
Sheets("Incomplete Records").Cells(1, 3).Value = "Date Reported"
Sheets("Incomplete Records").Cells(1, 4).Value = "Time Reported"
Sheets("Incomplete Records").Cells(1, 5).Value = "Date Discovered"
Sheets("Incomplete Records").Cells(1, 6).Value = "Status"
Sheets("Incomplete Records").Cells(1, 7).Value = "Date Opened"
Sheets("Incomplete Records").Cells(1, 8).Value = "Time Opened"
Sheets("Incomplete Records").Cells(1, 9).Value = "Date Closed"
Sheets("Incomplete Records").Cells(1, 10).Value = "Time Closed"
Sheets("Incomplete Records").Cells(1, 11).Value = "Working Hours"
Sheets("Incomplete Records").Cells(1, 12).Value = "Handler Name"

Sheets("Incomplete Records").Range("A1:L1").HorizontalAlignment = xlCenter
Sheets("Incomplete Records").Range("A1:L1").Font.Bold = True

Sheets("Open Records").Cells(1, 1).Value = "Subject"
Sheets("Open Records").Cells(1, 2).Value = "Category"
Sheets("Open Records").Cells(1, 3).Value = "Date Reported"
Sheets("Open Records").Cells(1, 4).Value = "Time Reported"
Sheets("Open Records").Cells(1, 5).Value = "Date Discovered"
Sheets("Open Records").Cells(1, 6).Value = "Status"
Sheets("Open Records").Cells(1, 7).Value = "Date Opened"
Sheets("Open Records").Cells(1, 8).Value = "Time Opened"
Sheets("Open Records").Cells(1, 9).Value = "Date Closed"
Sheets("Open Records").Cells(1, 10).Value = "Time Closed"
Sheets("Open Records").Cells(1, 11).Value = "Working Hours"
Sheets("Open Records").Cells(1, 12).Value = "Handler Name"

Sheets("Open Records").Range("A1:L1").HorizontalAlignment = xlCenter
Sheets("Open Records").Range("A1:L1").Font.Bold = True

Workbooks.Open (directory & fileName)

 maxRows = 0

        Set ws = Workbooks(fileName).Sheets(1)
        With ws
        maxRows = .Range("F" & .Rows.count).End(xlUp).Row
        End With

RangeString = "F2:F" & Trim(Str(maxRows))
        Set r = Workbooks(fileName).Sheets(1).Range(RangeString)

SubString = "A2:A" & Trim(Str(maxRows))
        CatString = "B2:B" & Trim(Str(maxRows))
        DateRepString = "C2:C" & Trim(Str(maxRows))
        TimeRepString = "D2:D" & Trim(Str(maxRows))
        DateDiscString = "E2:E" & Trim(Str(maxRows))
        DateOpenString = "G2:G" & Trim(Str(maxRows))
        TimeOpenString = "H2:H" & Trim(Str(maxRows))
        DateClString = "I2:I" & Trim(Str(maxRows))
        TimeClString = "J2:J" & Trim(Str(maxRows))
        WorkingString = "K2:K" & Trim(Str(maxRows))
        HandlerString = "L2:L" & Trim(Str(maxRows))

        OpenRecords = 2
        Badrecords = 2
        rawdatanumber = 2
 
 For Each cell In r

                SubString = "A" & Trim(Str(rawdatanumber))
                CatString = "B" & Trim(Str(rawdatanumber))
                DateRepString = "C" & Trim(Str(rawdatanumber))
                TimeRepString = "D" & Trim(Str(rawdatanumber))
                StatusString = "F" & Trim(Str(rawdatanumber))
                DateDiscString = "E" & Trim(Str(rawdatanumber))
                DateOpenString = "G" & Trim(Str(rawdatanumber))
                TimeOpenString = "H" & Trim(Str(rawdatanumber))
                DateClString = "I" & Trim(Str(rawdatanumber))
                TimeClString = "J" & Trim(Str(rawdatanumber))
                WorkingString = "K" & Trim(Str(rawdatanumber))
                HandlerString = "L" & Trim(Str(rawdatanumber))

                SubCheck = Workbooks(fileName).Sheets(1).Range(SubString).Value
                CatCheck = Workbooks(fileName).Sheets(1).Range(CatString).Value
                DateRepCheck = Workbooks(fileName).Sheets(1).Range(DateRepString).Value
                TimeRepCheck = Workbooks(fileName).Sheets(1).Range(TimeRepString).Value
                DateDiscCheck = Workbooks(fileName).Sheets(1).Range(DateDiscString).Value
                DateOpenCheck = Workbooks(fileName).Sheets(1).Range(DateOpenString).Value
                TimeOpenCheck = Workbooks(fileName).Sheets(1).Range(TimeOpenString).Value
                DateClCheck = Workbooks(fileName).Sheets(1).Range(DateClString).Value
                TimeClCheck = Workbooks(fileName).Sheets(1).Range(TimeClString).Value
                WorkingCheck = Workbooks(fileName).Sheets(1).Range(WorkingString).Value
                HandlerCheck = Workbooks(fileName).Sheets(1).Range(HandlerString).Value
                StatusCheck = Workbooks(fileName).Sheets(1).Range(StatusString).Value

                myDate1 = Format(DateRepCheck, "dd/mm/20yy")
                myDate2 = Format(DateDiscCheck, "dd/mm/20yy")
                myDate3 = Format(DateOpenCheck, "dd/mm/20yy")
                myDate4 = Format(DateClCheck, "dd/mm/20yy")

                myTime1 = Format(TimeRepCheck, "##:##")
                myTime2 = Format(TimeOpenCheck, "##:##")
                myTime3 = Format(TimeClCheck, "##:##")

         If StatusCheck = "Open" Then

                    CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":L" & Trim(Str(rawdatanumber))
                    CopyRangeTo = "A" & Trim(Str(OpenRecords)) & ":L" & Trim(Str(OpenRecords))

                    Workbooks("Incident Response Dashboard - Modified version of deployed").Worksheets("Open Records").Activate

                    Sheets("Open Records").Range(CopyRangeTo).NumberFormat = "@"
                    Sheets("Open Records").Range(CopyRangeTo).Value = Sheets(1).Range(CopyRangeFrom).Value

                    OpenRecords = OpenRecords + 1

                    ElseIf StatusCheck = "Closed" And SubCheck = "Unknown" Or CatCheck = "Unknown" Or DateRepCheck = "Unknown" Or TimeRepCheck = "Unknown" Or DateDiscCheck = "Unknown" Or DateOpenCheck = "Unknown" Or TimeOpenCheck = "Unknown" Or DateClCheck = "Unknown" Or TimeClCheck = "Unknown" Or WorkingCheck = "Unknown" Or SubCheck = "" Or CatCheck = "" Or DateRepCheck = "" Or TimeRepCheck = "" Or DateDiscCheck = "" Or DateOpenCheck = "" Or TimeOpenCheck = "" Or DateClCheck = "" Or TimeClCheck = "" Or WorkingCheck = "" Or Not IsDate(DateRepCheck) Or Not IsDate(DateOpenCheck) Or Not IsDate(DateClCheck) Then

                    CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":L" & Trim(Str(rawdatanumber))
                    CopyRangeTo = "A" & Trim(Str(Badrecords)) & ":L" & Trim(Str(Badrecords))
                    
                    Workbooks("Incident Response Dashboard - Modified version of deployed").Worksheets("Incomplete Records").Activate
                    
                    Sheets("Incomplete Records").Range(CopyRangeTo).NumberFormat = "@"
                    Sheets("Incomplete Records").Range(CopyRangeTo).Value = Sheets(1).Range(CopyRangeFrom).Value
                    
                    Badrecords = Badrecords + 1
                
        End If
                
                rawdatanumber = rawdatanumber + 1

                Next

    Workbooks(fileName).Close
    fileName = Dir()

Sheets("Open Records").UsedRange.Columns.AutoFit
Sheets("Incomplete Records").UsedRange.Columns.AutoFit

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Sub Open_Closed_Trend_Chart()

Worksheets("Trend_Chart_Data").Activate

Set ws = Sheets("Trend_Chart_Data")
  With ws
    maxRows = .Range("A" & .Rows.count).End(xlUp).Row
  End With

RangeString = "'Trend_Chart_Data'!$A$1:$C$" & Trim(Str(maxRows))

Range(RangeString).Select

Worksheets("Dashboard").Activate

ActiveSheet.Shapes.AddChart.Select
ActiveChart.ChartType = xlLine
ActiveChart.SetSourceData Source:=Range(RangeString)

ActiveChart.ChartArea.Height = 290
ActiveChart.ChartArea.Width = 442
ActiveChart.ChartArea.Left = 15
ActiveChart.ChartArea.Top = 600

ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MajorUnit = 10
ActiveChart.Axes(xlValue).MajorUnit = 5

ActiveChart.Axes(xlValue).MajorGridlines.Select
ActiveChart.ChartArea.Select
ActiveChart.ApplyLayout (1)

ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue

ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Security Incident Record Evolution"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Security Incident Record Evolution"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 34).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    
    ActiveChart.ApplyLayout (3)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    
'This code below loads the number of open/closed records when the data is reloaded

Worksheets("Trend_Chart_Data").Activate

Set ws2 = Sheets("Trend_Chart_Data")
  With ws2
    maxRows3 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

LastRecordOpen = "B" & Trim(Str(maxRows3))
LastRecordIncomp = "C" & Trim(Str(maxRows3))

NumberOpen = Sheets("Trend_Chart_Data").Range(LastRecordOpen).Value
NumberIncomp = Sheets("Trend_Chart_Data").Range(LastRecordIncomp).Value

Worksheets("Dashboard").Activate

Sheets("Dashboard").Range("B37").Value = "Total Open Records: " & NumberOpen
Sheets("Dashboard").Range("B37").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Sheets("Dashboard").Range("B38").Value = "Total Closed and Incomplete Records: " & NumberIncomp
Sheets("Dashboard").Range("B38").Select
Selection.Font.Bold = True
Selection.Font.Size = 13

Range("B37:C37").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

Range("B38:C38").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub SnapShotDb()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim FolderPath As String, path As String, FileType As String, count As Integer

FolderPath = "C:\test"

path = FolderPath & "\*.????"

fileName = Dir(path)

    Do While fileName <> ""
       count = count + 1
        fileName = Dir()
    Loop

NumberofFIle = count + 1

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "TempDBSheet" Then ws.Delete
Next

Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "RawDataTemp" Then ws2.Delete
Next

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "TempDBSheet"
Worksheets("TempDBSheet").Visible = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "RawDataTemp"
Worksheets("RawDataTemp").Visible = False

Worksheets("Raw Data").Activate

Worksheets("Raw Data").UsedRange.Copy
Worksheets("RawDataTemp").Paste

Sheets("RawDataTemp").UsedRange.Columns.AutoFit

Worksheets("RawDataTemp").Activate

Set ws2 = Sheets("RawDataTemp")
  With ws2
    maxRows = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows = 1 Then

MsgBox "No data has been loaded, the data will be loaded now", vbCritical, "Error - No Data"

Call Load_Data

End If

found = 0

For i = 2 To maxRows
    
    StringRange = "A" & Trim(Str(i))
    CatString = "B" & Trim(Str(i))
    DateRepString = "C" & Trim(Str(i))
    TimeRepString = "D" & Trim(Str(i))
    DateDiscString = "E" & Trim(Str(i))
    StatusString = "F" & Trim(Str(i))
    DateOpenString = "G" & Trim(Str(i))
    TimeOpenString = "H" & Trim(Str(i))
    DateClString = "I" & Trim(Str(i))
    TimeClString = "J" & Trim(Str(i))
    WorkingString = "K" & Trim(Str(i))

    SubCheck = Sheets("RawDataTemp").Range(StringRange).Value
    CatCheck = Sheets("RawDataTemp").Range(CatString).Value
    DateRepCheck = Sheets("RawDataTemp").Range(DateRepString).Value
    TimeRepCheck = Sheets("RawDataTemp").Range(TimeRepString).Value
    DateDiscCheck = Sheets("RawDataTemp").Range(DateDiscString).Value
    StatusCheck = Sheets("RawDataTemp").Range(StatusString).Value
    DateOpenCheck = Sheets("RawDataTemp").Range(DateOpenString).Value
    TimeOpenCheck = Sheets("RawDataTemp").Range(TimeOpenString).Value
    DateClCheck = Sheets("RawDataTemp").Range(DateClString).Value
    TimeClCheck = Sheets("RawDataTemp").Range(TimeClString).Value
    WorkingCheck = Sheets("RawDataTemp").Range(WorkingString).Value
    
    If SubCheck = "2014-023<CLOSED> EDisclosure Hadingham & Sons" Then
    
    found = 1
    Exit For
    
    Else
    
    CopyRangeFrom = "A" & Trim(Str(i)) & ":K" & Trim(Str(i))
    CopyRangeTo = "A" & Trim(Str(i)) & ":K" & Trim(Str(i))
    Sheets("TempDBSheet").Range(CopyRangeTo).NumberFormat = "@"
    Sheets("TempDBSheet").Range(CopyRangeTo).Value = Sheets("RawDataTemp").Range(CopyRangeFrom).Value
    
    found = 0
    
    End If
    
    Next

Worksheets("TempDBSheet").Activate
Sheets("TempDBSheet").UsedRange.Columns.AutoFit

DayTemp = Format("d", Date)

Select Case DayTemp

    Case 1, 21, 31
    
    DayEnd = "st"
    
    Case 2, 22
    
    DayEnd = "nd"
    
    Case 3, 23
    
    DayEnd = "rd"
    
    Case Else
    
    DayEnd = "th"

End Select

NewFileType = "Excel Files 2007 (*.xlsx), *.xlsx," & _
               "All files (*.*), *.*"

FileLocation = "C:\test\" & NumberofFIle & " " & Format(Date, "d") & DayEnd & " " & Format(Date, "mmmm yyyy")

ActiveSheet.Copy
    With ActiveWorkbook
        .SaveAs FileLocation
        .Close 0
    End With

Worksheets("TempDBSheet").Visible = False
Application.DisplayAlerts = True
Application.ScreenUpdating = True
  
End Sub

Sub ShowRawData()

Worksheets("Raw Data").Visible = True
Worksheets("Raw Data").Activate

End Sub

Sub HideRawData()

Worksheets("Raw Data").Visible = False
Worksheets("Dashboard").Visible = True
Worksheets("Dashboard").Activate

End Sub

Sub OpenMTRER(EventType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTRER Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRER Data"
Worksheets("MTRER Data").Visible = False

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTRER Data").Cells(1, 1).Value = "Subject"
Sheets("MTRER Data").Cells(1, 2).Value = "Category"
Sheets("MTRER Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTRER Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTRER Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTRER Data").Cells(1, 6).Value = "Time Closed"
Sheets("MTRER Data").Cells(1, 7).Value = "Number of Hours"

Sheets("MTRER Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTRER Data").Range("A1:G1").Font.Bold = True
Sheets("MTRER Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r
               
               EventCheckString = "B" & Trim(Str(counter))
               EventCheck = Sheets("CompleteRecords").Range(EventCheckString).Value
                          
               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
               If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = EventType Then
               
                    CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                    CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))
        
                    Sheets("MTRER Data").Range(CopyRangeTo).NumberFormat = "@"
                    Sheets("MTRER Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
        
                    ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                    ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))
        
                    Sheets("MTRER Data").Range(ClosedStringTo).NumberFormat = "@"
                    Sheets("MTRER Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
        
                    DateRepRange = "C" & Trim(Str(mynumber))
                    DateReported = Sheets("MTRER Data").Range(DateRepRange)
                    TimeRepRange = "D" & Trim(Str(mynumber))
                    TimeReported = Sheets("MTRER Data").Range(TimeRepRange)
        
                    CombReported = DateReported & " " & TimeReported
                    DateCloRange = "E" & Trim(Str(mynumber))
                    DateClosed = Sheets("MTRER Data").Range(DateCloRange)
                    TimeCloRange = "F" & Trim(Str(mynumber))
                    TimeClosed = Sheets("MTRER Data").Range(TimeCloRange)
                    CombClosed = DateClosed & " " & TimeClosed
        
                    HoursDiff = DateDiff("h", CombReported, CombClosed)
        
                    HoursRange = "G" & Trim(Str(mynumber))
                    Sheets("MTRER Data").Range(HoursRange) = HoursDiff
        
                    If HoursDiff = 0 Then
        
                    MinutesDiff = DateDiff("n", TimeReported, TimeClosed)
                    MinutesDiff2 = MinutesDiff / 60
        
                    HoursRange = "G" & Trim(Str(mynumber))
                    Sheets("MTRER Data").Range(HoursRange) = MinutesDiff2
                    End If
        
                    mynumber = mynumber + 1
        
                    End If
        
                      rawdatanumber = rawdatanumber + 1
                      counter = counter + 1
                 Next

Sheets("MTRER Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTRER Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows2 >= 2 Then

MTREventsCount = 0

For i = 2 To maxRows2

RangeString = "G" & Trim(Str(i))
Temp = Sheets("MTRER Data").Range(RangeString).Value
MTREventsCount = MTREventsCount + Temp

Next

maxRows2 = maxRows2 - 1
MTREREvents = MTREventsCount / maxRows2
MTREREvents = Math.Round(MTREREvents, 2)

EventType = Mid(EventType, 8)

MsgBox ("The MTRR for " & maxRows2 & " " & EventType & " events for the previous three months is: " & MTREREvents & " hours")

Else

MsgBox ("No events to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1
End Sub

Sub MTERSelect()

Dim EventTemp As String
Dim IncidentTemp As String

If dropdownindex <> Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex Then

  dropdownindex = Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex

 Select Case dropdownindex

    Case 2
    EventTemp = "Event - Audit"
    Call OpenMTER(EventTemp)
    
    Case 3
    EventTemp = "Event - Customer Dispute"
    Call OpenMTER(EventTemp)
    
    Case 4
    EventTemp = "Event - Data Subject Access Request (DSAR)"
    Call OpenMTER(EventTemp)
    
    Case 5
    EventTemp = "Event - Date Loss Event"
    Call OpenMTER(EventTemp)
    
    Case 6
    EventTemp = "Event - E-Disclosure"
    Call OpenMTER(EventTemp)
    
    Case 7
    EventTemp = "Event - Equipment Theft/Loss"
    Call OpenMTER(EventTemp)
    
    Case 8
    EventTemp = "Event - Human Resources Investigation"
    Call OpenMTER(EventTemp)
    
    Case 9
    EventTemp = "Event - Policies"
    Call OpenMTER(EventTemp)
    
    Case 10
    EventTemp = "Event - Internal Usage Investigation"
    Call OpenMTER(EventTemp)
    
    Case 11
    EventTemp = "Event - Regulatory Investigation"
    Call OpenMTER(EventTemp)
    
    Case 12
    EventTemp = "Event - Security Assistance"
    Call OpenMTER(EventTemp)
       
    Case 13
    IncidentTemp = "Incident - Data Exposure"
    Call OpenMTIR(IncidentTemp)
    
    Case 14
    IncidentTemp = "Incident - Fraudulent Activity"
    Call OpenMTIR(IncidentTemp)
    
    Case 15
    IncidentTemp = "Incident - Malware Incidents"
    Call OpenMTIR(IncidentTemp)
    
    Case 16
    IncidentTemp = "Incident - Policies"
    Call OpenMTIR(IncidentTemp)
    
    Case 17
    IncidentTemp = "Incident - Service Outage"
    Call OpenMTIR(IncidentTemp)
    
    Case 18
    IncidentTemp = "Incident - Unauthorised Access to Information"
    Call OpenMTIR(IncidentTemp)
    
    Case 19
    IncidentTemp = "Incident - Unauthorised Modification of Information"
    Call OpenMTIR(IncidentTemp)
    
    Case 20
    Call MTERAllEvents
    
    Case 21
    Call MTERAllIncidents
    
    Case 22
    Call MTERAll

End Select
End If
End Sub

Sub MTWESelect()

Dim EventTemp As String
Dim IncidentTemp As String

If dropdownindex <> Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex Then

  dropdownindex = Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex

 Select Case dropdownindex
 
 Case 2
    EventTemp = "Event - Audit"
    Call OpenMTWE(EventTemp)
    
    Case 3
    EventTemp = "Event - Customer Dispute"
    Call OpenMTWE(EventTemp)
    
    Case 4
    EventTemp = "Event - Data Subject Access Request (DSAR)"
    Call OpenMTWE(EventTemp)
    
    Case 5
    EventTemp = "Event - Date Loss Event"
    Call OpenMTWE(EventTemp)
    
    Case 6
    EventTemp = "Event - E-Disclosure"
    Call OpenMTWE(EventTemp)
    
    Case 7
    EventTemp = "Event - Equipment Theft/Loss"
    Call OpenMTWE(EventTemp)
    
    Case 8
    EventTemp = "Event - Human Resources Investigation"
    Call OpenMTWE(EventTemp)
    
    Case 9
    EventTemp = "Event - Policies"
    Call OpenMTWE(EventTemp)
    
    Case 10
    EventTemp = "Event - Internal Usage Investigation"
    Call OpenMTWE(EventTemp)
    
    Case 11
    EventTemp = "Event - Regulatory Investigation"
    Call OpenMTWE(EventTemp)
    
    Case 12
    EventTemp = "Event - Security Assistance"
    Call OpenMTWE(EventTemp)
       
    Case 13
    IncidentTemp = "Incident - Data Exposure"
    Call OpenMTWI(IncidentTemp)
    
    Case 14
    IncidentTemp = "Incident - Fraudulent Activity"
    Call OpenMTWI(IncidentTemp)
    
    Case 15
    IncidentTemp = "Incident - Malware Incidents"
    Call OpenMTWI(IncidentTemp)
    
    Case 16
    IncidentTemp = "Incident - Policies"
    Call OpenMTWI(IncidentTemp)
    
    Case 17
    IncidentTemp = "Incident - Service Outage"
    Call OpenMTWI(IncidentTemp)
    
    Case 18
    IncidentTemp = "Incident - Unauthorised Access to Information"
    Call OpenMTWI(IncidentTemp)
    
    Case 19
    IncidentTemp = "Incident - Unauthorised Modification of Information"
    Call OpenMTWI(IncidentTemp)
    
    Case 20
    Call MTWEAllEvents
    
    Case 21
    Call MTWEAllIncidents
    
    Case 22
    Call MTWEAll

End Select
End If
End Sub

Sub MTRERSelect()

Dim EventTemp2 As String
Dim IncidentTemp2 As String

If dropdownindex <> Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex Then

  dropdownindex = Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex

 Select Case dropdownindex

    Case 2
    EventTemp2 = "Event - Audit"
    Call OpenMTRER(EventTemp2)
    
    Case 3
    EventTemp2 = "Event - Customer Dispute"
    Call OpenMTRER(EventTemp2)
    
    Case 4
    EventTemp2 = "Event - Data Subject Access Request (DSAR)"
    Call OpenMTRER(EventTemp2)
    
    Case 5
    EventTemp2 = "Event - Date Loss Event"
    Call OpenMTRER(EventTemp2)
    
    Case 6
    EventTemp2 = "Event - E-Disclosure"
    Call OpenMTRER(EventTemp2)
    
    Case 7
    EventTemp2 = "Event - Equipment Theft/Loss"
    Call OpenMTRER(EventTemp2)
    
    Case 8
    EventTemp2 = "Event - Human Resources Investigation"
    Call OpenMTRER(EventTemp2)
    
    Case 9
    EventTemp2 = "Event - Policies"
    Call OpenMTRER(EventTemp2)
    
    Case 10
    EventTemp2 = "Event - Internal Usage Investigation"
    Call OpenMTRER(EventTemp2)
    
    Case 11
    EventTemp2 = "Event - Regulatory Investigation"
    Call OpenMTRER(EventTemp2)
    
    Case 12
    EventTemp2 = "Event - Security Assistance"
    Call OpenMTRER(EventTemp2)
       
    Case 13
    IncidentTemp2 = "Incident - Data Exposure"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 14
    IncidentTemp2 = "Incident - Fraudulent Activity"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 15
    IncidentTemp2 = "Incident - Malware Incidents"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 16
    IncidentTemp2 = "Incident - Policies"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 17
    IncidentTemp2 = "Incident - Service Outage"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 18
    IncidentTemp2 = "Incident - Unauthorised Access to Information"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 19
    IncidentTemp2 = "Incident - Unauthorised Modification of Information"
    Call OpenMTRIR(IncidentTemp2)
    
    Case 20
    Call MTRERAllEvents

    Case 21
    Call MTRERAllIncidents

    Case 22
    Call All_MTRER

End Select
End If
End Sub

Sub OpenMTER(EventType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTER_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTER_Real_Data"
Worksheets("MTER_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTER Data" Then ws2.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTER Data"
Worksheets("MTER Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTER Data").Cells(1, 1).Value = "Subject"
Sheets("MTER Data").Cells(1, 2).Value = "Category"
Sheets("MTER Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTER Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTER Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTER Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTER Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTER Data").Range("A1:G1").Font.Bold = True
Sheets("MTER Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

            For Each cell In r

               EventCheckString = "B" & Trim(Str(counter))
               EventCheck = Sheets("CompleteRecords").Range(EventCheckString).Value

               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)

               If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = EventType Then

                    CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                    CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                    Sheets("MTER Data").Range(CopyRangeTo).NumberFormat = "@"
                    Sheets("MTER Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

                    ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                    ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                    Sheets("MTER Data").Range(ClosedStringTo).NumberFormat = "@"
                    Sheets("MTER Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value

                    DateRepRange = "C" & Trim(Str(mynumber))
                    DateReported = Sheets("MTER Data").Range(DateRepRange)
                    
                    TimeRepRange = "D" & Trim(Str(mynumber))
                    TimeReported = Sheets("MTER Data").Range(TimeRepRange)

                    DateTimeReported = DateReported & " " & TimeReported
                    DateRepRange = "A" & Trim(Str(mynumber))
                    Sheets("MTER_Real_Data").Range(DateRepRange).NumberFormat = "@"
                    Sheets("MTER_Real_Data").Range(DateRepRange) = DateTimeReported

                    DateCloRange = "E" & Trim(Str(mynumber))
                    DateClosed = Sheets("MTER Data").Range(DateCloRange)

                    TimeCloRange = "F" & Trim(Str(mynumber))
                    TimeClosed = Sheets("MTER Data").Range(TimeCloRange)

                    TimeDateClosed = DateClosed & " " & TimeClosed
                    DateCloRange = "B" & Trim(Str(mynumber))
                    Sheets("MTER_Real_Data").Range(DateCloRange).NumberFormat = "@"
                    Sheets("MTER_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                    MTERRange = "D" & Trim(Str(mynumber))
                    Sheets("MTER_Real_Data").Range(MTERRange).NumberFormat = "0.0"
                    
                    Sheets("MTER_Real_Data").Range(MTERRange).Formula = "=9*(NETWORKDAYS(RC[-3],RC[-2])-1)-24*((MOD(RC[-3],1)-MOD(RC[-2],1)))"

                    mynumber = mynumber + 1

                    End If

                      rawdatanumber = rawdatanumber + 1
                      counter = counter + 1
                 Next

Sheets("MTER_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTER_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofEvents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "D" & Trim(Str(i))
Temp = Sheets("MTER_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofEvents = NumberofEvents + 1

Next

MTERMetric = TotalTime / NumberofEvents
MTERMetric = Math.Round(MTERMetric, 2)

EventType = Mid(EventType, 8)

MsgBox ("The MTR for " & NumberofEvents & " " & EventType & " events for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No events to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1

End Sub

Sub MTERAllEvents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTERAll_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTERAll_Real_Data"
Worksheets("MTERAll_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTER_All Data" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTER_All Data"
Worksheets("MTER_All Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTER_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTER_All Data").Cells(1, 2).Value = "Category"
Sheets("MTER_All Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTER_All Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTER_All Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTER_All Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTER_All Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTER_All Data").Range("A1:G1").Font.Bold = True
Sheets("MTER_All Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               EventCheckPos = "B" & Trim(Str(counter))
               EventCheckString = Sheets("CompleteRecords").Range(EventCheckPos).Value
               EventCheck = InStr(1, EventCheckString, "Event", vbTextCompare)
               
               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = 1 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("MTER_All Data").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("MTER_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("MTER_All Data").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("MTER_All Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("MTER_All Data").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("MTER_All Data").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("MTERAll_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("MTERAll_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("MTER_All Data").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("MTER_All Data").Range(TimeCloRange)
                        
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("MTERAll_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("MTERAll_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTERRange = "D" & Trim(Str(mynumber))
                        Sheets("MTERAll_Real_Data").Range(MTERRange).NumberFormat = "0.0"
                    
                        Sheets("MTERAll_Real_Data").Range(MTERRange).Formula = "=9*(NETWORKDAYS(RC[-3],RC[-2])-1)-24*((MOD(RC[-3],1)-MOD(RC[-2],1)))"

                        mynumber = mynumber + 1

                    End If

                      rawdatanumber = rawdatanumber + 1
                      counter = counter + 1
Next

Sheets("MTERAll_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTERAll_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofEvents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "D" & Trim(Str(i))
Temp = Sheets("MTERAll_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofEvents = NumberofEvents + 1

Next

MTERMetric = TotalTime / NumberofEvents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTR for " & NumberofEvents & " events for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No events to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1

End Sub

Sub MTRERAllEvents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTRERAll_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRERAll_Real_Data"
Worksheets("MTRERAll_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTRER_All Data" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRER_All Data"
Worksheets("MTRER_All Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTRER_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTRER_All Data").Cells(1, 2).Value = "Category"
Sheets("MTRER_All Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTRER_All Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTRER_All Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTRER_All Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTRER_All Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTRER_All Data").Range("A1:G1").Font.Bold = True
Sheets("MTRER_All Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               EventCheckPos = "B" & Trim(Str(counter))
               EventCheckString = Sheets("CompleteRecords").Range(EventCheckPos).Value
               EventCheck = InStr(1, EventCheckString, "Event", vbTextCompare)
               
               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = 1 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("MTRER_All Data").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("MTRER_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("MTRER_All Data").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("MTRER_All Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("MTRER_All Data").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("MTRER_All Data").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("MTRERAll_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("MTRERAll_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("MTRER_All Data").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("MTRER_All Data").Range(TimeCloRange)
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("MTRERAll_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("MTRERAll_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTRERRange = "D" & Trim(Str(mynumber))
                        Sheets("MTRERAll_Real_Data").Range(MTRERRange).NumberFormat = "0.0"
                        
                        HoursDiff = DateDiff("h", DateTimeReported, TimeDateClosed)
                        HoursRange = "C" & Trim(Str(mynumber))
                        Sheets("MTRERAll_Real_Data").Range(HoursRange) = HoursDiff

                       If HoursDiff = 0 Then

                           MinutesDiff = DateDiff("n", TimeReported, TimeClosed)
                           MinutesDiff2 = MinutesDiff / 60

                            HoursRange = "C" & Trim(Str(mynumber))
                            Sheets("MTRERAll_Real_Data").Range(HoursRange) = MinutesDiff2
                        End If
                        
                mynumber = mynumber + 1
                End If

       rawdatanumber = rawdatanumber + 1
       counter = counter + 1
Next

Sheets("MTRERAll_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTRERAll_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofEvents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "C" & Trim(Str(i))
Temp = Sheets("MTRERAll_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofEvents = NumberofEvents + 1

Next

MTERMetric = TotalTime / NumberofEvents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTRR for " & NumberofEvents & " events for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No events to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1

End Sub

Sub MTERAllIncidents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTIRAll_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTIRAll_Real_Data"
Worksheets("MTIRAll_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTIR_All Data" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTIR_All Data"
Worksheets("MTIR_All Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTIR_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTIR_All Data").Cells(1, 2).Value = "Category"
Sheets("MTIR_All Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTIR_All Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTIR_All Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTIR_All Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTIR_All Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTIR_All Data").Range("A1:G1").Font.Bold = True
Sheets("MTIR_All Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               IncidentCheckPos = "B" & Trim(Str(counter))
               IncidentCheckString = Sheets("CompleteRecords").Range(IncidentCheckPos).Value
               IncidentCheck = InStr(1, IncidentCheckString, "Incident", vbTextCompare)
               
               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = 1 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("MTIR_All Data").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("MTIR_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("MTIR_All Data").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("MTIR_All Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("MTIR_All Data").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("MTIR_All Data").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("MTIRAll_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("MTIRAll_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("MTIR_All Data").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("MTIR_All Data").Range(TimeCloRange)
                        
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("MTIRAll_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("MTIRAll_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTERRange = "D" & Trim(Str(mynumber))
                        Sheets("MTIRAll_Real_Data").Range(MTERRange).NumberFormat = "0.0"
                    
                        Sheets("MTIRAll_Real_Data").Range(MTERRange).Formula = "=9*(NETWORKDAYS(RC[-3],RC[-2])-1)-24*((MOD(RC[-3],1)-MOD(RC[-2],1)))"

                        mynumber = mynumber + 1

                    End If

                      rawdatanumber = rawdatanumber + 1
                      counter = counter + 1
Next

Sheets("MTIRAll_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTIRAll_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofIncidents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "D" & Trim(Str(i))
Temp = Sheets("MTIRAll_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofIncidents = NumberofIncidents + 1

Next

MTERMetric = TotalTime / NumberofIncidents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTR for " & NumberofIncidents & " incidents for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No incidents to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1

End Sub

Sub MTRERAllIncidents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTRIRAll_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRIRAll_Real_Data"
Worksheets("MTRIRAll_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTRIR_All Data" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRIR_All Data"
Worksheets("MTRIR_All Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTRIR_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTRIR_All Data").Cells(1, 2).Value = "Category"
Sheets("MTRIR_All Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTRIR_All Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTRIR_All Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTRIR_All Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTRIR_All Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTRIR_All Data").Range("A1:G1").Font.Bold = True
Sheets("MTRIR_All Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               IncidentCheckPos = "B" & Trim(Str(counter))
               IncidentCheckString = Sheets("CompleteRecords").Range(IncidentCheckPos).Value
               IncidentCheck = InStr(1, IncidentCheckString, "Incident", vbTextCompare)
               
               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = 1 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("MTRIR_All Data").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("MTRIR_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("MTRIR_All Data").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("MTRIR_All Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("MTRIR_All Data").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("MTRIR_All Data").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("MTRIRAll_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("MTRIRAll_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("MTRIR_All Data").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("MTRIR_All Data").Range(TimeCloRange)
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("MTRIRAll_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("MTRIRAll_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTRERRange = "D" & Trim(Str(mynumber))
                        Sheets("MTRIRAll_Real_Data").Range(MTRERRange).NumberFormat = "0.0"
                        
                        HoursDiff = DateDiff("h", DateTimeReported, TimeDateClosed)
                        HoursRange = "C" & Trim(Str(mynumber))
                        Sheets("MTRIRAll_Real_Data").Range(HoursRange) = HoursDiff

                       If HoursDiff = 0 Then

                           MinutesDiff = DateDiff("n", TimeReported, TimeClosed)
                           MinutesDiff2 = MinutesDiff / 60

                            HoursRange = "C" & Trim(Str(mynumber))
                            Sheets("MTRIRAll_Real_Data").Range(HoursRange) = MinutesDiff2
                        End If
                        
                mynumber = mynumber + 1
                End If

       rawdatanumber = rawdatanumber + 1
       counter = counter + 1
Next

Sheets("MTRIRAll_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTRIRAll_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofIncidents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "C" & Trim(Str(i))
Temp = Sheets("MTRIRAll_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofIncidents = NumberofIncidents + 1

Next

MTERMetric = TotalTime / NumberofIncidents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTRR for " & NumberofIncidents & " incidents for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No incidents to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1

End Sub

Sub MTERAll()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "All_MTIR_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "All_MTIR_Real_Data"
Worksheets("All_MTIR_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTIR_All Data_2" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTIR_All Data_2"
Worksheets("MTIR_All Data_2").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTIR_All Data_2").Cells(1, 1).Value = "Subject"
Sheets("MTIR_All Data_2").Cells(1, 2).Value = "Category"
Sheets("MTIR_All Data_2").Cells(1, 3).Value = "Date Reported"
Sheets("MTIR_All Data_2").Cells(1, 4).Value = "Time Reported"
Sheets("MTIR_All Data_2").Cells(1, 5).Value = "Date Closed"
Sheets("MTIR_All Data_2").Cells(1, 6).Value = "Time Closed"

Sheets("MTIR_All Data_2").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTIR_All Data_2").Range("A1:G1").Font.Bold = True
Sheets("MTIR_All Data_2").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("MTIR_All Data_2").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("MTIR_All Data_2").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("MTIR_All Data_2").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("MTIR_All Data_2").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("MTIR_All Data_2").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("MTIR_All Data_2").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("All_MTIR_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("All_MTIR_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("MTIR_All Data_2").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("MTIR_All Data_2").Range(TimeCloRange)
                        
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("All_MTIR_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("All_MTIR_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTERRange = "D" & Trim(Str(mynumber))
                        Sheets("All_MTIR_Real_Data").Range(MTERRange).NumberFormat = "0.0"
                    
                        Sheets("All_MTIR_Real_Data").Range(MTERRange).Formula = "=9*(NETWORKDAYS(RC[-3],RC[-2])-1)-24*((MOD(RC[-3],1)-MOD(RC[-2],1)))"

                        mynumber = mynumber + 1

                    End If

                      rawdatanumber = rawdatanumber + 1
                      counter = counter + 1
Next

Sheets("All_MTIR_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("All_MTIR_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofIncidents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "D" & Trim(Str(i))
Temp = Sheets("All_MTIR_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofIncidents = NumberofIncidents + 1

Next

MTERMetric = TotalTime / NumberofIncidents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTR for " & NumberofIncidents & " events and incidents for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No occurrences to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1

End Sub

Sub All_MTRER()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "All_MTRER_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "All_MTRER_Real_Data"
Worksheets("All_MTRER_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "All_MTRER_Data" Then ws2.Delete
Next

Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "All_MTRER_Data"
Worksheets("All_MTRER_Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("All_MTRER_Data").Cells(1, 1).Value = "Subject"
Sheets("All_MTRER_Data").Cells(1, 2).Value = "Category"
Sheets("All_MTRER_Data").Cells(1, 3).Value = "Date Reported"
Sheets("All_MTRER_Data").Cells(1, 4).Value = "Time Reported"
Sheets("All_MTRER_Data").Cells(1, 5).Value = "Date Closed"
Sheets("All_MTRER_Data").Cells(1, 6).Value = "Time Closed"

Sheets("All_MTRER_Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("All_MTRER_Data").Range("A1:G1").Font.Bold = True
Sheets("All_MTRER_Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r

               LowerDateDiff = DateDiff("d", Date2, cell.Value)
               UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
                    If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then
                        
                        CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
                        CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

                        Sheets("All_MTRER_Data").Range(CopyRangeTo).NumberFormat = "@"
                        Sheets("All_MTRER_Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value
                        
                        ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
                        ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))

                        Sheets("All_MTRER_Data").Range(ClosedStringTo).NumberFormat = "@"
                        Sheets("All_MTRER_Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value
                        
                        DateRepRange = "C" & Trim(Str(mynumber))
                        DateReported = Sheets("All_MTRER_Data").Range(DateRepRange)
                    
                        TimeRepRange = "D" & Trim(Str(mynumber))
                        TimeReported = Sheets("All_MTRER_Data").Range(TimeRepRange)
                        
                        DateTimeReported = DateReported & " " & TimeReported
                        DateRepRange = "A" & Trim(Str(mynumber))
                        Sheets("All_MTRER_Real_Data").Range(DateRepRange).NumberFormat = "@"
                        Sheets("All_MTRER_Real_Data").Range(DateRepRange) = DateTimeReported
                        
                        DateCloRange = "E" & Trim(Str(mynumber))
                        DateClosed = Sheets("All_MTRER_Data").Range(DateCloRange)

                        TimeCloRange = "F" & Trim(Str(mynumber))
                        TimeClosed = Sheets("All_MTRER_Data").Range(TimeCloRange)
                        TimeDateClosed = DateClosed & " " & TimeClosed
                        
                        DateCloRange = "B" & Trim(Str(mynumber))
                        Sheets("All_MTRER_Real_Data").Range(DateCloRange).NumberFormat = "@"
                        Sheets("All_MTRER_Real_Data").Range(DateCloRange) = TimeDateClosed
                    
                        MTRERRange = "D" & Trim(Str(mynumber))
                        Sheets("All_MTRER_Real_Data").Range(MTRERRange).NumberFormat = "0.0"
                        
                        HoursDiff = DateDiff("h", DateTimeReported, TimeDateClosed)
                        HoursRange = "C" & Trim(Str(mynumber))
                        Sheets("All_MTRER_Real_Data").Range(HoursRange) = HoursDiff

                       If HoursDiff = 0 Then

                           MinutesDiff = DateDiff("n", TimeReported, TimeClosed)
                           MinutesDiff2 = MinutesDiff / 60

                            HoursRange = "C" & Trim(Str(mynumber))
                            Sheets("All_MTRER_Real_Data").Range(HoursRange) = MinutesDiff2
                        End If
                        
                mynumber = mynumber + 1
                End If

       rawdatanumber = rawdatanumber + 1
       counter = counter + 1
Next

Sheets("All_MTRER_Real_Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("All_MTRER_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With
  
If maxRows2 >= 2 Then
  
NumberofEvents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "C" & Trim(Str(i))
Temp = Sheets("All_MTRER_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofEvents = NumberofEvents + 1

Next

MTERMetric = TotalTime / NumberofEvents
MTERMetric = Math.Round(MTERMetric, 2)

MsgBox ("The MTRR for " & NumberofEvents & " events and incidents for the previous three months is: " & MTERMetric & " hours")

Else

MsgBox ("No events to report for this metric")

End If

Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1

End Sub

Sub OpenMTRIR(IncidentType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTRIR Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTRIR Data"
Worksheets("MTRIR Data").Visible = False

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTRIR Data").Cells(1, 1).Value = "Subject"
Sheets("MTRIR Data").Cells(1, 2).Value = "Category"
Sheets("MTRIR Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTRIR Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTRIR Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTRIR Data").Cells(1, 6).Value = "Time Closed"
Sheets("MTRIR Data").Cells(1, 7).Value = "Number of Hours"

Sheets("MTRIR Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTRIR Data").Range("A1:G1").Font.Bold = True
Sheets("MTRIR Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

For Each cell In r
        
            IncidentCheckString = "B" & Trim(Str(counter))
            IncidentCheck = Sheets("CompleteRecords").Range(IncidentCheckString).Value

             LowerDateDiff = DateDiff("d", Date2, cell.Value)
             UpperDateDiff = DateDiff("d", Date1, cell.Value)

             If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = IncidentType Then

             CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
             CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

             Sheets("MTRIR Data").Range(CopyRangeTo).NumberFormat = "@"
             Sheets("MTRIR Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

             ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
             ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))
             
             Sheets("MTRIR Data").Range(ClosedStringTo).NumberFormat = "@"
             Sheets("MTRIR Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value

             DateRepRange = "C" & Trim(Str(mynumber))
             DateReported = Sheets("MTRIR Data").Range(DateRepRange)

             TimeRepRange = "D" & Trim(Str(mynumber))
             TimeReported = Sheets("MTRIR Data").Range(TimeRepRange)

             DateTimeReported = DateReported & " " & TimeReported

             DateCloRange = "E" & Trim(Str(mynumber))
             DateClosed = Sheets("MTRIR Data").Range(DateCloRange)
             TimeCloRange = "F" & Trim(Str(mynumber))
             TimeClosed = Sheets("MTRIR Data").Range(TimeCloRange)
             
             TimeDateClosed = DateClosed & " " & TimeClosed
             
             HoursDiff = DateDiff("h", DateTimeReported, TimeDateClosed)
             
             HoursRange = "G" & Trim(Str(mynumber))
             Sheets("MTRIR Data").Range(HoursRange) = HoursDiff
                    
             If HoursDiff = 0 Then
        
             MinutesDiff = DateDiff("n", TimeReported, TimeClosed)
             MinutesDiff2 = MinutesDiff / 60
             
             HoursRange = "G" & Trim(Str(mynumber))
             Sheets("MTRIR Data").Range(HoursRange) = MinutesDiff2
             
             End If
                    
             mynumber = mynumber + 1
             End If

             rawdatanumber = rawdatanumber + 1
             counter = counter + 1
         Next

Sheets("MTRIR Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTRIR Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows2 >= 2 Then

MTRIncidentsCount = 0

For i = 2 To maxRows2

RangeString = "G" & Trim(Str(i))
Temp = Sheets("MTRIR Data").Range(RangeString).Value
MTRIncidentsCount = MTRIncidentsCount + Temp

Next

maxRows2 = maxRows2 - 1
MTRIRIncidents = MTRIncidentsCount / maxRows2
MTRIRIncidents = Math.Round(MTRIRIncidents, 2)

IncidentType = Mid(IncidentType, 12)

MsgBox ("The MTRR for " & maxRows2 & " " & IncidentType & " incidents for the previous three months is: " & MTRIRIncidents & " hours")

Else

MsgBox ("No incidents to report for this metric")

End If
Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1

End Sub

Sub PopulateDropDowns()

  Sheets("Dashboard").DropDowns("DropDownBox1").RemoveAllItems
  Sheets("Dashboard").DropDowns("DropDownBox3").RemoveAllItems
  Sheets("Dashboard").DropDowns("DropDownBox5").RemoveAllItems
  
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Please Select", Index:=1
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Audit", Index:=2
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Customer Dispute", Index:=3
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Data Subject Access Request (DSAR)", Index:=4
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Date Loss Event", Index:=5
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - E-Disclosure", Index:=6
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Equipment Theft/Loss", Index:=7
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Human Resources Investigation", Index:=8
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Policies, Procedure or Process", Index:=9
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Internal Usage Investigation", Index:=10
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Regulatory Investigation", Index:=11
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Event - Security Assistance", Index:=12
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Data Exposure", Index:=13
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Fraudulent Activity", Index:=14
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Malware", Index:=15
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Policies, Procedure or Process ", Index:=16
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Service Outage", Index:=17
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Unauthorised Access to Information", Index:=18
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="Incident - Unauthorised Modification of Information", Index:=19
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="All Events", Index:=20
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="All Incidents", Index:=21
  Sheets("Dashboard").DropDowns("DropDownBox1").AddItem Text:="All Occurrences", Index:=22
  
  Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1
    
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Please Select", Index:=1
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Audit", Index:=2
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Customer Dispute", Index:=3
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Data Subject Access Request (DSAR)", Index:=4
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Date Loss Event", Index:=5
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - E-Disclosure", Index:=6
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Equipment Theft/Loss", Index:=7
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Human Resources Investigation", Index:=8
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Policies, Procedure or Process", Index:=9
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Internal Usage Investigation", Index:=10
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Regulatory Investigation", Index:=11
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Event - Security Assistance", Index:=12
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Data Exposure", Index:=13
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Fraudulent Activity", Index:=14
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Malware", Index:=15
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Policies, Procedure or Process ", Index:=16
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Service Outage", Index:=17
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Unauthorised Access to Information", Index:=18
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="Incident - Unauthorised Modification of Information", Index:=19
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="All Events", Index:=20
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="All Incidents", Index:=21
  Sheets("Dashboard").DropDowns("DropDownBox3").AddItem Text:="All Occurrences", Index:=22

  Sheets("Dashboard").DropDowns("DropDownBox3").ListIndex = 1
  
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Please Select", Index:=1
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Audit", Index:=2
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Customer Dispute", Index:=3
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Data Subject Access Request (DSAR)", Index:=4
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Date Loss Event", Index:=5
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - E-Disclosure", Index:=6
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Equipment Theft/Loss", Index:=7
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Human Resources Investigation", Index:=8
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Policies, Procedure or Process", Index:=9
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Internal Usage Investigation", Index:=10
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Regulatory Investigation", Index:=11
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Event - Security Assistance", Index:=12
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Data Exposure", Index:=13
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Fraudulent Activity", Index:=14
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Malware", Index:=15
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Policies, Procedure or Process ", Index:=16
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Service Outage", Index:=17
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Unauthorised Access to Information", Index:=18
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="Incident - Unauthorised Modification of Information", Index:=19
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="All Events", Index:=20
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="All Incidents", Index:=21
  Sheets("Dashboard").DropDowns("DropDownBox5").AddItem Text:="All Occurrences", Index:=22
  
  Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1
  
  Sheets("Dashboard").DropDowns("DropDownBox7").AddItem Text:="Select Desired Action", Index:=1
  Sheets("Dashboard").DropDowns("DropDownBox7").AddItem Text:="Show Open Records", Index:=2
  Sheets("Dashboard").DropDowns("DropDownBox7").AddItem Text:="Show Incomplete Records", Index:=3
    
  Sheets("Dashboard").DropDowns("DropDownBox7").ListIndex = 1
    
End Sub

Sub OpenMTIR(IncidentType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTIR_Real_Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTIR_Real_Data"
Worksheets("MTIR_Real_Data").Visible = False

Application.DisplayAlerts = False
Dim ws2 As Worksheet
For Each ws2 In Worksheets
If ws2.Name = "MTIR Data" Then ws2.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTIR Data"
Worksheets("MTIR Data").Visible = False

Date1 = Date

DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")

Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Sheets("MTIR Data").Cells(1, 1).Value = "Subject"
Sheets("MTIR Data").Cells(1, 2).Value = "Category"
Sheets("MTIR Data").Cells(1, 3).Value = "Date Reported"
Sheets("MTIR Data").Cells(1, 4).Value = "Time Reported"
Sheets("MTIR Data").Cells(1, 5).Value = "Date Closed"
Sheets("MTIR Data").Cells(1, 6).Value = "Time Closed"

Sheets("MTIR Data").Range("A1:G1").HorizontalAlignment = xlCenter
Sheets("MTIR Data").Range("A1:G1").Font.Bold = True
Sheets("MTIR Data").UsedRange.Columns.AutoFit

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

         For Each cell In r

             IncidentCheckString = "B" & Trim(Str(counter))
             IncidentCheck = Sheets("CompleteRecords").Range(IncidentCheckString).Value

             LowerDateDiff = DateDiff("d", Date2, cell.Value)
             UpperDateDiff = DateDiff("d", Date1, cell.Value)

             If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = IncidentType Then

             CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":D" & Trim(Str(rawdatanumber))
             CopyRangeTo = "A" & Trim(Str(mynumber)) & ":D" & Trim(Str(mynumber))

             Sheets("MTIR Data").Range(CopyRangeTo).NumberFormat = "@"
             Sheets("MTIR Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

             ClosedStringFrom = "I" & Trim(Str(rawdatanumber)) & ":J" & Trim(Str(rawdatanumber))
             ClosedStringTo = "E" & Trim(Str(mynumber)) & ":F" & Trim(Str(mynumber))
             Sheets("MTIR Data").Range(ClosedStringTo).NumberFormat = "@"
             Sheets("MTIR Data").Range(ClosedStringTo).Value = Sheets("CompleteRecords").Range(ClosedStringFrom).Value

             DateRepRange = "C" & Trim(Str(mynumber))
             DateReported = Sheets("MTIR Data").Range(DateRepRange)

             TimeRepRange = "D" & Trim(Str(mynumber))
             TimeReported = Sheets("MTIR Data").Range(TimeRepRange)

             DateTimeReported = DateReported & " " & TimeReported

             DateRepRange = "A" & Trim(Str(mynumber))
             Sheets("MTIR_Real_Data").Range(DateRepRange).NumberFormat = "@"
             Sheets("MTIR_Real_Data").Range(DateRepRange) = DateTimeReported

             DateCloRange = "E" & Trim(Str(mynumber))
             DateClosed = Sheets("MTIR Data").Range(DateCloRange)
             TimeCloRange = "F" & Trim(Str(mynumber))
             TimeClosed = Sheets("MTIR Data").Range(TimeCloRange)
             
             TimeDateClosed = DateClosed & " " & TimeClosed
             
             DateCloRange = "B" & Trim(Str(mynumber))
             Sheets("MTIR_Real_Data").Range(DateCloRange).NumberFormat = "@"
             Sheets("MTIR_Real_Data").Range(DateCloRange) = TimeDateClosed
             
             MTIRRange = "D" & Trim(Str(mynumber))
             Sheets("MTIR_Real_Data").Range(MTIRRange).NumberFormat = "0.0"
                    
             Sheets("MTIR_Real_Data").Range(MTIRRange).Formula = "=9*(NETWORKDAYS(RC[-3],RC[-2])-1)-24*((MOD(RC[-3],1)-MOD(RC[-2],1)))"
             
             mynumber = mynumber + 1
             End If

             rawdatanumber = rawdatanumber + 1
             counter = counter + 1
         Next

Sheets("MTIR Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTIR_Real_Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows2 >= 2 Then

NumberofIncidents = 0
TotalTime = 0#

For i = 2 To maxRows2

RangeString = "D" & Trim(Str(i))
Temp = Sheets("MTIR_Real_Data").Range(RangeString).Value
TotalTime = TotalTime + Temp
NumberofIncidents = NumberofIncidents + 1

Next

MTIRMetric = TotalTime / NumberofIncidents
MTIRMetric = Math.Round(MTIRMetric, 2)

IncidentType = Mid(IncidentType, 12)

MsgBox ("The MTR for " & NumberofIncidents & " " & IncidentType & " incidents for the previous three months is: " & MTIRMetric & " hours")

Else

MsgBox ("No incidents to report for this metric")

End If
Sheets("Dashboard").DropDowns("DropDownBox1").ListIndex = 1
End Sub

Sub OpenMTWE(EventType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTWE Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTWE Data"
Worksheets("MTWE Data").Visible = False

Sheets("MTWE Data").Cells(1, 1).Value = "Subject"
Sheets("MTWE Data").Cells(1, 2).Value = "Category"
Sheets("MTWE Data").Cells(1, 3).Value = "Working Hours"
Sheets("MTWE Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("MTWE Data").Range("A1:C1").Font.Bold = True
Sheets("MTWE Data").UsedRange.Columns.AutoFit

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

         For Each cell In r

           EventCheckString = "B" & Trim(Str(counter))
           EventCheck = Sheets("CompleteRecords").Range(EventCheckString).Value

           LowerDateDiff = DateDiff("d", Date2, cell.Value)
           UpperDateDiff = DateDiff("d", Date1, cell.Value)

           If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = EventType Then

             CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":B" & Trim(Str(rawdatanumber))
             CopyRangeTo = "A" & Trim(Str(mynumber)) & ":B" & Trim(Str(mynumber))
             Sheets("MTWE Data").Range(CopyRangeTo).NumberFormat = "@"
             Sheets("MTWE Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

             WorkingStringFrom = "K" & Trim(Str(rawdatanumber))
             WorkingStringTo = "C" & Trim(Str(mynumber))

             Sheets("MTWE Data").Range(WorkingStringTo).NumberFormat = Number
             Sheets("MTWE Data").Range(WorkingStringTo).Value = Sheets("CompleteRecords").Range(WorkingStringFrom).Value

             mynumber = mynumber + 1
           End If

           rawdatanumber = rawdatanumber + 1
           counter = counter + 1
         Next

Sheets("MTWE Data").Columns("C").Cells.HorizontalAlignment = xlCenter
Sheets("MTWE Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTWE Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows >= 2 Then

MTWEHoursCount = 0
valid = 0

 For i = 2 To maxRows2
         RangeString = "C" & Trim(Str(i))
         Hours = Sheets("MTWE Data").Range(RangeString).Value
         If Hours <> "Unknown" Then
         MTWEHoursCount = MTWEHoursCount + Hours
         valid = valid + 1
         End If
 Next

If MTWEHoursCount <> 0 Then

 MTWE = MTWEHoursCount / valid
 MTWE = Math.Round(MTWE, 2)
 
 MeanCostofEvents = MTWEHoursCount * 53.75 / valid
 MeanCostofEvents = Math.Round(MeanCostofEvents, 2)

 EventType = Mid(EventType, 8)

MsgBox ("The MTWO for " & valid & " " & EventType & " events for the previous three months is: " & MTWE & " hours. The average cost of these events is: £" & MeanCostofEvents)

Else

MsgBox ("No events to report for this metric")

End If
End If

Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1
End Sub

Sub MTWEAllEvents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTWE_All Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTWE_All Data"
Worksheets("MTWE_All Data").Visible = False

Sheets("MTWE_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTWE_All Data").Cells(1, 2).Value = "Category"
Sheets("MTWE_All Data").Cells(1, 3).Value = "Working Hours"
Sheets("MTWE_All Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("MTWE_All Data").Range("A1:C1").Font.Bold = True
Sheets("MTWE_All Data").UsedRange.Columns.AutoFit

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

         For Each cell In r

            EventCheckPos = "B" & Trim(Str(counter))
            EventCheckString = Sheets("CompleteRecords").Range(EventCheckPos).Value
            EventCheck = InStr(1, EventCheckString, "Event", vbTextCompare)
               
            LowerDateDiff = DateDiff("d", Date2, cell.Value)
            UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
            If LowerDateDiff >= 0 And UpperDateDiff <= 0 And EventCheck = 1 Then

             CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":B" & Trim(Str(rawdatanumber))
             CopyRangeTo = "A" & Trim(Str(mynumber)) & ":B" & Trim(Str(mynumber))
             Sheets("MTWE_All Data").Range(CopyRangeTo).NumberFormat = "@"
             Sheets("MTWE_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

             WorkingStringFrom = "K" & Trim(Str(rawdatanumber))
             WorkingStringTo = "C" & Trim(Str(mynumber))

             Sheets("MTWE_All Data").Range(WorkingStringTo).NumberFormat = Number
             Sheets("MTWE_All Data").Range(WorkingStringTo).Value = Sheets("CompleteRecords").Range(WorkingStringFrom).Value

             mynumber = mynumber + 1
           End If

           rawdatanumber = rawdatanumber + 1
           counter = counter + 1
         Next

Sheets("MTWE_All Data").Columns("C").Cells.HorizontalAlignment = xlCenter
Sheets("MTWE_All Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTWE_All Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows >= 2 Then

MTWEHoursCount = 0
valid = 0

 For i = 2 To maxRows2
         RangeString = "C" & Trim(Str(i))
         Hours = Sheets("MTWE_All Data").Range(RangeString).Value
         If Hours <> "Unknown" Then
         MTWEHoursCount = MTWEHoursCount + Hours
         valid = valid + 1
         End If
 Next

If MTWEHoursCount <> 0 Then

 MTWE = MTWEHoursCount / valid
 MTWE = Math.Round(MTWE, 2)
 
 MeanCostofEvents = MTWEHoursCount * 53.75 / valid
 MeanCostofEvents = Math.Round(MeanCostofEvents, 2)

 EventType = Mid(EventType, 8)

MsgBox ("The MTWO for " & valid & " " & EventType & " events for the previous three months is: " & MTWE & " hours. The average cost of these events is: £" & MeanCostofEvents)

Else

MsgBox ("No events to report for this metric")

End If
End If

Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1
End Sub

Sub MTWEAll()

Dim r As Range, cell As Range

Application.DisplayAlerts = False

Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTWE_All_2 Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTWE_All_2 Data"
Worksheets("MTWE_All_2 Data").Visible = False

Sheets("MTWE_All_2 Data").Cells(1, 1).Value = "Subject"
Sheets("MTWE_All_2 Data").Cells(1, 2).Value = "Category"
Sheets("MTWE_All_2 Data").Cells(1, 3).Value = "Working Hours"
Sheets("MTWE_All_2 Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("MTWE_All_2 Data").Range("A1:C1").Font.Bold = True
Sheets("MTWE_All_2 Data").UsedRange.Columns.AutoFit

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

         For Each cell In r
           
            LowerDateDiff = DateDiff("d", Date2, cell.Value)
            UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
            If LowerDateDiff >= 0 And UpperDateDiff <= 0 Then

             CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":B" & Trim(Str(rawdatanumber))
             CopyRangeTo = "A" & Trim(Str(mynumber)) & ":B" & Trim(Str(mynumber))
             Sheets("MTWE_All_2 Data").Range(CopyRangeTo).NumberFormat = "@"
             Sheets("MTWE_All_2 Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

             WorkingStringFrom = "K" & Trim(Str(rawdatanumber))
             WorkingStringTo = "C" & Trim(Str(mynumber))

             Sheets("MTWE_All_2 Data").Range(WorkingStringTo).NumberFormat = Number
             Sheets("MTWE_All_2 Data").Range(WorkingStringTo).Value = Sheets("CompleteRecords").Range(WorkingStringFrom).Value

             mynumber = mynumber + 1
           End If

           rawdatanumber = rawdatanumber + 1
           counter = counter + 1
         Next

Sheets("MTWE_All_2 Data").Columns("C").Cells.HorizontalAlignment = xlCenter
Sheets("MTWE_All_2 Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTWE_All_2 Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows >= 2 Then

MTWEHoursCount = 0
valid = 0

 For i = 2 To maxRows2
         RangeString = "C" & Trim(Str(i))
         Hours = Sheets("MTWE_All_2 Data").Range(RangeString).Value
         If Hours <> "Unknown" Then
         MTWEHoursCount = MTWEHoursCount + Hours
         valid = valid + 1
         End If
 Next

If MTWEHoursCount <> 0 Then

 MTWE = MTWEHoursCount / valid
 MTWE = Math.Round(MTWE, 2)
 
 MeanCostofEvents = MTWEHoursCount * 53.75 / valid
 MeanCostofEvents = Math.Round(MeanCostofEvents, 2)

 EventType = Mid(EventType, 8)

MsgBox ("The MTWO for " & valid & " events and incidents for the previous three months is: " & MTWE & " hours. The average cost of these events is: £" & MeanCostofEvents)

Else

MsgBox ("No occurrences to report for this metric")

End If
End If

Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1

End Sub

Sub MTWEAllIncidents()

Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTWI_All Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTWI_All Data"
Worksheets("MTWI_All Data").Visible = False

Sheets("MTWI_All Data").Cells(1, 1).Value = "Subject"
Sheets("MTWI_All Data").Cells(1, 2).Value = "Category"
Sheets("MTWI_All Data").Cells(1, 3).Value = "Working Hours"
Sheets("MTWI_All Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("MTWI_All Data").Range("A1:C1").Font.Bold = True
Sheets("MTWI_All Data").UsedRange.Columns.AutoFit

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

        For Each cell In r

           IncidentCheckPos = "B" & Trim(Str(counter))
           IncidentCheckString = Sheets("CompleteRecords").Range(IncidentCheckPos).Value
           IncidentCheck = InStr(1, IncidentCheckString, "Incident", vbTextCompare)
               
           LowerDateDiff = DateDiff("d", Date2, cell.Value)
           UpperDateDiff = DateDiff("d", Date1, cell.Value)
               
           If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = 1 Then
           
                CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":B" & Trim(Str(rawdatanumber))
                CopyRangeTo = "A" & Trim(Str(mynumber)) & ":B" & Trim(Str(mynumber))
                Sheets("MTWI_All Data").Range(CopyRangeTo).NumberFormat = "@"
                Sheets("MTWI_All Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

                WorkingStringFrom = "K" & Trim(Str(rawdatanumber))
                WorkingStringTo = "C" & Trim(Str(mynumber))
                Sheets("MTWI_All Data").Range(WorkingStringTo).NumberFormat = Number
                Sheets("MTWI_All Data").Range(WorkingStringTo).Value = Sheets("CompleteRecords").Range(WorkingStringFrom).Value

               mynumber = mynumber + 1
              End If

              rawdatanumber = rawdatanumber + 1
              counter = counter + 1
         Next

Sheets("MTWI_All Data").Columns("C").Cells.HorizontalAlignment = xlCenter
Sheets("MTWI_All Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTWI_All Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows >= 2 Then

MTWIHoursCount = 0
valid = 0

   For i = 2 To maxRows2
         RangeString = "C" & Trim(Str(i))
         Hours = Sheets("MTWI_All Data").Range(RangeString).Value
         If Hours <> "Unknown" Then
         MeanTimeIncidents = MeanTimeIncidents + Hours
         valid = valid + 1
         End If
  Next

 If MeanTimeIncidents <> 0 Then

    MTWI = MeanTimeIncidents / valid
    MTWI = Math.Round(MTWI, 2)
    MeanCostofIncidents = MeanTimeIncidents * 53.75 / valid
    MeanCostofIncidents = Math.Round(MeanCostofIncidents, 2)

    IncidentType = Mid(IncidentType, 12)

MsgBox ("The MTWO for " & valid & " " & IncidentType & " incidents for the previous three months is: " & MTWI & " hours. The average cost of these incidents is: £" & MeanCostofIncidents)

Else

MsgBox ("No incidents to report for this metric")

End If
End If
Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1
End Sub
Sub OpenMTWI(IncidentType As String)

Dim r As Range, cell As Range

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In Worksheets
If ws.Name = "MTWI Data" Then ws.Delete
Next
Application.DisplayAlerts = True

Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MTWI Data"
Worksheets("MTWI Data").Visible = False

Sheets("MTWI Data").Cells(1, 1).Value = "Subject"
Sheets("MTWI Data").Cells(1, 2).Value = "Category"
Sheets("MTWI Data").Cells(1, 3).Value = "Working Hours"
Sheets("MTWI Data").Range("A1:C1").HorizontalAlignment = xlCenter
Sheets("MTWI Data").Range("A1:C1").Font.Bold = True
Sheets("MTWI Data").UsedRange.Columns.AutoFit

Date1 = Date
DateTemp = DateAdd("m", -3, Date1)
DayValue = Format(DateTemp, "dd")
Date2 = DateAdd("d", -DayValue + 1, DateTemp)

Set ws2 = Sheets("CompleteRecords")
  With ws2
    maxRows = .Range("C" & .Rows.count).End(xlUp).Row
  End With

RangeString = "C2:C" & Trim(Str(maxRows))
Set r = Sheets("CompleteRecords").Range(RangeString)

mynumber = 2
rawdatanumber = 2
counter = 2

        For Each cell In r

           IncidentCheckString = "B" & Trim(Str(counter))
           IncidentCheck = Sheets("CompleteRecords").Range(IncidentCheckString).Value

           LowerDateDiff = DateDiff("d", Date2, cell.Value)
           UpperDateDiff = DateDiff("d", Date1, cell.Value)

              If LowerDateDiff >= 0 And UpperDateDiff <= 0 And IncidentCheck = IncidentType Then

                CopyRangeFrom = "A" & Trim(Str(rawdatanumber)) & ":B" & Trim(Str(rawdatanumber))
                CopyRangeTo = "A" & Trim(Str(mynumber)) & ":B" & Trim(Str(mynumber))
                Sheets("MTWI Data").Range(CopyRangeTo).NumberFormat = "@"
                Sheets("MTWI Data").Range(CopyRangeTo).Value = Sheets("CompleteRecords").Range(CopyRangeFrom).Value

                WorkingStringFrom = "K" & Trim(Str(rawdatanumber))
                WorkingStringTo = "C" & Trim(Str(mynumber))
                Sheets("MTWI Data").Range(WorkingStringTo).NumberFormat = Number
                Sheets("MTWI Data").Range(WorkingStringTo).Value = Sheets("CompleteRecords").Range(WorkingStringFrom).Value

               mynumber = mynumber + 1
              End If

              rawdatanumber = rawdatanumber + 1
              counter = counter + 1
         Next

Sheets("MTWI Data").Columns("C").Cells.HorizontalAlignment = xlCenter
Sheets("MTWI Data").UsedRange.Columns.AutoFit

Set ws3 = Sheets("MTWI Data")
  With ws3
    maxRows2 = .Range("A" & .Rows.count).End(xlUp).Row
  End With

If maxRows >= 2 Then

MTWIHoursCount = 0
valid = 0

   For i = 2 To maxRows2
         RangeString = "C" & Trim(Str(i))
         Hours = Sheets("MTWI Data").Range(RangeString).Value
         If Hours <> "Unknown" Then
         MeanTimeIncidents = MeanTimeIncidents + Hours
         valid = valid + 1
         End If
  Next

 If MeanTimeIncidents <> 0 Then

    MTWI = MeanTimeIncidents / valid
    MTWI = Math.Round(MTWI, 2)
    MeanCostofIncidents = MeanTimeIncidents * 53.75 / valid
    MeanCostofIncidents = Math.Round(MeanCostofIncidents, 2)

    IncidentType = Mid(IncidentType, 12)

MsgBox ("The MTWO for " & valid & " " & IncidentType & " incidents for the previous three months is: " & MTWI & " hours. The average cost of these incidents is: £" & MeanCostofIncidents)

Else

MsgBox ("No incidents to report for this metric")

End If
End If
Sheets("Dashboard").DropDowns("DropDownBox5").ListIndex = 1

End Sub

Sub BuildDashboard()

Application.ScreenUpdating = False
Application.StatusBar = "Building Dashboard..."

Call CreateDashboard
Call Collect_Attack_Type_Data

Application.StatusBar = "Collecting Data..."

Call BreakDownEvents
Call BreakDownIncidents
Call IncidentEvents_Chart
Call Attack_Incidents_Chart

Application.StatusBar = "Creating Charts..."

Call MostRecentChart
Call DataForPastTwelveMonths

Application.StatusBar = "Collecting Data..."

Call CreateAverageCharts
Call CollectCompleteRecords

Application.StatusBar = "Building Dashboard..."

Call LoadNewTrendData
Call Open_Closed_Trend_Chart

Application.ScreenUpdating = True
Worksheets("Dashboard").Visible = True
Sheets("Dashboard").Range("A1").Select

Application.StatusBar = "Ready"

End Sub

Sub CallRecordStatus()

If dropdownindex <> Sheets("Dashboard").DropDowns("DropDownBox7").ListIndex Then

  dropdownindex = Sheets("Dashboard").DropDowns("DropDownBox7").ListIndex

 Select Case dropdownindex

    Case 2
    Worksheets("Open Records").Visible = True
    Worksheets("Open Records").Activate

    Case 3
    Worksheets("Incomplete Records").Visible = True
    Worksheets("Incomplete Records").Activate

   End Select

End If
Sheets("Dashboard").DropDowns("DropDownBox7").ListIndex = 1
End Sub

Sub CloseIncompleteRecords()

Worksheets("Incomplete Records").Visible = False

Worksheets("Dashboard").Visible = True
Worksheets("Dashboard").Activate

End Sub

Sub CloseOpenRecords()

Worksheets("Open Records").Visible = False

Worksheets("Dashboard").Visible = True
Worksheets("Dashboard").Activate

End Sub

Sub SummaryCollect(SummaryReportedBody)

 SummaryCheck1 = InStr(1, SummaryReportedBody, "Record Brief:", vbTextCompare)
 SummaryCheck2 = InStr(1, SummaryReportedBody, "Appendix for communications", vbTextCompare)
 SummaryDiff = SummaryCheck2 - SummaryCheck1
 
 If SummaryCheck1 <> 0 Then
  
 SummaryInfo = Mid(SummaryReportedBody, SummaryCheck1 + 15, SummaryDiff - 19)
 
 Else
 
 SummaryInfo = "Unknown"
 
 End If
 
 RangeString = "M" & Trim(Str(RowCount))
 Sheets("Raw Data").Range(RangeString).NumberFormat = "@"
 Sheets("Raw Data").Range(RangeString).Value = SummaryInfo
 
 Sheets("Raw Data").Columns("M").HorizontalAlignment = xlLeft
 Sheets("Raw Data").Columns("M").VerticalAlignment = xlBottom
 Sheets("Raw Data").Columns("M").WrapText = False
 Sheets("Raw Data").Columns("M").Orientation = 0
 Sheets("Raw Data").Columns("M").AddIndent = False
 Sheets("Raw Data").Columns("M").IndentLevel = 0
 Sheets("Raw Data").Columns("M").ShrinkToFit = False
 Sheets("Raw Data").Columns("M").ReadingOrder = xlContext
 
 Columns("M:M").EntireColumn.AutoFit
 
End Sub
