VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArrDateForm 
   Caption         =   "Insert Date"
   ClientHeight    =   3105
   ClientLeft      =   30
   ClientTop       =   480
   ClientWidth     =   4710
   OleObjectBlob   =   "ArrDateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArrDateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'InsertArrDateForm module
 Option Explicit
 Private Const MAX_UNDO As Long = 20
 Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
 
'Apr 2011 - Inserts dates from year 1000 to 3000 in active cell.
'Oct 2011 - Rewritten - MAX_UNDO determines the limit for the number of Undo's (in 3 subs).
'Nov 2011 - CountA fails on Indexed column in array in XL 2010 - replaced with loop.
'Dec 2011 - Changed to array from worksheet to hold dates.
'Feb 2013 - Revised format in 1 & 12 month calendars and added data overwrite check.
'James Cone - Portland, Oregon USA - Copyrighted - xxjamesconexx@gmail.com

Private Sub cmdButtonExit_Click()
 On Error Resume Next
 Unload Me
 DoEvents
End Sub

Private Sub cmdButtonInfo_Click()
On Error GoTo BadInfo
Dim M          As Long
Dim Y          As Long
Dim strAddress As String

'  If you set the value or formula of a cell to a date, Excel checks to see whether
'  that cell is already formatted with one of the date or time number formats.
'  If not, Excel changes the number format to the default short date number format.

'INFO MESSAGE
 On Error GoTo BadInfo
 If cmdButtonInfo.Caption = "Info" Then
   Call ShowInsertInfo
'UNDO
 Else
   On Error Resume Next
   For Y = LBound(vFormulas, 1) To UBound(vFormulas, 1)
       If Len(vFormulas(Y, 3)) > 0 Then M = M + 1
   Next
  'Is full address
   strAddress = vFormulas(M, 3)
   On Error Resume Next
   ActiveSheet.Range(strAddress).Parent.Activate
  'Workbook closed or sheet deleted.
   If Err.Number <> 0 Then
     ReDim vFormulas(1 To MAX_UNDO, 1 To 3)
     cmdButtonInfo.Caption = "Info"
     cmdButtonInfo.ForeColor = vbButtonText
     Err.Clear
     GoTo BadInfo
   End If
   On Error GoTo BadInfo
   ActiveSheet.Range(strAddress).Value2 = vFormulas(M, 1)
   ActiveSheet.Range(strAddress).NumberFormat = vFormulas(M, 2)
   vFormulas(M, 1) = Empty: vFormulas(M, 2) = Empty: vFormulas(M, 3) = Empty
   If M = 1 Then
     cmdButtonInfo.Caption = "Info"
     cmdButtonInfo.ForeColor = vbButtonText
     cmdButtonInfo.ControlTipText = vbNullString
     cmdButtonInsert.ControlTipText = "to append date: press shift key when inserting"
  'Insert makes it red at MAX_UNDO
   ElseIf M = MAX_UNDO - 1 Then
     cmdButtonInfo.ForeColor = vbButtonText
   End If
 End If
 Me.Frame1.SetFocus
 Exit Sub
BadInfo:
 Application.Cursor = xlDefault
 MsgBox "Unable to undo.     ", vbExclamation, "Insert Date"
 Me.Frame1.SetFocus
End Sub

Private Sub CmdButtonInsert_Click()
'ADD/APPEND DATE TO WORKSHEET.
 On Error GoTo DoesNotFit
 Dim strAddress As String
 Dim strFormat  As String
 Dim blnExists  As Boolean
 Dim objLB      As MSForms.ListBox
 Dim dteValue   As Variant 'date
 Dim d          As Long    'day
 Dim i          As Long
 Dim M          As Long
 Dim Y          As Long    'year
 
 For i = 1 To Me.Frame1.Controls.Count
    If Me.Controls("ListBox" & i).ListIndex > -1 Then
      Set objLB = Me.Controls("ListBox" & i)
      On Error Resume Next
     'Returns a string
      d = CLng(objLB.Value)
      On Error GoTo DoesNotFit
      Exit For
    End If
 Next 'i
 If Not d > 0 Then Err.Raise 56789, , "Date is not valid - Unable to insert"
 DoEvents
   
'SAVE INFO IN vFORMULAS
 strAddress = ActiveCell.Address(True, True, xlA1, True, Nothing)
 i = 0
 On Error Resume Next
 With Application.WorksheetFunction
   i = .Match(strAddress, .Index(vFormulas, 0, 3), 0)
 End With
 On Error GoTo DoesNotFit
'Only if a new cell.
 If i <> 0 And i <= MAX_UNDO Then
  'Cell address is in vFormulas
   M = i
   blnExists = True
 Else
   M = 0
   For Y = LBound(vFormulas, 1) To UBound(vFormulas, 1)
       If Len(vFormulas(Y, 3)) > 0 Then M = M + 1
   Next
   If M >= MAX_UNDO Then
     M = MAX_UNDO
    'Shuffle all array values up one row
     For Y = 1 To (MAX_UNDO - 1)
       For i = 1 To 3
           vFormulas(Y, i) = vFormulas(Y + 1, i)
       Next
     Next
     vFormulas(MAX_UNDO, 1) = Empty
     vFormulas(MAX_UNDO, 2) = Empty
     vFormulas(MAX_UNDO, 3) = Empty
   Else
     M = M + 1
   End If
 End If
 
'Only save data from new cell locations.
 If Not blnExists Then
   vFormulas(M, 1) = ActiveCell.Formula
   vFormulas(M, 2) = ActiveCell.NumberFormat
   vFormulas(M, 3) = strAddress
   If M = 1 Then
     cmdButtonInfo.ForeColor = vbBlue
     cmdButtonInfo.Caption = "Undo"
     If Val(Application.Version) >= 9 Then _
        cmdButtonInfo.ControlTipText = "undo limited to the last " & MAX_UNDO & " inserts"
     cmdButtonInsert.ControlTipText = "to append date: press shift key when inserting"
   ElseIf M = 3 Then
     cmdButtonInfo.ControlTipText = vbNullString
     cmdButtonInsert.ControlTipText = vbNullString
   ElseIf M = MAX_UNDO Then
     cmdButtonInfo.ForeColor = vbRed
     If Val(Application.Version) >= 9 Then _
        cmdButtonInfo.ControlTipText = "undo limited to the last " & MAX_UNDO & " inserts"
   Else
     cmdButtonInfo.ForeColor = vbBlue
     If M = (MAX_UNDO - 1) Then cmdButtonInfo.ControlTipText = vbNullString
   End If
 End If
 
'Changing cell dependents creates error values in cells with formulas
' so convert cell to value.
 On Error Resume Next
 ActiveCell.Value2 = ActiveCell.Value2
 If Err.Number <> 0 Then                     'belts and suspenders
   On Error GoTo DoesNotFit
   ActiveCell.Copy
   ActiveCell.PasteSpecial xlPasteValues
   Application.CutCopyMode = False
 Else
   On Error GoTo DoesNotFit
 End If

'INSERT DATE IN CELLL
'Determine date - month is spelled out.
 i = Me.sbMonth.Value
 Y = Me.sbYear.Value
 If objLB.ListIndex = 0 And objLB.Value > 7 Then
   i = i - 1
   If i = 0 Then
     i = 12
     Y = Y - 1
   End If
 End If
'DateSerial allows for international formats - DateValue does not.
 dteValue = VBA.DateSerial(Y, i, d)          'Hans Vogelaar
 If GetKeyState(vbKeyShift) < 0 Then         'APPENDING
    If Not IsEmpty(ActiveCell) Then
     'Using Str function will not add leading space.
      ActiveCell.Value = ActiveCell.Value & " " & dteValue
    Else 'blank cell
      ActiveCell.Value = dteValue
    End If
 Else                                        'INSERTING
    ActiveCell.Value = dteValue
 End If
 Me.Caption = VBA.UCase$(Format$(dteValue, "yyyy - mmmm ")) & d
 Me.Frame1.SetFocus
 Set objLB = Nothing
 Exit Sub
DoesNotFit:
 Application.Cursor = xlDefault
 MsgBox Err.Description & ".   ", vbExclamation, "Insert Date"
End Sub

Private Sub CmdButtonReset_Click()
'Resets date to current date or adds calendar to worksheet.
'Feb 14, 2013 - Reduced font size of prior month dates in first row of calendar.
'               Clearing entire data area before pasting a calendar.
 On Error GoTo Voided
 Dim FirstDay As Long
 Dim Col      As Long
 Dim Rw       As Long
 Dim M        As Long
 Dim Y        As Long
 Dim Awf      As Excel.WorksheetFunction
 Dim rngDates As Excel.Range
 
'COPY MONTH
 If GetKeyState(vbKeyShift) < 0 Then
 If ActiveCell.Row > Rows.Count - 7 Or ActiveCell.Column > Columns.Count - 6 Then
   MsgBox "Need a little more room." & vbCr & _
          "Select another cell further from the edge of the worksheet.   ", _
           vbInformation, "Add Calendar Month"
   GoTo Rehabilitation
 Else
   Set Awf = Application.WorksheetFunction
   If Awf.CountA(ActiveCell.Resize(8, 7)) > 0 Then
     Beep
     If MsgBox("Overwrite existing data ?" & vbCr & "(undo cannot restore data)     ", _
       vbYesNo + vbQuestion, "Add Calendar Month") <> vbYes Then
       GoTo Rehabilitation
     End If
   End If
   Application.ScreenUpdating = False
   ActiveCell.Resize(8, 7).Clear
   M = Me.sbMonth.Value
   Y = Me.sbYear.Value
   ActiveCell.Value2 = Awf.Proper(Format$(M & "/28/" & Y, "yyyy - mmmm"))
   ActiveCell.Resize(1, 7).HorizontalAlignment = xlHAlignCenterAcrossSelection
   ActiveCell.Offset(1, 0).Resize(1, 7).Value2 = GetDayNames
   ActiveCell.Offset(1, 0).Resize(1, 7).HorizontalAlignment = xlHAlignCenter
  'BorderAround method in xl2010 is not reliable.
   With ActiveCell.Offset(2, 0).Resize(6, 7)
     .Value2 = vArrDates
     .NumberFormat = "General_)"
     .Interior.Color = vbWhite
     .Borders(xlEdgeTop).Weight = xlHairline
     .Borders(xlEdgeLeft).Weight = xlHairline
     .Borders(xlEdgeRight).Weight = xlHairline
     .Borders(xlEdgeBottom).Weight = xlHairline
      On Error Resume Next
      FirstDay = Awf.Match(1, .Rows(1).Cells, 0)
      On Error GoTo Voided
      If FirstDay > 1 Then
       With .Cells(1, 1).Resize(1, FirstDay - 1).Font
         .Color = vbBlue
         .Size = .Size - 1.5
       End With
      End If     'Only make columns narrower, never wider.
      If .Columns(1).ColumnWidth >= ActiveSheet.StandardWidth * 0.41 Then     '3.3
         .EntireColumn.ColumnWidth = Awf.Max(ActiveSheet.StandardWidth * 0.41, 3.3)
      Else
         .EntireColumn.ColumnWidth = .Columns(1).ColumnWidth
      End If
   End With
   Call cmdButtonExit_Click
   Application.ScreenUpdating = True
   End If
   
'COPY TWELVE MONTHS
 ElseIf GetKeyState(vbKeyControl) < 0 Then
 If ActiveCell.Row > Rows.Count - 31 Or ActiveCell.Column > Columns.Count - 22 Then
   MsgBox "Need a little more room." & vbCr & _
          "Select another cell further from the edge of the worksheet.   ", _
           vbInformation, "Add Calendar Year"
   GoTo Rehabilitation
 Else
   Set Awf = Application.WorksheetFunction
   Set rngDates = ActiveCell.Resize(32, 23)
   If Awf.CountA(rngDates) > 0 Then
     Beep
     If MsgBox("Overwrite existing data ?" & vbCr & "(undo cannot restore data)     ", _
       vbYesNo + vbQuestion, "Add Calendar Year") <> vbYes Then
       GoTo Rehabilitation
     End If
   End If
   Application.ScreenUpdating = False
   rngDates.Clear
   M = Me.sbMonth.Value
   Y = Me.sbYear.Value
   For Rw = 1 To 25 Step 8
     For Col = 1 To 17 Step 8
       rngDates(Rw, Col).Value2 = Awf.Proper(Format$(M & "/28/" & Y, "yyyy - mmmm"))
       rngDates(Rw, Col).Resize(1, 7).HorizontalAlignment = xlHAlignCenterAcrossSelection
       rngDates(Rw, Col).Resize(1, 7).Interior.Color = RGB(221, 221, 221) 'light gray pre xl2007
       rngDates(Rw, Col).Offset(1, 0).Resize(1, 7).Value2 = GetDayNames
       rngDates(Rw, Col).Offset(1, 0).Resize(1, 7).HorizontalAlignment = xlHAlignCenter
       With rngDates(Rw, Col).Offset(2, 0).Resize(6, 7)
         .Value2 = vArrDates
         .NumberFormat = "General_)"
         .Interior.Color = vbWhite
         .Borders(xlEdgeTop).Weight = xlHairline
         .Borders(xlEdgeLeft).Weight = xlHairline
         .Borders(xlEdgeRight).Weight = xlHairline
         .Borders(xlEdgeBottom).Weight = xlHairline
          On Error Resume Next
          FirstDay = Awf.Match(1, .Rows(1).Cells, 0)
          On Error GoTo Voided
          If FirstDay > 1 Then
           With .Cells(1, 1).Resize(1, FirstDay - 1).Font
            .Color = vbBlue
            .Size = .Size - 1.5
           End With
          End If
       End With
       M = M + 1
       If M > 12 Then
          M = 1
          Y = Y + 1
       End If
       Call PutDaysInArray(M, Y)
     Next 'Col
     Call PutDaysInArray(M, Y)
   Next 'Rw
   
   With rngDates 'after paste
    .EntireColumn.ColumnWidth = Awf.Max(ActiveSheet.StandardWidth * 0.41, 3.3)
    .Columns(8).EntireColumn.ColumnWidth = Awf.Max(ActiveSheet.StandardWidth / 4, 2)
    .Columns(16).EntireColumn.ColumnWidth = Awf.Max(ActiveSheet.StandardWidth / 4, 2)
   End With
   Call cmdButtonExit_Click
   Application.ScreenUpdating = True
 End If
 
'RESET
 Else
 Me.sbYear.Value = VBA.Year(VBA.Date)
 Me.sbMonth.Value = VBA.Month(VBA.Date)
 ReDim vFormulas(1 To MAX_UNDO, 1 To 3)
 cmdButtonInfo.Caption = "Info"
 cmdButtonInfo.ForeColor = vbButtonText
 cmdButtonInsert.ControlTipText = "to append date: press shift key when inserting"
 cmdButtonInfo.ControlTipText = vbNullString
 Call FindLatestDate(True)
 End If 'GetKeyState(vbKeyShift)
Rehabilitation:
 Set Awf = Nothing
 Set rngDates = Nothing
 Exit Sub
Voided:
 Beep
 Resume Next
End Sub

Private Function FindLatestDate(ByRef bSetToday As Boolean) As Boolean
'Called by Form_Initialize, CmdButtonReset and both Scrollbar controls.
 On Error GoTo BlindDate
 Dim C As Long
 Dim R As Long
 Dim d As Variant         'Day
 Dim blnFound  As Boolean
 Dim objListBx As Control 'Listbox fails
 
 For C = 1 To 7
    Me.Controls("ListBox" & C).Clear
   For R = 2 To 7
    If Len(vArrDates(R, C)) Then
      Me.Controls("ListBox" & C).AddItem vArrDates(R, C)
    End If
   Next
 Next
 
'Only form initialize and reset.
 If bSetToday Then
   d = VBA.Day(VBA.Date)
   For C = 7 To 1 Step -1
     For R = 7 To 2 Step -1
      'Returns an Integer or Empty
       If vArrDates(R, C) = d Then
         If Not blnFound Then
           Set objListBx = Me.Controls("ListBox" & C)
           objListBx.SetFocus
           objListBx.ListIndex = R - 2
           blnFound = True
         End If
         Exit For
       End If
     Next 'C
    'Forces display of entire listbox.
     Me.Controls("ListBox" & C).SetFocus
   Next 'R
   objListBx.SetFocus 'needed
 End If
 Set objListBx = Nothing
 Exit Function
BlindDate:
 Beep
 Me.ListBox1.ListIndex = 0
End Function

Private Sub LabelCell_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
        ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 If ActiveCell.HasFormula Then Me.LabelCell.ControlTipText = "Has Formula" _
    Else Me.LabelCell.ControlTipText = "active cell"
End Sub

Private Sub ListBox1_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox1)
End Sub

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox2_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox2)
End Sub

Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox3_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox3)
End Sub

Private Sub ListBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox4_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox4)
End Sub

Private Sub ListBox4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox5_Click()
 On Error Resume Next
Call ResetListBoxIndexes(ListBox5)
End Sub

Private Sub ListBox5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox6_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox6)
End Sub

Private Sub ListBox6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Sub ListBox7_Click()
 On Error Resume Next
 Call ResetListBoxIndexes(ListBox7)
End Sub

Private Sub ListBox7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
End Sub

Private Function ResetListBoxIndexes(ByRef objLB As Control) As Boolean
'Called when a date is selected.
'Nov 27, 2011 - Added prior month date capability.
 On Error GoTo MakeSound
 Dim oList    As Control
 Dim M        As Long
 Dim Y        As Long
 Dim objValue As Long
 
 M = Me.sbMonth.Value
 Y = Me.sbYear.Value
 On Error Resume Next
 objValue = CLng(objLB.Value)
 On Error GoTo MakeSound
 
 For Each oList In Me.Frame1.Controls
    If Not oList Is objLB Then oList.ListIndex = -1
 Next 'oList
 If objValue > 0 Then
  'Adjust for prior month/year
   If objLB.ListIndex = 0 And objValue > 7 Then
     M = M - 1
     If M = 0 Then
       M = 12
       Y = Y - 1
     End If
   End If
   Me.Caption = VBA.UCase$(Format$(M & "/28/" & Y, "yyyy - mmmm ")) & objValue
 Else
   Me.Caption = VBA.UCase$(Format$(M & "/28/" & Y, "yyyy - mmmm"))
 End If
 Exit Function
MakeSound:
 Beep
 Resume Next
End Function

Private Sub sbmonth_Change()
 On Error GoTo Voided
 Dim M As Long
 Dim Y As Long
 Dim bShftKey   As Boolean
 
 bShftKey = GetKeyState(vbKeyShift) < 0
 M = Me.sbMonth.Value
 Y = Me.sbYear.Value
 If M = 0 Then 'Min is 0, Max is 13
    M = 12
    Me.sbMonth.Value = M
    If bShftKey Then
       Y = Me.sbYear.Value - 1
       Me.sbYear.Value = Y
    End If
 ElseIf M = 13 Then
    M = 1
    Me.sbMonth.Value = M
    If bShftKey Then
       Y = Me.sbYear.Value + 1
       Me.sbYear.Value = Y
    End If
 End If
 Me.Caption = VBA.UCase$(Format$(M & "/28/" & Y, "yyyy - mmmm"))
 If Len(Me.LabelCell.Caption) < 2 Then Exit Sub
 Call PutDaysInArray(M, Y)
 Call FindLatestDate(False)
 Exit Sub
Voided:
 Beep
 Resume Next
End Sub

Private Sub sbyear_Change()
 On Error GoTo Voided
 Dim M As Long
 Dim Y As Long
 
 M = Me.sbMonth.Value
 Y = Me.sbYear.Value
 Me.Caption = VBA.UCase$(Format$(M & "/28/" & Y, "yyyy - mmmm"))
 If Len(Me.LabelCell.Caption) < 2 Then Exit Sub
 Call PutDaysInArray(M, Y)
 Call FindLatestDate(False)
 Exit Sub
Voided:
 Beep
 Resume Next
End Sub

Private Sub UserForm_Activate()
'Moves form left or right if selection is in center of sheet.
 On Error GoTo Sloth
 Dim X             As Long
 Dim lngCount      As Long
 Dim C             As Single
 Dim Num           As Single
 Dim ColumnsInForm As Single
 Dim sngIncrement  As Single
 Dim rngVisible    As Excel.Range
 
'Can't determine actual visible range if window frozen.
 Application.ScreenUpdating = False
 If ActiveWindow.FreezePanes Then
   GoTo Resting
 Else
'If cell selected is near top/bottom - no need to position form.
  Set rngVisible = ActiveWindow.VisibleRange
  X = ActiveCell.Row
  Num = rngVisible.Rows.Count
  lngCount = rngVisible.Rows(Num).Row
  If X < (lngCount \ 5) Or X >= (lngCount * 0.7) Then GoTo Resting
 End If
 
 C = rngVisible.Width / 2
 X = ActiveCell.Column
 lngCount = rngVisible.Columns.Count
'Find middle column (where the form is centered)
 For Num = 1 To lngCount
   With rngVisible.Columns(Num)
     If .Left <= C And (.Left + .Width) > C Then
       C = Num
       Exit For
     End If
  End With
 Next
 If Num > lngCount Then C = lngCount / 2
 If C < 3 Then
   C = 3
 ElseIf C > lngCount - 2 Then
   C = lngCount - 2
 End If
 With rngVisible
   sngIncrement = .Range(.Columns(C - 2!), .Columns(C + 2!)).Width / 5!
 End With
 ColumnsInForm = Application.WorksheetFunction.Ceiling(Me.Width / sngIncrement, 0.5!)
 ColumnsInForm = ColumnsInForm / 2!
 Num = C - X
 If Abs(Num) <= Application.WorksheetFunction.Ceiling(ColumnsInForm, 1) Then
  'Checking for zero (if cell is in center column)
   If Num < 0.1 And Num > -0.1 Then Num = 0.666
  'The farther the cell is from center column, the less the form has to be moved.
   Me.Left = Me.Left + (1 / Num * sngIncrement * ColumnsInForm)
   Me.StartUpPosition = 0  'manual
 End If
Resting:
 Me.Repaint
 Set rngVisible = Nothing
 Application.ScreenUpdating = True
 Exit Sub
Sloth:
 Beep
 Resume Resting
End Sub

Private Sub UserForm_Initialize()
'Feb 2013 - DateValue not req'd in title bar date display.
 On Error GoTo BadForm
 Dim sngResult As Single
 Dim X As Long
 Dim Y As Long
 Dim vDays As Variant
 
 If Not IsArraySet(vFormulas()) Then
   ReDim vFormulas(1 To MAX_UNDO, 1 To 3)
 Else
   cmdButtonInfo.Caption = "Undo"
   cmdButtonInfo.ForeColor = vbBlue
 End If
 vDays = GetDayNames
 Me.LabelSu.Caption = vDays(0)
 Me.LabelMo.Caption = vDays(1)
 Me.LabelTu.Caption = vDays(2)
 Me.LabelWe.Caption = vDays(3)
 Me.LabelTh.Caption = vDays(4)
 Me.LabelFr.Caption = vDays(5)
 Me.LabelSa.Caption = vDays(6)
 sngResult = ResizeToRightSize
 Me.Width = Me.Width * sngResult
 Me.Height = Me.Height * sngResult
 Me.Zoom = sngResult * 100!
'Form, height_margin, width_margin, optional bottom most_Ctrl, optional right most_Ctrl
 Call RefinishTheForm(Me, 3.4!, 7.2!, sngResult, Me.cmdButtonExit, Me.Frame1)
 X = VBA.Month(VBA.Date)
 Y = VBA.Year(VBA.Date)
 Me.sbMonth.Value = X
 Me.sbYear.Value = Y
 Call PutDaysInArray(X, Y)
 Call FindLatestDate(True)
 Me.Caption = VBA.UCase$(Format$(VBA.Date, "yyyy - mmmm ")) & VBA.Day(VBA.Date) _
              & "   (" & VBA.Date & ")"
 If ActiveCell.HasFormula Then Me.LabelCell.ControlTipText = "Has Formula"
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
 Exit Sub
BadForm:
 Beep
 Resume Next
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
            ByVal X As Single, ByVal Y As Single)
 On Error Resume Next
 Me.LabelCell.Caption = ActiveCell.Address(False, False)
 If ActiveCell.HasFormula Then Me.LabelCell.ControlTipText = "Has Formula" _
    Else Me.LabelCell.ControlTipText = "active cell"
End Sub
