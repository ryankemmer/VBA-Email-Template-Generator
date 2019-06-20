Attribute VB_Name = "Module1"
Public TRNumber As String
Public Model As String
Public TestType As String
Public StartDate As String

Sub User_Form_Show()

Application.ScreenUpdating = False
Application.Visible = False

Form.Show

End Sub

Sub Create_New_Template()

Application.ScreenUpdating = False
Application.Visible = False

Dim wrkName As String
wrkName = "C:\Users\Public\Desktop\" & TRNumber & " " & Model & " Email.xlsm"


With ActiveWorkbook
    Range("A1").Value = 1
    Range("A2").Value = TRNumber
    Range("A3").Value = Model
    Range("A4").Value = TestType
    Range("A5").Value = StartDate
    Range("A2").NumberFormat = "@"
End With

ActiveWorkbook.SaveCopyAs Filename:=wrkName

With ActiveWorkbook
    Range("A1").Value = ""
    Range("A2").Value = ""
    Range("A3").Value = ""
    Range("A4").Value = ""
    Range("A5").Value = ""
End With

ActiveWorkbook.Save

Dim answer As Integer
answer = MsgBox("New template created. Would you like to create another template?", vbYesNo + vbQuestion)
If answer = vbYes Then
    Call User_Form_Show
ElseIf answer = vbNo Then
    ActiveWorkbook.Close
Else
    ActiveWorkbook.Close
End If

ActiveWorkbook.Close

End Sub

Sub Create_Email()

Application.ScreenUpdating = False
Application.Visible = False

Dim OutApp As Object
Dim Outmail As Object
Dim strbody As String
Dim Booth As String
Dim TestType_Email As String
Dim TRNumber_Email As String
Dim Model_Email As String

'See if outlook is open
Call check_is_running

'Check the Date
Call check_date

'Set up Outlook
Set OutApp = CreateObject("Outlook.Application")
Set Outmail = OutApp.CreateItem(0)
    
'Determine Variables

TRNumber_Email = Range("A2").Value
Model_Email = Range("A3").Value
TestType_Email = Range("A4").Value

'Determine correct time (what booth it is coming out of)
If Time < 12 / 24 Then
    Booth = "cold booth this morning"
ElseIf Time > 12 / 24 Then
    Booth = "hot booth this afternoon"
Else
    Booth = ""
End If

'Determine Correct Folder
Dim FolderNumber As String
TRInt = CLng(TRNumber_Email)
    If (38500 <= TRInt) And (TRInt <= 38999) Then
        FolderNumber = "38500-38999"
    ElseIf (39000 <= TRInt) And (TRInt <= 39499) Then
        FolderNumber = "39000-39499"
    ElseIf (39500 <= TRInt) And (TRInt <= 39999) Then
        FolderNumber = "39500-39999"
    ElseIf (40000 <= TRInt) And (TRInt <= 40499) Then
        FolderNumber = "40000-40499"
    ElseIf (40500 <= TRInt) And (TRInt <= 40999) Then
        FolderNumber = "40500-40999"
    ElseIf (41000 <= TRInt) And (TRInt <= 41499) Then
        FolderNumber = "41000-41499"
    ElseIf (41500 <= TRInt) And (TRInt <= 41999) Then
        FolderNumber = "41500-41999"
    ElseIf (42000 <= TRInt) And (TRInt <= 42499) Then
        FolderNumber = "42000-42499"
    ElseIf (42500 <= TRInt) And (TRInt <= 42999) Then
        FolderNumber = "42500-42999"
    ElseIf (43000 <= TRInt) And (TRInt <= 43499) Then
        FolderNumber = "43000-43499"
    ElseIf (43500 <= TRInt) And (TRInt <= 43999) Then
        FolderNumber = "43500-43999"
    ElseIf (44000 <= TRInt) And (TRInt <= 44499) Then
        FolderNumber = "44000-44499"
    ElseIf (44500 <= TRInt) And (TRInt <= 44999) Then
        FolderNumber = "44500-44999"
    ElseIf (45000 <= TRInt) And (TRInt <= 45499) Then
        FolderNumber = "45000-45499"
    ElseIf (45500 <= TRInt) And (TRInt <= 45999) Then
        FolderNumber = "45500-45999"
    ElseIf (46000 <= TRInt) And (TRInt <= 46499) Then
        FolderNumber = "46000-46499"
    ElseIf (46500 <= TRInt) And (TRInt <= 46999) Then
        FolderNumber = "46500-46999"
    ElseIf (47000 <= TRInt) And (TRInt <= 47499) Then
        FolderNumber = "47000-47499"
End If
        

'Create Email Body
strbody = "<b>" & Now() & "</b>" & "<br><br>" & _
           "TR #" & TRNumber_Email & " " & Model_Email & " HFO Barrier Layer HIPS " & TestType_Email & " came out " & _
           "of the " & Booth & " with no findings. <br><br>" & _
           "<a href = ""\\subzero.com\Wisconsin\Madison\Teamwork\Projects\Testdata\" & FolderNumber & "\" & TRNumber_Email & """>Click here to view TR Folder</a><br>" & _
           "<br>Best,"

'Create message in outlook
On Error Resume Next
With Outmail
    .Display
    .To = "frank.bogat@subzero.com; tim.ingman@subzero.com; chris.dietsch@subzero.com; michael.spindler@subzero.com; matthew.wurz@subzero.com; stephen.ginos@subzero.com"
    .CC = "christopher.obuch@subzero.com; bob.zoladz@subzero.com; ryan.kemmer@subzero.com"
    .Bcc = ""
    .Subject = "Thermal Shock TR #" & TRNumber_Email & " " & Model_Email & " " & TestType_Email & " HFO Barrier Layer HIPS"
    .HTMLBody = strbody & .HTMLBody
End With
On Error GoTo 0

'Clean up outlook
Set Outmail = Nothing
Set OutApp = Nothing

ActiveWorkbook.Close

End Sub

Sub check_date()
    Dim xdate As Date
    Dim newDate As Date
    xdate = CDate(Range("A5"))
    newDate = DateAdd("d", 5, xdate)
    If newDate < Date Then
        MsgBox ("Template Expired on " & newDate)
        ActiveWorkbook.Close
    End If
  
End Sub

Sub check_is_running()

Dim Process As Object
Dim proc_name As String
Dim bl  As Boolean
proc_name = "Outlook.exe" '\change process name here
bl = False
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & proc_name & "'")
        If UCase(Process.Name) = UCase(proc_name) Then
            bl = True
            Exit For
        End If
    Next Process
    
    If bl = False Then
        Shell ("OUTLOOK")
    Else

    End If
End Sub

