VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Template Generator"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   9460.001
   OleObjectBlob   =   "Email Template UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()

Unload Me

End Sub


Private Sub create_template_Click()

TRNumber = TRNumber_log.Value
Model = Model_log.Value
StartDate = mm.Value & "/" & dd.Value & "/" & yyyy.Value
If (DoorButton.Value = True) Then
    TestType = "Doors"
ElseIf (CabinetButton.Value = True) Then
    TestType = "Cabinets"
ElseIf (UDrawerButton.Value = True) Then
    TestType = "Upper Drawers"
ElseIf (LDrawerButton.Value = True) Then
    TestType = "Lower Drawers"
Else
    MsgBox ("Extremely Fatal Error")
End If

Unload Me

Call Create_New_Template

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Initialize()

'Initialize text boxes

TRNumber_log = ""
Model_log = ""
DoorButton.Value = True

'Initialize date boxes
With mm
    Dim i As Integer
        For i = 1 To 12
            .AddItem (i)
        Next i
End With

With dd
    Dim j As Integer
        For j = 1 To 31
            .AddItem (j)
        Next j
End With

With yyyy
    Dim d As Integer
    d = Year(Now)
    .AddItem (d)
    .AddItem (d + 1)
End With

End Sub
