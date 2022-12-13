VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Groups"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' TEXT BOXES===============================================================

Private Sub AddGroupButton_Click()
Call AddGroup(EditToggleCheckBox1.value)
End Sub

Private Sub StartDateTextBox_Change()
If EndDateTextBox.Text = "" Then
    EndDateTextBox.Text = StartDateTextBox.Text
ElseIf StartDateTextBox.Text = "" Then
    StartDateTextBox.Text = EndDateTextBox.Text
ElseIf CDate(EndDateTextBox.Text) < CDate(StartDateTextBox.Text) Then
    MsgBox ("The start date cannot be after the end date")
    StartDateTextBox.Text = ""
End If
End Sub

Private Sub StartDateTextBox_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call funcs.GetCalendar(UserForm2.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub StartDateTextBox_Enter()
Call funcs.GetCalendar(UserForm2.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub EndDateTextBox_Change()
If EndDateTextBox.Text = "" Then
    EndDateTextBox.Text = StartDateTextBox.Text
ElseIf CDate(EndDateTextBox.Text) < CDate(StartDateTextBox.Text) Then
    MsgBox ("The start date cannot be after the end date")
    EndDateTextBox.Text = ""
End If
End Sub

Private Sub EndDateTextBox_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call funcs.GetCalendar(UserForm2.EndDateTextBox) ' Show Date Picker
End Sub

Private Sub EndDateTextBox_Enter()
Call funcs.GetCalendar(UserForm2.EndDateTextBox) ' Show Date Picker
End Sub

Private Sub Userform_Initialize()
' Set rowsource of listboxes
Call funcs.RefreshListBox("Non-Specific Defaults", 4, CategoryListBox)
Call funcs.RefreshListBox("Type-Specific Defaults", 1, TypeListBox)

' Set default value of category
CategoryListBox.ListIndex = 0
End Sub

Private Sub AddGroup(mode As Boolean)
' Add or edit group

' mode = False = Add event
' mode = True = Edit event

' Store which mode we're in publically
If mode = True Then
    EditingCheck = True
Else
    EditingCheck = False
End If



End Sub
