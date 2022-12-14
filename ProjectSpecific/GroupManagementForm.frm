VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupManagementForm 
   Caption         =   "Groups"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   OleObjectBlob   =   "GroupManagementForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupManagementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' TEXT BOXES===============================================================

Private Sub AddGroupButton_Click()
Call AddGroup(EditToggleCheckBox1.value)
End Sub

Private Sub BackButton_Click()
GroupManagementForm.Hide
End Sub

Private Sub EditToggleCheckBox1_Click()
If EditToggleCheckBox1.value = True Then
    AddGroupButton.Caption = "Edit Selected Group"
Else
    AddGroupButton.Caption = "Add New Group"
End If
End Sub

Private Sub StartDateTextBox_change()
If EndDateTextBox.Text = "" Then
    EndDateTextBox.Text = StartDateTextBox.Text
ElseIf StartDateTextBox.Text = "" Then
    StartDateTextBox.Text = EndDateTextBox.Text
ElseIf CDate(EndDateTextBox.Text) < CDate(StartDateTextBox.Text) Then
    MsgBox ("The start date cannot be after the end date")
    StartDateTextBox.Text = ""
End If
End Sub

Private Sub StartDateTextBox_DblClick(ByVal cancel As MSForms.ReturnBoolean)
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

Private Sub EndDateTextBox_Dblclick(ByVal cancel As MSForms.ReturnBoolean)
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

' Ensure everything that needs to be entered has been entered
If GroupNameTextBox.Text = "" Then
    MsgBox ("Please enter a name")
    Exit Sub
ElseIf StartDateTextBox.Text = "" Then
    MsgBox ("Please enter a start date")
    Exit Sub
ElseIf EndDateTextBox.Text = "" Then
    MsgBox ("Please enter an end date")
    Exit Sub
ElseIf CategoryListBox.ListIndex = -1 Then
    MsgBox ("Please enter a category")
    Exit Sub
ElseIf TypeListBox.ListIndex = -1 Then
    MsgBox ("Please enter a type")
    Exit Sub
End If


End Sub
