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
Option Explicit

Public EditingCheck As Boolean

Public GroupCosts As String
Public GroupRevenue As String

'' BUTTONS===================================================================
Private Sub AddGroupButton_Click()
Call AddGroup(EditToggleCheckBox1.value)
End Sub

Private Sub BackButton_Click()
UserForm1.NewSearchTextBox.Text = GroupManagementForm.GroupNameTextBox.Text
GroupManagementForm.Hide
End Sub

Private Sub EditToggleCheckBox1_Click()
If EditToggleCheckBox1.value = True Then
    AddGroupButton.Caption = "Edit Selected Group"
Else
    AddGroupButton.Caption = "Add New Group"
End If
End Sub

'' TEXT BOXES===============================================================
Private Sub GroupCostsTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(GroupCostsTextBox, GroupCosts)
End Sub

Private Sub GroupCostsTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
GroupCostsTextBox.Text = StrManip.Convert2Currency(GroupCostsTextBox)
End Sub

Private Sub GroupRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(GroupRevenueTextBox, GroupRevenue)
End Sub

Private Sub GroupRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
GroupRevenueTextBox.Text = StrManip.Convert2Currency(GroupRevenueTextBox)
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
Call funcs.GetCalendar(GroupManagementForm.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub StartDateTextBox_Enter()
Call funcs.GetCalendar(GroupManagementForm.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub EndDateTextBox_Change()
If EndDateTextBox.Text = "" Then
    EndDateTextBox.Text = StartDateTextBox.Text
ElseIf StartDateTextBox.Text = "" Then
    ' Do nothing
ElseIf CDate(EndDateTextBox.Text) < CDate(StartDateTextBox.Text) Then
    MsgBox ("The start date cannot be after the end date")
    EndDateTextBox.Text = ""
End If
End Sub

Private Sub EndDateTextBox_Dblclick(ByVal cancel As MSForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call funcs.GetCalendar(GroupManagementForm.EndDateTextBox) ' Show Date Picker
End Sub

Private Sub EndDateTextBox_Enter()
Call funcs.GetCalendar(GroupManagementForm.EndDateTextBox) ' Show Date Picker
End Sub

'' USERFORM====================================================================
Private Sub Userform_Initialize()
' Set rowsource of listboxes
Call funcs.RefreshListBox("Non-Specific Defaults", 4, CategoryListBox)
Call funcs.RefreshListBox("Type-Specific Defaults", 1, TypeListBox)

' Don't do this because it doesn't always work
' Set default value of category
'CategoryListBox.ListIndex = 0
End Sub

Private Sub Userform_Activate()
' Set rowsource of listboxes
Call funcs.RefreshListBox("Non-Specific Defaults", 4, CategoryListBox)
Call funcs.RefreshListBox("Type-Specific Defaults", 1, TypeListBox)

' Don't do this because it doesn't always work
' Set default value of category
'CategoryListBox.ListIndex = 0
End Sub

'' FUNCTIONS=============================================================
Private Sub AddGroup(mode As Boolean)
' Add or edit group

' mode = False = Add group
' mode = True = Edit group

' Store which mode we're in publically
If mode = True Then
    EditingCheck = True
Else
    EditingCheck = False
End If

' Ensure everything that needs to be entered has been entered
If mode = True And UserForm1.GroupIDListBox.ListIndex = -1 Then
    MsgBox ("You are in edit mode. " & _
            "Please select a group to edit on the Manage Groups page")
ElseIf GroupNameTextBox.Text = "" Then
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
ElseIf GroupCostsTextBox.Text = "" Then
    MsgBox ("Please enter costs incurred specifically for the group of events, " _
            & "not easily divided across each event.")
ElseIf GroupRevenueTextBox.Text = "" Then
    MsgBox ("Please enter revenue acrued specifically for the group of events, " _
        & "not easily divided across each event.")
End If

Dim sheet As String
sheet = "Data"

Dim ws As Worksheet
Set ws = Sheets(sheet)

Dim my_row As Long
If mode = False Then
    ' Next available row
    my_row = ws.Cells(Rows.Count, 1).End(xlUp).row + 1
ElseIf mode = True Then
    ' Row corresponds to the row of the event selected by the user
    my_row = UserForm1.GroupIDListBox.ListIndex + 2
End If

' Add data given by user into spreadsheet
' G stands for "group"
ws.Cells(my_row, 1) = "G" & funcs.UUIDGenerator(CategoryListBox.value, _
                                StartDateTextBox.Text, GroupNameTextBox.Text)
ws.Cells(my_row, 72) = "G" & funcs.UUIDGenerator(CategoryListBox.value, _
                                StartDateTextBox.Text, GroupNameTextBox.Text)
ws.Cells(my_row, 2) = GroupNameTextBox.Text
ws.Cells(my_row, 3) = StrManip.ConvertDate(StartDateTextBox.Text)
ws.Cells(my_row, 24) = CategoryListBox.value
ws.Cells(my_row, 29) = TypeListBox.value
ws.Cells(my_row, 73) = True
ws.Cells(my_row, 74) = StrManip.ConvertDate(EndDateTextBox.Text)
ws.Cells(my_row, 75) = GroupCostsTextBox.Text
ws.Cells(my_row, 76) = GroupRevenueTextBox.Text
ws.Cells(my_row, 77) = GroupNameTextBox.Text

' Index of what was previously selected
Dim my_index As Integer
' Length of listbox. Lets us know if the edit means the group no longer appears
'   in the search
Dim my_length As Integer
If mode = False Then
    MsgBox ("Group Added")
Else
    my_index = UserForm1.GroupNameListBox.ListIndex
    my_length = UserForm1.GroupNameListBox.ListCount
    MsgBox ("Group Edited")
End If

' Update pivot table(s)
Call funcs.ChangeSource(sheet, "Analysis", "PivotTable1")

' Update listboxes
Call funcs.RefreshListBox("Data", 1, UserForm1.EventIDListBox)
Call funcs.RefreshListBox("Data", 1, UserForm1.GroupEventIDListBox)
Call funcs.RefreshListBox("Data", 72, UserForm1.GroupIDListBox)

' Refresh group search
Dim search As String
search = UserForm1.GroupSearchBox.Text
UserForm1.GroupSearchBox.Text = ""
UserForm1.GroupSearchBox.Text = search

' Change selected event to one just added, if one was added (and not edited)
If UserForm1.GroupIDListBox.ListCount > 0 And mode = False Then
    ' Select final item in listbox
    UserForm1.GroupIDListBox.ListIndex = UserForm1.GroupIDListBox.ListCount - 1
    ' Find the event ID
    UserForm1.HiddenGroupIDListBox.value = UserForm1.GroupIDListBox.value
    ' Set user-facing listboxes to correct event
    UserForm1.GroupNameListBox.ListIndex = UserForm1.HiddenGroupIDListBox.ListIndex
ElseIf UserForm1.GroupNameListBox.ListCount = my_length Then
    UserForm1.GroupNameListBox.ListIndex = my_index
End If
End Sub
