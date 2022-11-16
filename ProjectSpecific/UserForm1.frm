VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Events"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Used to prevent unecessary refreshing of ListBoxes
Public counter As Integer

' Store user input so that it can be restored if they make a mistake
Public AuditoriumCapacity As String
Public EgremontCapacity As String
Public TotalCapacity As String

'' BUTTON CLICKING===============================================================

Private Sub AddEventButton1_Click()
Call AddEvent
End Sub

Private Sub AddEventButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call AddEvent
End Sub

Private Sub AddEventButton2_Click()
Call AddEvent
End Sub

Private Sub AddEventButton2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call AddEvent
End Sub

Private Sub SearchButton_Click()

' Not sure what this stuff does
Dim c As Range
Dim firstAddress As String

' Search for a UUID. Display address if found.
' Future improvement: a "goto" button which takes the user to the cell.
' Future improvement: some way of knowing if multiple references are found.

If SearchBox.value = "" Then
    MsgBox ("Please enter a UUID in the search box.") ' Error msg
Else
    With Worksheets(2).Range("A:A") ' Look in worksheet 2 over this range of cells
        Set c = .Find(SearchBox.Text, LookIn:=xlValues)
        If Not c Is Nothing Then ' If anything is found, then...
            firstAddress = c.Address
            MsgBox ("UUID Found in cell " & c.Address)
        Else
            MsgBox ("UUID Not Found")
        End If
    End With
End If
   
End Sub

'' TEXT BOXES===============================================================

Private Sub EventDateTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call GetCalendar
End Sub

Private Sub EventDateTextBox_Enter()
Call GetCalendar
End Sub

Private Sub AuditoriumCapacityTextBox_Change()
' Input validation
If CheckIfNonNegInt(AuditoriumCapacityTextBox.Text) = False Then
    AuditoriumCapacityTextBox.Text = AuditoriumCapacity
Else
    AuditoriumCapacityTextBox.Text = RmSpecialChars(AuditoriumCapacityTextBox.Text)
    AuditoriumCapacity = AuditoriumCapacityTextBox.Text
End If

Call TotalCapDecider
End Sub

Private Sub EgremontCapacityTextBox_Change()
' Input validation
If CheckIfNonNegInt(EgremontCapacityTextBox.Text) = False Then
    EgremontCapacityTextBox.Text = EgremontCapacity
Else
    EgremontCapacityTextBox.Text = RmSpecialChars(EgremontCapacityTextBox.Text)
    EgremontCapacity = EgremontCapacityTextBox.Text
End If

Call TotalCapDecider
End Sub

Private Sub TotalCapacityTextBox_Change()
If CheckIfNonNegInt(TotalCapacityTextBox.Text) = False Then
    TotalCapacityTextBox.Text = TotalCapacity
Else
    TotalCapacity = TotalCapacityTextBox.Text
End If
End Sub

'' LIST BOXES===============================================================

Private Sub LocationListBox_Change()
' Throw an error if the location doesn't match the room selected (external should be selected for anything not in Kirkgate)
If LocationListBox.value <> "Kirkgate" And RoomListBox.value <> "External" And RoomListBox.ListIndex <> -1 Then
    MsgBox ("'External' should refer to an Arts out West venue.")
ElseIf LocationListBox.value = "Kirkgate" And RoomListBox.value = "External" Then
    MsgBox ("'External' should refer to an Arts out West venue.")
End If
End Sub

Private Sub RoomListBox_Change()
If LocationListBox.value <> "Kirkgate" And RoomListBox.value <> "External" And LocationListBox.ListIndex <> -1 Then
    MsgBox ("'External' should refer to an Arts out West venue.")
ElseIf LocationListBox.value = "Kirkgate" And RoomListBox.value = "External" Then
    MsgBox ("'External' should refer to an Arts out West venue.")
End If

' Decide what to show regarding capacity options
CapacityListBoxDecider

End Sub

Private Sub AuditoriumLayoutListBox_Change()
AuditoriumCapacityTextBox.Text = Worksheets("NonSpecificDefaults").Cells(AuditoriumLayoutListBox.ListIndex + 2, 7)
End Sub

Private Sub EgremontLayoutListBox_Change()
EgremontCapacityTextBox.Text = Worksheets("NonSpecificDefaults").Cells(EgremontLayoutListBox.ListIndex + 2, 9)
End Sub

Private Sub AudienceListBox_Change()
' Decide what to show regarding capacity options
Call CapacityListBoxDecider
End Sub
'' MULTIPAGE===============================================================

Private Sub MultiPage1_Change()
' Add items into listboxes based on cells in specified worksheets
If counter <> 1 Then
    LocationListBox.RowSource = ("NonSpecificDefaults!A2:A1024")
    RoomListBox.RowSource = ("NonSpecificDefaults!B2:B1024")
    CategoryListBox.RowSource = ("NonSpecificDefaults!D2:D1024")
    AudienceListBox.RowSource = ("NonSpecificDefaults!E2:E1024")
    AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1024")
    EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1024")
    TypeListBox.RowSource = ("UserFormData!A2:A1024")
End If

' Sop this from happening again
counter = 1
End Sub

'' FUNCTIONS===============================================================

Sub GetCalendar() ' Calendar format
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate(DateFontSize:=11, _
        BackgroundColor:=RGB(242, 248, 238), _
        HeaderColor:=RGB(84, 130, 53), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(226, 239, 218), _
        SubHeaderFontColor:=RGB(55, 86, 35), _
        DateColor:=RGB(242, 248, 238), _
        DateFontColor:=RGB(55, 86, 35), _
        TrailingMonthFontColor:=RGB(106, 163, 67), _
        DateHoverColor:=RGB(198, 224, 180), _
        DateSelectedColor:=RGB(169, 208, 142), _
        TodayFontColor:=RGB(255, 0, 0))
If dateVariable <> 0 Then UserForm1.EventDateTextBox = Format(dateVariable, "dd/mm/yyyy")
End Sub

Public Function UUIDGenerator(category As String, eventDate As String, name As String) As String
' Generate uniqueish UUID. Not unique if the same event is added twice within a second
UUIDGenerator = RmSpecialChars(category) & RmSpecialChars(eventDate) _
    & RmSpecialChars(name) & Format(Now, "ss")
End Function

Public Function RmSpecialChars(inputStr As String) As String
' List of chars we want to remove
Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?, ,/,:,',."
Const CommaCharacter As String = ","
Dim char As Variant

RmSpecialChars = inputStr

' Iterate over SpecialCharacters and remove everything that matches
For Each char In Split(SpecialCharacters, ",")
    RmSpecialChars = Replace(RmSpecialChars, char, "")
Next

' Remove commas
For Each char In Split(CommaCharacter, ".")
    RmSpecialChars = Replace(RmSpecialChars, char, "")
Next
End Function

Public Function CheckIfNonNegInt(inputStr As String) As Boolean
If inputStr = "" Then ' If blank, ignore
    CheckIfNonNegInt = True
ElseIf IsNumeric(inputStr) = False Then ' Check it can be conveted to a number
    CheckIfNonNegInt = False
ElseIf Round(CDbl(inputStr)) <> CDbl(inputStr) Then ' Check it is an integer
    CheckIfNonNegInt = False
ElseIf CDbl(inputStr) < 0 Then ' Check it is >= 0. CDbl used to prevent overflow.
    CheckIfNonNegInt = False
Else ' Then it must be a non-negative integer
    CheckIfNonNegInt = True
End If
End Function

Public Function max(x, y As Variant) As Variant
' Find max of two numbers
  max = IIf(x > y, x, y)
End Function

Public Function min(x, y As Variant) As Variant
' Find min of two numbers
   min = IIf(x < y, x, y)
End Function

Public Function TotalCapDecider() As String ' Highly non-generic function. Sorry!
If AuditoriumCapacityTextBox.Text & EgremontCapacityTextBox.Text = "" Then ' Set total capacity to blank if both are blank
    TotalCapacityTextBox.Text = ""
ElseIf AuditoriumCapacityTextBox.Text = "0" Then ' Ignore Auditorium Capacity if it is 0
    TotalCapacityTextBox.Text = EgremontCapacityTextBox.Text
ElseIf EgremontCapacityTextBox.Text = "0" Then ' Ignore Egremont Capacity if it is 0
    TotalCapacityTextBox.Text = AuditoriumCapacityTextBox.Text
ElseIf AuditoriumCapacityTextBox.Text = "" Then ' Ignore Auditorium Capacity if it is blank
    TotalCapacityTextBox.Text = EgremontCapacityTextBox.Text
ElseIf EgremontCapacityTextBox.Text = "" Then ' Ignore Egremont Capacity if it is blank
    TotalCapacityTextBox.Text = AuditoriumCapacityTextBox.Text
Else ' Find the max of the two
    TotalCapacityTextBox.Text = min(CDbl(AuditoriumCapacityTextBox.Text), CDbl(EgremontCapacityTextBox.Text))
End If

TotalCapacity = TotalCapacityTextBox.Text
End Function

Private Function AddEvent()

' Check whether all of the information has been completed or not
If NameTextBox.Text = "" Then
    MsgBox ("Please enter an event name")
    Exit Function
ElseIf EventDateTextBox.Text = "" Then
    MsgBox ("Please select a date using the calendar. Double click on the text box to show the calendar.")
    Exit Function
ElseIf CategoryListBox.ListIndex = -1 Then
    MsgBox ("Please select a category")
    Exit Function
ElseIf TypeListBox.ListIndex = -1 Then
    MsgBox ("Please enter a type")
    Exit Function
ElseIf LocationListBox.ListIndex = -1 Then
    MsgBox ("Please select a location")
    Exit Function
ElseIf RoomListBox.ListIndex = -1 Then
    MsgBox ("Please select a room")
    Exit Function
ElseIf AudienceListBox.ListIndex = -1 Then
    MsgBox ("Please enter an audience type")
Exit Function
ElseIf MorningCheckBox.value = 0 And AfternoonCheckBox.value = 0 And EveningCheckBox.value = 0 Then
    MsgBox ("Please select a time")
    Exit Function
ElseIf AuditoriumLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Auditorium")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
ElseIf EgremontLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Egremont Room")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
Else ' The user is allowed to create a new event
End If

' Find next empty row
Dim empty_row As Long
empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row + 1

' Add default data into spreadsheet can be overridden by user in the future
Dim i As Integer
For i = 0 To 5
    ' Add default minutes worked by each volunteer category, depending on event category selected
    Worksheets("Data").Cells(empty_row, i + 18) = Worksheets("UserFormData").Cells(CategoryListBox.ListIndex + 2, i + 3)
Next

' Bar gross profit bit
Worksheets("Data").Cells(empty_row, 26) = Worksheets("NonSpecificDefaults").Cells(2, 3)
Worksheets("Data").Cells(empty_row, 27) = "=RC[-2]*RC[-1]" ' Worksheets("NonSpecificDefaults").Cells(2, 3) * Worksheets("Data").Cells(empty_row, 25)

' Add data given by user into spreadsheet
Worksheets("Data").Cells(empty_row, "A") = UUIDGenerator(CategoryListBox.value, EventDateTextBox.Text, NameTextBox.Text)
Worksheets("Data").Cells(empty_row, "B") = NameTextBox.Text
Worksheets("Data").Cells(empty_row, "C") = EventDateTextBox.Text
Worksheets("Data").Cells(empty_row, "D") = LocationListBox.value
Worksheets("Data").Cells(empty_row, "X") = CategoryListBox.value
Worksheets("Data").Cells(empty_row, 28) = RoomListBox.value
Worksheets("Data").Cells(empty_row, 5) = MorningCheckBox.value
Worksheets("Data").Cells(empty_row, 6) = AfternoonCheckBox.value
Worksheets("Data").Cells(empty_row, 7) = EveningCheckBox.value
Worksheets("Data").Cells(empty_row, 29) = TypeListBox.value
Worksheets("Data").Cells(empty_row, 30) = AudienceListBox.value
Worksheets("Data").Cells(empty_row, 31) = EgremontLayoutListBox.value
Worksheets("Data").Cells(empty_row, 32) = AuditoriumLayoutListBox.value
Worksheets("Data").Cells(empty_row, 33) = TotalCapacityTextBox.Text
End Function

Private Function AuditoriumUsed()
AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1024")
AuditoriumCapacityTextBox.Locked = False
TotalCapacityTextBox.Locked = False
End Function

Private Function EgremontUsed()
EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1024")
EgremontCapacityTextBox.Locked = False
TotalCapacityTextBox.Locked = False
End Function

Private Function BothUsed()
Call AuditoriumUsed
Call EgremontUsed
End Function

Private Function AuditoriumNotUsed()
AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2")
AuditoriumCapacityTextBox.Locked = True
AuditoriumCapacityTextBox.value = "0"
End Function

Private Function EgremontNotUsed()
EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2")
EgremontCapacityTextBox.Locked = True
EgremontCapacityTextBox = "0"
End Function

Private Function NoneUsed()
Call AuditoriumNotUsed
Call EgremontNotUsed
TotalCapacityTextBox.Locked = True
TotalCapacityTextBox.Text = "0"
End Function

Private Function CapacityListBoxDecider()
' Decide what to show and not show when the user selects options for audience and room
If AudienceListBox.value = "None" Then ' "None" for audience overrides all
    Call NoneUsed
ElseIf RoomListBox.value = "Auditorium" Then ' Auditorium only
    Call AuditoriumUsed
    Call EgremontNotUsed
ElseIf RoomListBox.value = "Egremont Room" Then ' Egremont room only
    Call EgremontUsed
    Call AuditoriumNotUsed
ElseIf RoomListBox.value = "Both" Then ' Both rooms used
    Call BothUsed
ElseIf RoomListBox.value = "External" Then ' Neither rooms will be used
    Call NoneUsed
    ' The user must, then, enter the total capacity. So we must allow that.
    TotalCapacityTextBox.Locked = False
Else ' AKA nothing selected
    Call BothUsed
End If
End Function
