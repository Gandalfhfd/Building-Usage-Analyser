VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Events"
   ClientHeight    =   6768
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11172
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

Private Sub AddEventButton_Click()
' Check whether all of the information has been completed or not
If CategoryListBox.ListIndex = -1 Then
    MsgBox ("Please select a category")
    Exit Sub
ElseIf EventDateTextBox.Text = "" Then
    MsgBox ("Please select a date using the calendar. Double click on the text box to show the calendar.")
    Exit Sub
ElseIf NameTextBox.Text = "" Then
    MsgBox ("Please enter an event name")
    Exit Sub
ElseIf LocationListBox.ListIndex = -1 Then
    MsgBox ("Please select a location")
    Exit Sub
ElseIf RoomListBox.ListIndex = -1 Then
    MsgBox ("Please select a room")
    Exit Sub
ElseIf MorningCheckBox.value = 0 And AfternoonCheckBox.value = 0 And EveningCheckBox.value = 0 Then
    MsgBox ("Please select a time")
    Exit Sub
ElseIf TypeListBox.ListIndex = -1 Then
    MsgBox ("Please enter a type")
    Exit Sub
ElseIf AudienceListBox.ListIndex = -1 Then
    MsgBox ("Please enter an audience type")
    Exit Sub
ElseIf AuditoriumLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Auditorium")
    Exit Sub
ElseIf EgremontLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Egremont Room")
    Exit Sub
Else ' The user is allowed to create a new event

End If

Dim empty_row As Long
empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row + 1

Dim UUID As String
UUID = UUIDGenerator(CategoryListBox.value, EventDateTextBox.Text, NameTextBox.Text)

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
Worksheets("Data").Cells(empty_row, "A") = UUID
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

End Sub

Private Sub AddEventButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call AddEventButton_Click
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

If RoomListBox.value = "Auditorium" Then
    AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1024")
ElseIf RoomListBox.value = "Egremont Room" Then
    EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1024")
ElseIf RoomListBox.value = "Both" Then
    AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1024")
    EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1024")
ElseIf RoomListBox.value = "External" Then
    AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2")
    EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2")
End If

End Sub

Private Sub AuditoriumLayoutListBox_Change()
AuditoriumCapacityTextBox.Text = Worksheets("NonSpecificDefaults").Cells(AuditoriumLayoutListBox.ListIndex + 2, 7)
End Sub

Private Sub EgremontLayoutListBox_Change()
EgremontCapacityTextBox.Text = Worksheets("NonSpecificDefaults").Cells(EgremontLayoutListBox.ListIndex + 2, 9)
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
  max = IIf(x > y, x, y)
End Function

Public Function min(x, y As Variant) As Variant
   min = IIf(x < y, x, y)
End Function

Public Function TotalCapDecider() As String
If AuditoriumCapacityTextBox.Text & EgremontCapacityTextBox.Text = "" Then ' Set total capacity to blank if both are blank
    TotalCapacityTextBox.Text = ""
    TotalCapacity = TotalCapacityTextBox.Text
ElseIf AuditoriumCapacityTextBox.Text = "0" Then ' Ignore Auditorium Capacity if it is 0
    TotalCapacityTextBox.Text = EgremontCapacityTextBox.Text
    TotalCapacity = TotalCapacityTextBox.Text
ElseIf EgremontCapacityTextBox.Text = "0" Then ' Ignore Egremont Capacity if it is 0
    TotalCapacityTextBox.Text = AuditoriumCapacityTextBox.Text
    TotalCapacity = TotalCapacityTextBox.Text
ElseIf AuditoriumCapacityTextBox.Text = "" Then ' Ignore Auditorium Capacity if it is blank
    TotalCapacityTextBox.Text = EgremontCapacityTextBox.Text
    TotalCapacity = TotalCapacityTextBox.Text
ElseIf EgremontCapacityTextBox.Text = "" Then ' Ignore Egremont Capacity if it is blank
    TotalCapacityTextBox.Text = AuditoriumCapacityTextBox.Text
    TotalCapacity = TotalCapacityTextBox.Text
Else ' Find the max of the two
    TotalCapacityTextBox.Text = max(CDbl(AuditoriumCapacityTextBox.Text), CDbl(EgremontCapacityTextBox.Text))
    TotalCapacity = TotalCapacityTextBox.Text
End If
End Function
