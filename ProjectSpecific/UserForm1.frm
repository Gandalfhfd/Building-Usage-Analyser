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

End Sub

Private Sub AddEventButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call AddEventButton_Click
End Sub

Private Sub MultiPage1_Open() ' Doesn't seem to be called
MsgBox ("MultiPage1_Open Was Called")
End Sub

Private Sub EventDateTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call GetCalendar
End Sub

Private Sub EventDateTextBox_Enter()
Call GetCalendar
End Sub

Private Sub MultiPage1_Change()
LocationListBox.RowSource = ("NonSpecificDefaults!A2:A1024")
'Create list of categories based on some cells in the specified worksheet
CategoryListBox.RowSource = ("UserFormData!A2:A1024")
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

Function UUIDGenerator(category As String, eventDate As String, name As String) As String
UUIDGenerator = RmSpecialChars(category) & RmSpecialChars(eventDate) & RmSpecialChars(name)
End Function

Function RmSpecialChars(inputStr As String) As String
' List of chars we want to remove
Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?, ,/,:,',."

Dim char As Variant

RmSpecialChars = inputStr

' Iterate over SpecialCharacters and remove everything that matches
For Each char In Split(SpecialCharacters, ",")
    RmSpecialChars = Replace(RmSpecialChars, char, "")
Next
End Function
