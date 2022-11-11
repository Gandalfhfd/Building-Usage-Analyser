VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6765
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

' Set up a public variable which stores the total number of events
Public numEvents As Integer

Private Sub AddEventButton_Click()

' Check whether a category has been selected or not
If CategoryListBox.ListIndex = -1 Then
    MsgBox ("Please select a category")
    Exit Sub
Else
    ' Add 1 to the number of events for that category
    Worksheets("UserFormData").Cells(CategoryListBox.ListIndex + 2, 2).value = _
        Worksheets("UserFormData").Cells(CategoryListBox.ListIndex + 2, 2).value + 1
End If

Dim empty_row As Long
empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row + 1

' Add data into spreadsheet
Worksheets("Data").Cells(empty_row, "A") = CategoryListBox.value & EventDateTextBox.Text _
    & NameTextBox.Text

Worksheets("Data").Cells(empty_row, "B") = NameTextBox.Text
Worksheets("Data").Cells(empty_row, "C") = EventDateTextBox.Text
Worksheets("Data").Cells(empty_row, "X") = CategoryListBox.value
End Sub

Private Sub AddEventButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call AddEventButton_Click
End Sub

Private Sub AddEventButton_Enter()

End Sub

Private Sub MultiPage1_Open() ' Doesn't seem to be called

MsgBox ("MultiPage1 Was Called")

End Sub

Private Sub EventDateTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call GetCalendar
End Sub

Private Sub EventDate0TextBox_Change()

End Sub

Private Sub MultiPage1_Change()
'Create list of categories based on some cells in the specified worksheet
CategoryListBox.RowSource = ("UserFormData!A2:A1024")
End Sub

Private Sub NoEventsButton_Click()
MsgBox (numEvents & " events in total.")
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

Sub UserForm_Initialize()

numEvents = Worksheets("UserFormData").Range("B2")

End Sub

Sub GetCalendar()
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

Function UUIDGenerator()
 
End Function

