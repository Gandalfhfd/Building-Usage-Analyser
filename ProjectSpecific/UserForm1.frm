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
Public NumTicketsSold As String
Public BoxOfficeRevenue As String
Public SupportRevenue As String
Public FilmCost As String
Public Accommodation As String
Public ArtistFood As String
Public Heating As String
Public Lighting As String
Public MiscCost As String
Public BarRevenue As String
Public BarMargin As String

'' BUTTON CLICKING===============================================================

Private Sub EventButton1_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub DeleteButton_Click()
' This breaks if EventID is not unique

' If nothing has been selected, do nothing but tell the user
If EventIDListBox.ListIndex = -1 Then
    MsgBox ("Please select an event to delete")
    Exit Sub
End If

' Store response
Dim response
response = MsgBox("Are you sure you want to delete " & "test", vbYesNo)
If response = vbNo Then
    ' Exit the sub because they don't want to delete anything
    Exit Sub
End If

' Store location of row to be deleted
Dim location As Variant
location = funcs.search(EventIDListBox.value, "Data")

' Unlikely, but possible that the search fails.
If location(0) = 0 Then
    MsgBox ("EventID Not found. ")
End If

' Delete entire row corresponding to selected event
Sheets("Data").Rows(location(0)).Delete
End Sub

Private Sub ImportSelectedButton_Click()
' Find the row of the selected event
Dim event_row As Long
event_row = funcs.search(SearchBox.Text, "Data")(0)

' Input validation. If tests fail, sub must be exited.
If SearchBox.Text = "" Then
    MsgBox ("Please enter the Event ID into the search box so that " & _
            "we know which event to import the data into")
    Exit Sub
ElseIf event_row = "0" Then
    MsgBox ("Event ID could not be found")
    Exit Sub
End If

' Import data into the event that has been selected
Call ImportFromTicketsolve("Selected")
End Sub

Private Sub ImportPreviousButton_Click()
' Import data into the event which was most recently added
Call ImportFromTicketsolve("Previous")
End Sub

Private Sub NameSearchButton_Click()
Dim myAddress As Variant

' Search for an Event ID. Display address if found.
' Future improvement: a "goto" button which takes the user to the cell.
' Future improvement: some way of knowing if multiple references are found.

If NameSearchTextBox.value = "" Then
    MsgBox ("Please enter the name of the event in the search box.")
ElseIf funcs.search(NameSearchTextBox.Text, "Data")(0) = 0 Then
    MsgBox ("Event Not Found")
Else
    myAddress = funcs.search(NameSearchTextBox.Text, "Data")
    MsgBox ("Event found on row " & myAddress(0))
End If

End Sub

Private Sub SearchButton_Click()

Dim myAddress As Variant

' Search for an Event ID. Display address if found.
' Future improvement: a "goto" button which takes the user to the cell.
' Future improvement: some way of knowing if multiple references are found.

If SearchBox.value = "" Then
    MsgBox ("Please enter an Event ID in the search box.")
ElseIf funcs.search(SearchBox.Text, "Data")(0) = 0 Then
    MsgBox ("Event ID Not Found")
Else
    myAddress = funcs.search(SearchBox.Text, "Data")
    MsgBox ("Event ID found on row " & myAddress(0))
End If

End Sub

'' TEXT BOXES===============================================================

Private Sub EventDateTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean) ' Re-show Date Picker if box has already been entered
Call GetCalendar ' Show Date Picker
End Sub

Private Sub EventDateTextBox_Enter()
Call GetCalendar ' Show Date Picker
End Sub

Private Sub AuditoriumCapacityTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(AuditoriumCapacityTextBox, AuditoriumCapacity)
Call TotalCapDecider
End Sub

Private Sub EgremontCapacityTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(EgremontCapacityTextBox, EgremontCapacity)
Call TotalCapDecider
End Sub

Private Sub TotalCapacityTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(TotalCapacityTextBox, TotalCapacity)
End Sub

Private Sub NumTicketsSoldTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(NumTicketsSoldTextBox, NumTicketsSold)
End Sub

Private Sub BoxOfficeRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BoxOfficeRevenueTextBox, BoxOfficeRevenue)
End Sub

Private Sub BoxOfficeRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
BoxOfficeRevenueTextBox.Text = StrManip.Convert2Currency(BoxOfficeRevenueTextBox)
End Sub

Private Sub SupportRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(SupportRevenueTextBox, SupportRevenue)
End Sub

Private Sub SupportRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
SupportRevenueTextBox.Text = StrManip.Convert2Currency(SupportRevenueTextBox)
End Sub

Private Sub FilmCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmCostTextBox, FilmCost)
End Sub

Private Sub FilmCostTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
FilmCostTextBox.Text = StrManip.Convert2Currency(FilmCostTextBox)
End Sub

Private Sub AccommodationTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(AccommodationTextBox, Accommodation)
End Sub

Private Sub AccommodationTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
AccommodationTextBox.Text = StrManip.Convert2Currency(AccommodationTextBox)
End Sub

Private Sub ArtistFoodTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(ArtistFoodTextBox, ArtistFood)
End Sub

Private Sub ArtistFoodTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
ArtistFoodTextBox.Text = StrManip.Convert2Currency(ArtistFoodTextBox)
End Sub

Private Sub HeatingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HeatingTextBox, Heating)
End Sub

Private Sub HeatingTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
HeatingTextBox.Text = StrManip.Convert2Currency(HeatingTextBox)
End Sub

Private Sub LightingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(LightingTextBox, Lighting)
End Sub

Private Sub LightingTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
LightingTextBox.Text = StrManip.Convert2Currency(LightingTextBox)
End Sub

Private Sub MiscCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(MiscCostTextBox, MiscCost)
End Sub

Private Sub MiscCostTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
MiscCostTextBox.Text = StrManip.Convert2Currency(MiscCostTextBox)
End Sub

Private Sub BarRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BarRevenueTextBox, BarRevenue)
End Sub

Private Sub BarRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only real numbers are input
BarRevenueTextBox.Text = StrManip.Convert2Currency(BarRevenueTextBox)
End Sub

Private Sub BarMarginTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BarMarginTextBox, BarMargin)
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

Private Sub EventIDListBox_Change()
' Display Event ID in search box
SearchBox.Text = EventIDListBox.value

Dim nameLocation As Variant
nameLocation = funcs.search(EventIDListBox.value, "Data")

If nameLocation(0) = 0 Then
    MsgBox ("Event ID could not be found. YOU SHOULD NOT SEE THIS. ERROR IN EventIDListBox_Change())")
End If

' Display the event name in search box.
' It will always be found in the second column, so hard coding is ok.
NameSearchTextBox.Text = Worksheets("Data").Cells(nameLocation(0), 2)

' So the user always knows which event has been selected.
' Might want to hide that if they're not in edit mode, idk.
EventIDUpdaterLabel1.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel2.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel3.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel4.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel5.Caption = "Selected Event ID: " & EventIDListBox.value
End Sub

'' MULTIPAGE===============================================================

Private Sub MultiPage1_Change()
' Add items into listboxes based on cells in specified worksheets
If counter <> 1 Then
    LocationListBox.RowSource = ("NonSpecificDefaults!A2:A1048576")
    RoomListBox.RowSource = ("NonSpecificDefaults!B2:B1048576")
    CategoryListBox.RowSource = ("NonSpecificDefaults!D2:D1048576")
    AudienceListBox.RowSource = ("NonSpecificDefaults!E2:E1048576")
    AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1048576")
    EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1048576")
    TypeListBox.RowSource = ("UserFormData!A2:A1048576")
    EventIDListBox.RowSource = ("Data!A2:A32768")
End If

' Sop this from happening again
counter = 1

End Sub

'' CHECKBOXES==============================================================
Private Sub EditToggleCheckBox1_Click()
' Change some button captions and checkbox values
Call ToggleEditMode(EditToggleCheckBox1.value)
End Sub

Private Sub EditToggleCheckBox2_Click()
' Change some button captions and checkbox values
Call ToggleEditMode(EditToggleCheckBox2.value)
End Sub

Private Sub EditToggleCheckBox3_Click()
' Change some button captions and checkbox values
Call ToggleEditMode(EditToggleCheckBox3.value)
End Sub

'' FUNCTIONS===============================================================
' Should move these into their own modules

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
UUIDGenerator = InptValid.RmSpecialChars(name) & InptValid.RmSpecialChars(category) _
                & funcsRmSpecialChars(eventDate) & Format(Now, "ss")
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
Else ' Find the min of the two
    TotalCapacityTextBox.Text = funcs.min(CDbl(AuditoriumCapacityTextBox.Text), CDbl(EgremontCapacityTextBox.Text))
End If

TotalCapacity = TotalCapacityTextBox.Text
End Function

Private Function AddEvent(mode As Boolean)
' mode = True = Add event
' mode = False = Edit event

' Check whether the required information has been completed or not
If mode = True And EventIDListBox.ListIndex = -1 Then
    MsgBox ("You are in edit mode. Please select an event to edit on the Home page")
    Exit Function
ElseIf NameTextBox.Text = "" Then
    MsgBox ("Please enter an event name")
    Exit Function
ElseIf EventDateTextBox.Text = "" Then
    MsgBox ("Please select a date using the calendar. Double click on the text box to show the calendar.")
    Exit Function
ElseIf CategoryListBox.ListIndex = -1 Then ' .ListIndex = -1 means nothing has been selected yet
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
ElseIf TicketedOptionButton.value = False And NonTicketedOptionButton.value = False Then
    MsgBox ("Please select 'Ticketed' or 'Non-Ticketed'")
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

Dim my_row As Long
If mode = False Then
    ' Next available row
    my_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row + 1
ElseIf mode = True Then
    ' Row corresponds to the row of the event selected by the user
    my_row = funcs.search(SearchBox.Text, "Data")(0)
End If

' Add default data into spreadsheet can be overridden by user in the future
Dim i As Integer
For i = 0 To 5
    ' Add default minutes worked by each volunteer category, depending on event category selected
    Worksheets("Data").Cells(my_row, i + 18) = Worksheets("UserFormData").Cells(CategoryListBox.ListIndex + 2, i + 3)
Next

Dim sheet As String
sheet = "Data"

' Bar gross profit bit
Worksheets(sheet).Cells(my_row, 26) = Worksheets("NonSpecificDefaults").Cells(2, 3)
Worksheets(sheet).Cells(my_row, 27) = "=RC[-2]*RC[-1]"

' Add data given by user into spreadsheet
' Basic Info
Worksheets(sheet).Cells(my_row, "A") = UUIDGenerator(CategoryListBox.value, EventDateTextBox.Text, NameTextBox.Text)
Worksheets(sheet).Cells(my_row, "B") = NameTextBox.Text
Worksheets(sheet).Cells(my_row, "C") = StrManip.ConvertDate(EventDateTextBox.Text)
Worksheets(sheet).Cells(my_row, "D") = LocationListBox.value
Worksheets(sheet).Cells(my_row, "X") = CategoryListBox.value
Worksheets(sheet).Cells(my_row, 28) = RoomListBox.value
Worksheets(sheet).Cells(my_row, 5) = MorningCheckBox.value
Worksheets(sheet).Cells(my_row, 6) = AfternoonCheckBox.value
Worksheets(sheet).Cells(my_row, 7) = EveningCheckBox.value
Worksheets(sheet).Cells(my_row, 29) = TypeListBox.value
Worksheets(sheet).Cells(my_row, 30) = AudienceListBox.value
' If we're selling tickets ourselves or not
If TicketedOptionButton.value = True Then
    Worksheets(sheet).Cells(my_row, 46) = "True"
Else
    ' Since validation has already taken place, at least one of the options has been selected
    Worksheets(sheet).Cells(my_row, 46) = "False"
End If

' Layout & Capacity
Worksheets(sheet).Cells(my_row, 31) = EgremontLayoutListBox.value
Worksheets(sheet).Cells(my_row, 32) = AuditoriumLayoutListBox.value
Worksheets(sheet).Cells(my_row, 33) = TotalCapacityTextBox.Text

' Revenue & Costs


MsgBox ("Event Added")
End Function

Private Function AuditoriumUsed()
' To be called when Auditorium is being used
AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2:F1048576")
AuditoriumCapacityTextBox.Locked = False
TotalCapacityTextBox.Locked = False
End Function

Private Function EgremontUsed()
' To be called when Egremont room is being used
EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2:H1048576")
EgremontCapacityTextBox.Locked = False
TotalCapacityTextBox.Locked = False
End Function

Private Function BothUsed()
' To be called when both Auditorium and Egremont rooms are being used
Call AuditoriumUsed
Call EgremontUsed
End Function

Private Function AuditoriumNotUsed()
' To be called when Auditorium is not being used
AuditoriumLayoutListBox.RowSource = ("NonSpecificDefaults!F2")
AuditoriumCapacityTextBox.Locked = True
AuditoriumCapacityTextBox.value = "0"
End Function

Private Function EgremontNotUsed()
' To be called when Egremont room is not being used
EgremontLayoutListBox.RowSource = ("NonSpecificDefaults!H2")
EgremontCapacityTextBox.Locked = True
EgremontCapacityTextBox = "0"
End Function

Private Function NoneUsed()
' To be called when neither room is being used
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

Function ImportCell(keyword As String, importSheet As String, exportSheet As String, _
                        exportCell() As Variant, offset() As Variant, Optional errorMsg As String, _
                        Optional substitute As String = "N/A") As Boolean
' MOVE TO FUNCTIONS SECTION, OR funcs WHEN DONE

' Designed to transplant value of a cell from one sheet to another in the same workbook.
' Could be modified to go from one workbook to another at some point, but that would require modification of the search function too.

' INPUTS
' All arrays must contain 2 items, like coordinates. (row, column) 1 indexing.
' keyword = word found near our target cell
' importSheet = name of sheet target cell is found in
' exportSheet = name of sheet we want to send target value to
' exportCell = position of cell we want to send target value to. Must be 2-item array.
' offset = 2-item array describing the offset of the target cell to the keyword
' errorMsg = (OPTIONAL) message displayed in MsgBox if keyword is not found. errorMsg is always defined
' substitute = (OPTIONAL) what will go into the exportCell if keyword is not found

' OUTPUTS
' ImportCell = Boolean which states whether this was successful or not

' Set default value of errorMsg, if user hasn't supplied one
If IsMissing(errorMsg) Then
    errorMsg = keyword & " could not be found when importing."
End If

' Store address of various things
Dim myAddress As Variant

' Store success status of import
Dim succeeded As Boolean
succeeded = True

'Find total sales
Dim value As String ' Store target cell's value
myAddress = funcs.search(keyword, importSheet)
If myAddress(0) = "0" Then
    value = substitute
    MsgBox (errorMsg)
    succeeded = False
Else
    value = Worksheets(importSheet).Cells(CInt(myAddress(0) + offset(0)), CInt(myAddress(1) + offset(1)))
End If

' Update cell with target cell value, if found, or value of 'substitue' variable if not
Worksheets(exportSheet).Cells(exportCell(0), exportCell(1)) = value

ImportCell = succeeded
End Function

Private Sub ToggleEditMode(state As Boolean)
If state = True Then
    ' Update captions
    EventButton1.Caption = "Edit Selected Event"
    EventButton2.Caption = "Edit Selected Event"
    EventButton3.Caption = "Edit Selected Event"
    
    ' Tick all of the other edit checkboxes
    EditToggleCheckBox1.value = True
    EditToggleCheckBox2.value = True
    EditToggleCheckBox3.value = True
ElseIf state = False Then
    ' Update captions
    EventButton1.Caption = "Add New Event"
    EventButton2.Caption = "Add New Event"
    EventButton3.Caption = "Add New Event"
    
    ' Untick all of the other edit checkboxes
    EditToggleCheckBox1.value = False
    EditToggleCheckBox2.value = False
    EditToggleCheckBox3.value = False
End If
End Sub

Private Sub ImportFromTicketsolve(mode As String)
' Highly non-general function/sub

' Store which row we're working on. Depends on what we're after.
Dim my_row As Long

' Not strictly necessary. Used to avoid too much hard-coding.
Dim dataSheetName As String
dataSheetName = "Data"

' Decide which event to import info into
If mode = "Selected" Then
    my_row = funcs.search(SearchBox.Text, dataSheetName)(0)
ElseIf mode = "Previous" Then
    my_row = Worksheets(dataSheetName).Cells(Rows.Count, 1).End(xlUp).Row
Else
    MsgBox ("Programmer error in Private Sub TransferInfo")
    Exit Sub
End If

' Check if we were selling tickets. If not, prompt the user to update
'   the event before importing
If Worksheets(dataSheetName).Cells(my_row, 46) = "False" Then
    MsgBox ("We are not selling tickets for this event. If you wish to import " _
            & "data from Ticketsolve, please mark the event as 'Ticketed'.")
    Exit Sub
End If

' Not strictly necessary. Used to avoid too much hard-coding.
Dim sheetName As String
sheetName = "Import"

Dim csvImportSuccessCheck As Boolean
' Import the csv selected by the user into sheet "sheetName"
csvImportSuccessCheck = funcs.csv_Import(sheetName)
If csvImportSuccessCheck = True Then
    ' Continue with sub
Else
    ' Exit sub, since a file wasn't selected
    Exit Sub
End If

' Store whether import is successful or not.
Dim succeeded As Boolean
succeeded = True

'Find total number of ticket sales
Dim tempSuccessCheck As Boolean
Dim exportCell() As Variant
Dim offset() As Variant

exportCell = Array(my_row, 14)
offset = Array(0, 1)

tempSuccessCheck = ImportCell("Sold", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find event capacity
exportCell(1) = 15
tempSuccessCheck = ImportCell("Capacity", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find number of blocked seats
exportCell(1) = 34
offset = Array(1, 0)
tempSuccessCheck = ImportCell("Blocked", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find total ticket revenue
exportCell(1) = 42
offset = Array(1, 5)
tempSuccessCheck = ImportCell("Allocation Type", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find total Support the Kirkgate revenue
exportCell(1) = 44
offset = Array(0, 3)
tempSuccessCheck = ImportCell("Support the Kirkgate", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Must go at the bottom
If succeeded = True Then
    MsgBox ("Import successful")
End If

' Determine actual capacity and write it to a cell
Dim trueCapacity As Integer
trueCapacity = Worksheets(dataSheetName).Cells(exportCell(0), 15) - Worksheets(dataSheetName).Cells(exportCell(0), 34)
Worksheets(dataSheetName).Cells(exportCell(0), 45) = trueCapacity
End Sub

Public Function Foo(Bar As String) As String

End Function
