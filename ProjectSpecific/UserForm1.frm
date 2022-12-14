VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Events"
   ClientHeight    =   10770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14595
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Dumb, but why not
Public EventDeleteIndicator As Integer
Public GroupDeleteIndicator As Integer

' Store whether editing is happening
Public EditingCheck As Boolean

' Store whether autofilling is happening
Public AutofillCheck As Boolean

' Used to prevent unecessary refreshing of ListBoxes
Public counter As Integer

' Store user input so that it can be restored if they make a mistake
' Capacity
Public AuditoriumCapacity As String
Public EgremontCapacity As String
Public TotalCapacity As String
Public BlockedSeats As String

' Revenue & Costs
Public NumTicketsSold As String
Public BoxOfficeRevenue As String
Public SupportRevenue As String ' Revenue from Support the Kirkgate donations
Public RoomHireRevenue As String
Public MiscRevenue As String
' Bar
Public BarRevenue As String
Public BarMargin As String
' Film
Public FilmCost As String
Public FilmTransport As String
' Other
Public Accommodation As String
Public ArtistFood As String
Public HiredPersonnel As String
Public Heating As String
Public Lighting As String
Public MiscCost As String

' Time
' Event Time
Public SetupStartTime As String
Public DoorsTime As String
Public EventStartTime As String
Public EventEndTime As String
Public TakedownEndTime As String
Public EventDuration As String
Public SetupToTakedownEndDuration As String
Public SetupAvailableDuration As String
Public SetupTakedown As String
'Bar Time
Public BarSetupTime As String
Public BarOpenTime As String
Public BarCloseTime As String
Public BarOpenDuration As String
Public BarSetupToTakedownEndDuration As String
Public BarSetupTakedown As String

' Volunteers
' Volunteer Minutes Worked
Public FoH As String
Public DM As String
Public Tech As String
Public BoxOffice As String
Public Bar As String
Public AoWVol As String
Public MiscVol As String
' Volunteer Nominal Pay
Public FoHPay As String
Public DMPay As String
Public TechPay As String
Public BoxOfficePay As String
Public BarPay As String
Public AoWVolPay As String
Public MiscVolPay As String

' Group management
Public GroupName As String
Public EditGroupStatus As Boolean

Private Sub AutofillCheckBox_Click()
If AutofillCheckBox.value = True Then
    ' Toggle edit mode
    EditToggleCheckBox1.value = True
Else
    ' Toggle edit mode
    EditToggleCheckBox1.value = False
End If

If EventIDListBox.ListIndex = -1 Then
    ' No event has been selected, so do nothing
    Exit Sub
End If

' Store row we're autofilling into
Dim nameLocation As Integer
nameLocation = EventIDListBox.ListIndex + 2

If AutofillCheckBox.value = True Then
    ' Autofill data into form
    Call AutofillEventFromSelected(nameLocation)
End If

' Toggle edit mode

End Sub

Private Sub BarOpenOptionButton_Change()
' Hide the bar stuff if it isn't needed.
If BarOpenOptionButton = True Then
    BarTimeFrame.Visible = True
Else
    BarTimeFrame.Visible = False
End If
End Sub

Private Sub CreditsButton_Click()
CreditsForm.Show
End Sub

'' BUTTON CLICKING===============================================================

Private Sub EventButton1_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton1_DblClick(ByVal cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_DblClick(ByVal cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton3_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton3_DblClick(ByVal cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_DblClick(ByVal cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_DblClick(ByVal cancel As MSForms.ReturnBoolean)
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

' Store index of row to be deleted
Dim row As Integer
row = EventIDListBox.ListIndex + 2

Dim my_index As Long
my_index = EventIDListBox.ListIndex

' Mess with listindex to prevent breakage
If EventIDListBox.ListCount < 1 Then
    ' There are no more items, so select nothing
    EventIDListBox.ListIndex = -1
    EventDeleteIndicator = 1
    ' Delete entire row corresponding to selected event
    Sheets("Data").Rows(row).Delete
ElseIf EventIDListBox.ListCount = my_index + 1 Then
    ' We are at end of list, so go up one
    EventIDListBox.ListIndex = my_index - 1
    EventDeleteIndicator = 1
    ' Delete entire row corresponding to selected event
    Sheets("Data").Rows(row).Delete
Else
    ' Delete entire row corresponding to selected event
    EventDeleteIndicator = 1
    Sheets("Data").Rows(row).Delete
End If

'' Update listboxes
' Update EventIDListBox
Call funcs.RefreshListBox("Data", 1, EventIDListBox)
' Update search listboxes
' HiddenEventIDListBox must be updated first so that the internal record
'   of events is correct. If this doesn't make sense, dw, I got confused while
'   writing this comment. It needs to go first though. Try it without and see
'   what you get.
HiddenEventIDListBox.RemoveItem (SearchNameListBox.ListIndex)
SearchNameListBox.RemoveItem (SearchNameListBox.ListIndex)
SearchDateListBox.RemoveItem (SearchNameListBox.ListIndex)
SearchTypeListBox.RemoveItem (SearchNameListBox.ListIndex)

' Update pivot table(s)
Call funcs.ChangeSource("Data", "Analysis", "PivotTable1")
End Sub

Private Sub GenreListBox_Change()
' Stop it from getting stroppy for comparing a string to null
If GenreListBox.ListIndex = -1 Then
    Exit Sub
End If

If GenreListBox.value = "Jazz" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 10, "Jazz")
End If
End Sub

Private Sub GroupButton1_Click()
Dim my_row As Integer

' Pre-fill in details to make adding groups quicker
If GroupNameListBox.ListIndex = -1 Then ' Nothing is selected, so use event info
    GroupManagementForm.GroupNameTextBox.Text = UserForm1.GroupSearchBox.Text
    GroupManagementForm.StartDateTextBox.Text = UserForm1.StartDateTextBox.Text
    GroupManagementForm.EndDateTextBox.Text = UserForm1.EndDateTextBox.Text
    GroupManagementForm.CategoryListBox.value = UserForm1.CategoryListBox.value
    GroupManagementForm.TypeListBox.value = UserForm1.TypeListBox.value
Else
    ' find row group is on
    my_row = funcs.search(GroupNameListBox.value, "Data")(0)
    ' Use info from group
    Call AutofillGroupFromSelected(my_row)
End If

GroupManagementForm.EditToggleCheckBox1.value = GroupEditToggleCheckBox.value

' This needs to be last in the sub so that the other code executes
GroupManagementForm.Show
End Sub

Private Sub GroupEditToggleCheckBox_Click()
If GroupEditToggleCheckBox.value = True Then
    GroupButton1.Caption = "Edit Group"
Else
    GroupButton1.Caption = "Add New Group"
End If
End Sub

Private Sub TicketsolveImportSelectedButton_Click()
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
Call Import.ImportFromTicketsolve("Selected")
End Sub

Private Sub TicketsolveImportPreviousButton_Click()
' Import data into the event which was most recently added
Call Import.ImportFromTicketsolve("Previous")
End Sub

Private Sub ZettleImportSelectedButton_Click()
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

' Import info from Zettle
Call Import.ImportFromZettle("Selected")
End Sub
Private Sub ZettleImportPreviousButton_Click()
' Import info from Zettle
Call Import.ImportFromZettle("Previous")
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

Private Sub ClearTimeButton_Click()
' Clear all input from Time tab
' Event
SetupStartTimeTextBox.Text = ""
DoorsTimeTextBox.Text = ""
EventStartTimeTextBox.Text = ""
EventEndTimeTextBox.Text = ""
TakedownEndTimeTextBox.Text = ""
EventDurationTextBox.Text = ""
SetupToTakedownEndDurationTextBox.Text = ""
SetupTakedownTextBox.Text = ""
' Bar
BarSetupTimeTextBox.Text = ""
BarOpenTimeTextBox.Text = ""
BarCloseTimeTextBox.Text = ""
BarOpenDurationTextBox.Text = ""
BarSetupToTakedownEndDurationTextBox.Text = ""
BarSetupTakedownTextBox.Text = ""
End Sub

Private Sub VolRef1Button_Click()
If UpdateVolunteerMinutes = False Then
    MsgBox ("Please fill all time and duration boxes in on this page before refreshing")
End If
End Sub

Private Sub VolRef2Button_Click()
If UpdateVolunteerMinutes = False Then
    MsgBox ("Please fill all time and duration boxes in on the time page before refreshing")
End If
End Sub

'' TEXT BOXES===============================================================

' Basic Info============================================================================
Private Sub StartDateTextBox_DblClick(ByVal cancel As MSForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call funcs.GetCalendar(UserForm1.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub StartDateTextBox_Enter()
Call funcs.GetCalendar(UserForm1.StartDateTextBox) ' Show Date Picker
End Sub

Private Sub StartDateTextBox_change()
' Update end date to match start date
EndDateTextBox.Text = StartDateTextBox.Text
' Update times if date is changed and type matches what we want
If TypeListBox.value = "Film" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 2, "Film")
ElseIf TypeListBox.value = "Live Music" And GenreListBox.value = "Jazz" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 10, "Jazz")
End If
End Sub

Private Sub EndDateTextBox_Dblclick(ByVal cancel As MSForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call funcs.GetCalendar(UserForm1.EndDateTextBox) ' Show Date Picker
End Sub

Private Sub EndDateTextBox_Enter()
Call funcs.GetCalendar(UserForm1.EndDateTextBox) ' Show Date Picker
End Sub

' Layout & Capacity============================================================================
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

Private Sub BlockedSeatsTextBox_Change()
' Sanitise input to ensure only real numbers <= 100 are input
Call InptValid.SanitiseNonNegInt(BlockedSeatsTextBox, BlockedSeats)
End Sub

Private Sub BlockedSeatsTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Stop it from being left blank
If BlockedSeatsTextBox.Text = "" Then
    BlockedSeatsTextBox.Text = 0
End If
End Sub

' Time============================================================================

Private Sub SetupStartTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(SetupStartTimeTextBox, SetupStartTime)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

' Make rest of sub easier to read
Dim sheet As String
sheet = "Type-Specific Defaults"

' Find out which row we are getting our defaults from
Dim row As Integer
row = TypeListBox.ListIndex + 2

' Prevent crashes
If row = 1 Or SetupStartTimeTextBox.Text = "" Or TypeListBox.ListIndex = -1 Then
    Exit Sub
End If

' Change doors open time
If DoorsTimeTextBox.Text = "" Then
    DoorsTimeTextBox.Text = Format(DateAdd("n", (Worksheets(sheet).Cells(row, 9) - _
        Worksheets(sheet).Cells(row, 11)), SetupStartTimeTextBox.value), "hh:mm")
End If

' Change event start time
If EventStartTimeTextBox.Text = "" Then
    EventStartTimeTextBox.Text = Format(DateAdd("n", Worksheets(sheet).Cells(row, 9), _
        SetupStartTimeTextBox.value), "hh:mm")
End If

' Change event end time
If EventEndTimeTextBox.Text = "" Then
    EventEndTimeTextBox.Text = Format(DateAdd("n", (Worksheets(sheet).Cells(row, 9) + _
        Worksheets(sheet).Cells(row, 10)), SetupStartTimeTextBox.value), "hh:mm")
End If

' Change takedown end time
If TakedownEndTimeTextBox.Text = "" Then
    TakedownEndTimeTextBox.Text = Format(DateAdd("n", Worksheets(sheet).Cells(row, 13), _
        SetupStartTimeTextBox.value), "hh:mm")
End If

' Change setup to takedown complete duration
If TakedownEndTimeTextBox.Text = "" Then
ElseIf CDate(TakedownEndTimeTextBox.Text) - _
        CDate(SetupStartTimeTextBox.Text) < 0 Then
    MsgBox ("The setup cannot begin after the takedown.")
    SetupStartTimeTextBox.Text = ""
Else
    SetupToTakedownEndDurationTextBox.Text = (CDate(TakedownEndTimeTextBox.Text) - _
        CDate(SetupStartTimeTextBox.Text)) * 24 * 60
End If
End Sub

Private Sub DoorsTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(DoorsTimeTextBox, DoorsTime)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

If DoorsTimeTextBox.Text = "" Or EventStartTimeTextBox.Text = "" Then
    Exit Sub
' Sanity check doors time
ElseIf TimeValue(EventStartTimeTextBox.Text) < TimeValue(DoorsTimeTextBox.Text) Then
    MsgBox ("Doors cannot open after event starts")
    DoorsTimeTextBox.Text = ""
End If

End Sub

Private Sub EventStartTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(EventStartTimeTextBox, EventStartTime)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

' Make rest of sub easier to read
Dim sheet As String
sheet = "Type-Specific Defaults"

' Find out which row we are getting our defaults from
Dim row As Integer
row = TypeListBox.ListIndex + 2

' Prevent crashes
If row = 1 Or EventStartTimeTextBox.Text = "" Or TypeListBox.ListIndex = -1 Then
    Exit Sub
End If

' Change setup start time
SetupStartTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 9), _
    EventStartTimeTextBox.value), "hh:mm")

' Change doors open time
DoorsTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 11), _
    EventStartTimeTextBox.Text), "hh:mm")

' Change event end time
If EventDurationTextBox.Text <> "" Then
    EventEndTimeTextBox.Text = Format(DateAdd("n", EventDurationTextBox.Text, _
        EventStartTimeTextBox.Text), "hh:mm")
End If

' Change takedown time
If EventDurationTextBox.Text <> "" Then
    TakedownEndTimeTextBox.Text = Format(DateAdd("n", EventDurationTextBox.Text + _
        Worksheets(sheet).Cells(row, 12), EventStartTimeTextBox.Text), "hh:mm")
End If

' BAR
' Change bar setup start time
BarSetupTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 11) - _
    Worksheets(sheet).Cells(row, 16), EventStartTimeTextBox.Text), "hh:mm")

' Change bar open time: event start time - doors open time
BarOpenTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 11), _
    EventStartTimeTextBox.Text), "hh:mm")
    
' Change bar close time: bar open time + bar open duration
BarCloseTimeTextBox.Text = Format(DateAdd("n", Worksheets(sheet).Cells(row, 15), _
                            BarOpenTimeTextBox.Text), "hh:mm")

' Update volunteer hours
Call UpdateVolunteerMinutes
End Sub

Private Sub EventEndTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(EventEndTimeTextBox, EventEndTime)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

' Make rest of sub easier to read
Dim sheet As String
sheet = "Type-Specific Defaults"

' Find out which row we are getting our defaults from
Dim row As Integer
row = TypeListBox.ListIndex + 2

' Prevent crashes
If row = 1 Or EventEndTimeTextBox.Text = "" Or TypeListBox.ListIndex = -1 Then
    Exit Sub
End If

' Change takedown time
TakedownEndTimeTextBox.Text = Format(DateAdd("n", Worksheets(sheet).Cells(row, 12), _
    EventEndTimeTextBox.value), "hh:mm")

' Change event duration
If EventStartTimeTextBox.Text <> "" Then
    EventDurationTextBox.Text = (CDate(EventEndTimeTextBox.Text) - _
        CDate(EventStartTimeTextBox.Text)) * 24 * 60
End If

' Change setup to takedown complete duration
If SetupStartTimeTextBox.Text <> "" And TakedownEndTimeTextBox.Text <> "" Then
    SetupToTakedownEndDurationTextBox.Text = (CDate(TakedownEndTimeTextBox.Text) - _
        CDate(SetupStartTimeTextBox.Text)) * 24 * 60
End If
End Sub

Private Sub TakedownEndTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(TakedownEndTimeTextBox, TakedownEndTime)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

' No end time
If TakedownEndTimeTextBox.Text = "" Then
    Exit Sub
' Both start and end time
ElseIf TakedownEndTimeTextBox.Text <> "" And SetupStartTimeTextBox.Text <> "" Then
    SetupToTakedownEndDurationTextBox.Text = (CDate(TakedownEndTimeTextBox.Text) - _
        CDate(SetupStartTimeTextBox.Text)) * 24 * 60
Else ' End time only
    
End If
End Sub

Private Sub EventDurationTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(EventDurationTextBox, EventDuration)

' Don't automatically adjust unless wanted
If AutoTimeCheckBox = False Then
    Exit Sub
End If

' Make rest of sub easier to read
Dim sheet As String
sheet = "Type-Specific Defaults"

' Find out which row we are getting our defaults from
Dim row As Integer
row = TypeListBox.ListIndex + 2

Dim takedownDuration As Integer

' Event duration is blank, so exit sub
If EventDurationTextBox.Text = "" Then
    Exit Sub
' No start or end time
ElseIf EventStartTimeTextBox.Text = "" And EventEndTimeTextBox.Text = "" Then
    ' Don't complete this if statement
    
' No start time, but there is an end time
ElseIf EventStartTimeTextBox.Text = "" Then
    ' Set start time according to end time
    EventStartTimeTextBox.Text = Format(DateAdd("n", -EventDurationTextBox.Text, _
        EventEndTimeTextBox.Text), "hh:mm")
Else ' Is a start time, may or may not be an end time. Will be by the end
    
    ' Find takedown duration
    If EventEndTimeTextBox.Text = "" Then ' No end time, so use default
        takedownDuration = Worksheets(sheet).Cells(row, 12)
    Else
        ' Use start and end time to find takedown duration
        takedownDuration = (CDate(TakedownEndTimeTextBox.Text) - _
            CDate(EventEndTimeTextBox.Text)) * 24 * 60
    End If
    
    ' Set end time according to start time
    EventEndTimeTextBox.Text = Format(DateAdd("n", EventDurationTextBox.Text, _
        EventStartTimeTextBox.Text), "hh:mm")
        
    ' Set takedown end time according to end time and takedown duration
    TakedownEndTimeTextBox.Text = Format(DateAdd("n", takedownDuration, _
        EventEndTimeTextBox.Text), "hh:mm")
End If

' Can get here without a start or end time

' Now we want to know the setup to takedown complete duration

Dim setupDuration As Integer

' Change setup to takedown end duration
If SetupStartTimeTextBox.Text <> "" And TakedownEndTimeTextBox.Text <> "" Then
    SetupToTakedownEndDurationTextBox.Text = (CDate(TakedownEndTimeTextBox.Text) - _
            CDate(SetupStartTimeTextBox.Text)) * 24 * 60
    ' We've found ideal setup to takedown end duration, so no need to continue
    Exit Sub
End If

' Find setup duration
If SetupStartTimeTextBox.Text <> "" And EventStartTimeTextBox.Text <> "" Then
    setupDuration = (CDate(EventStartTimeTextBox.Text) - _
        CDate(SetupStartTimeTextBox.Text)) * 24 * 60
Else
    setupDuration = Worksheets(sheet).Cells(row, 9)
End If

' Find takedown duration
If EventEndTimeTextBox.Text <> "" And TakedownEndTimeTextBox.Text <> "" Then
    takedownDuration = (CDate(TakedownEndTimeTextBox.Text) - _
            CDate(EventEndTimeTextBox.Text)) * 24 * 60
Else
    takedownDuration = Worksheets(sheet).Cells(row, 12)
End If

SetupToTakedownEndDurationTextBox.Text = setupDuration + _
    EventDurationTextBox.Text + takedownDuration

End Sub

Private Sub SetupToTakedownEndDurationTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(SetupToTakedownEndDurationTextBox, SetupToTakedownEndDuration)
End Sub

Private Sub SetupAvailableDurationTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(SetupAvailableDurationTextBox, SetupAvailableDuration)
End Sub

Private Sub SetupTakedownTextBox_Change()
' This box shouldn't affect anything else, but should be affected
'   by other timings in a minor way.

' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(SetupTakedownTextBox, SetupTakedown)
End Sub

Private Sub BarSetupTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarSetupTimeTextBox, BarSetupTime)
End Sub

Private Sub BarOpenTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarOpenTimeTextBox, BarOpenTime)
End Sub

Private Sub BarCloseTimeTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarCloseTimeTextBox, BarCloseTime)
End Sub

Private Sub BarOpenDurationTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BarOpenDurationTextBox, BarOpenDuration)
End Sub

Private Sub BarSetupToTakedownEndDurationTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BarSetupToTakedownEndDurationTextBox, BarSetupToTakedownEndDuration)
End Sub

Private Sub BarSetupTakedownTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BarSetupTakedownTextBox, BarSetupTakedown)
End Sub

' Costs & Income============================================================================
Private Sub NumTicketsSoldTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(NumTicketsSoldTextBox, NumTicketsSold)
End Sub

Private Sub BoxOfficeRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BoxOfficeRevenueTextBox, BoxOfficeRevenue)
End Sub

Private Sub BoxOfficeRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
BoxOfficeRevenueTextBox.Text = StrManip.Convert2Currency(BoxOfficeRevenueTextBox)
End Sub

Private Sub SupportRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(SupportRevenueTextBox, SupportRevenue)
End Sub

Private Sub SupportRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
SupportRevenueTextBox.Text = StrManip.Convert2Currency(SupportRevenueTextBox)
End Sub

Private Sub RoomHireRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(RoomHireRevenueTextBox, RoomHireRevenue)
End Sub

Private Sub RoomHireRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
RoomHireRevenueTextBox.Text = StrManip.Convert2Currency(RoomHireRevenueTextBox)
End Sub

Private Sub MiscRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(MiscRevenueTextBox, MiscRevenue)
End Sub

Private Sub MiscRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
MiscRevenueTextBox.Text = StrManip.Convert2Currency(MiscRevenueTextBox)
End Sub

Private Sub BarRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BarRevenueTextBox, BarRevenue)
End Sub

Private Sub BarRevenueTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
BarRevenueTextBox.Text = StrManip.Convert2Currency(BarRevenueTextBox)
End Sub

Private Sub BarMarginTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitisePercentage(BarMarginTextBox, BarMargin)
End Sub

Private Sub FilmCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmCostTextBox, FilmCost)
End Sub

Private Sub FilmCostTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
FilmCostTextBox.Text = StrManip.Convert2Currency(FilmCostTextBox)
End Sub

Private Sub FilmTransportTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmTransportTextBox, FilmTransport)
End Sub

Private Sub FilmTransportTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
FilmTransportTextBox.Text = StrManip.Convert2Currency(FilmTransportTextBox)
End Sub

Private Sub AccommodationTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(AccommodationTextBox, Accommodation)
End Sub

Private Sub AccommodationTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
AccommodationTextBox.Text = StrManip.Convert2Currency(AccommodationTextBox)
End Sub

Private Sub ArtistFoodTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(ArtistFoodTextBox, ArtistFood)
End Sub

Private Sub ArtistFoodTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
ArtistFoodTextBox.Text = StrManip.Convert2Currency(ArtistFoodTextBox)
End Sub

Private Sub HeatingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HeatingTextBox, Heating)
End Sub

Private Sub HeatingTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
HeatingTextBox.Text = StrManip.Convert2Currency(HeatingTextBox)
End Sub

Private Sub LightingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(LightingTextBox, Lighting)
End Sub

Private Sub LightingTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
LightingTextBox.Text = StrManip.Convert2Currency(LightingTextBox)
End Sub

Private Sub MiscCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(MiscCostTextBox, MiscCost)
End Sub

Private Sub MiscCostTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
MiscCostTextBox.Text = StrManip.Convert2Currency(MiscCostTextBox)
End Sub

Private Sub HiredPersonnelTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HiredPersonnelTextBox, HiredPersonnel)
End Sub

Private Sub HiredPersonnelTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
HiredPersonnelTextBox.Text = StrManip.Convert2Currency(HiredPersonnelTextBox)
End Sub

' Volunteers==================================================================
' Volunteer Minutes Worked
Private Sub FoHTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(FoHTextBox, FoH)

' Set nominal pay based on this
If FoHTextBox.Text <> "" Then
    FoHPayTextBox.Text = Format( _
            FoHTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 14) / 60, _
            "Currency")
Else
    FoHPayTextBox.Text = "0.00"
End If
End Sub

Private Sub DMTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(DMTextBox, DM)

' Set nominal pay based on this
If DMTextBox.Text <> "" Then
    DMPayTextBox.Text = Format( _
            DMTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 15) / 60, _
            "Currency")
Else
    DMPayTextBox.Text = "0.00"
End If
End Sub

Private Sub TechTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(TechTextBox, Tech)

' Set nominal pay based on this
If TechTextBox.Text <> "" Then
    TechPayTextBox.Text = Format( _
            TechTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 16) / 60, _
            "Currency")
Else
    TechPayTextBox.Text = "0.00"
End If
End Sub

Private Sub BoxOfficeTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BoxOfficeTextBox, BoxOffice)

' Set nominal pay based on this
If BoxOfficeTextBox.Text <> "" Then
    BoxOfficePayTextBox.Text = Format( _
            BoxOfficeTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 14) / 60, _
            "Currency")
Else
    BoxOfficePayTextBox.Text = "0.00"
End If
End Sub

Private Sub BarTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BarTextBox, Bar)

' Set nominal pay based on this
If BarTextBox.Text <> "" Then
    BarPayTextBox.Text = Format( _
            BarTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 14) / 60, _
            "Currency")
Else
    BarPayTextBox.Text = "0.00"
End If
End Sub

Private Sub AoWVolTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(AoWVolTextBox, AoWVol)

' Set nominal pay based on this
If AoWVolTextBox.Text <> "" Then
    AoWVolPayTextBox.Text = Format( _
            AoWVolTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 14) / 60, _
            "Currency")
Else
    AoWVolPayTextBox.Text = "0.00"
End If
End Sub

Private Sub MiscVolTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(MiscVolTextBox, MiscVol)

' Set nominal pay based on this
If MiscVolTextBox.Text <> "" Then
    MiscVolPayTextBox.Text = Format( _
            MiscVolTextBox.Text * Worksheets("Non-Specific Defaults").Cells(2, 14) / 60, _
            "Currency")
Else
    MiscVolPayTextBox.Text = "0.00"
End If
End Sub

'Volunteer Nominal Pay
Private Sub FoHPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(FoHPayTextBox, FoHPay)
End Sub

Private Sub FoHPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
FoHPayTextBox.Text = StrManip.Convert2Currency(FoHPayTextBox)
End Sub

Private Sub DMPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(DMPayTextBox, DMPay)
End Sub

Private Sub DMPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
DMPayTextBox.Text = StrManip.Convert2Currency(DMPayTextBox)
End Sub

Private Sub TechPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(TechPayTextBox, TechPay)
End Sub

Private Sub TechPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
TechPayTextBox.Text = StrManip.Convert2Currency(TechPayTextBox)
End Sub

Private Sub BoxOfficePayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(BoxOfficePayTextBox, BoxOfficePay)
End Sub

Private Sub BoxOfficePayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
BoxOfficePayTextBox.Text = StrManip.Convert2Currency(BoxOfficePayTextBox)
End Sub

Private Sub BarPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(BarPayTextBox, BarPay)
End Sub

Private Sub BarPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
BarPayTextBox.Text = StrManip.Convert2Currency(BarPayTextBox)
End Sub

Private Sub AoWVolPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(AoWVolPayTextBox, AoWVolPay)
End Sub

Private Sub AoWVolPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
AoWVolPayTextBox.Text = StrManip.Convert2Currency(AoWVolPayTextBox)
End Sub

Private Sub MiscVolPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(MiscVolPayTextBox, MiscVolPay)
End Sub

Private Sub MiscVolPayTextBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
MiscVolPayTextBox.Text = StrManip.Convert2Currency(MiscVolPayTextBox)
End Sub

Private Sub NewSearchTextBox_Change()
Dim non_empty_row As Long
Dim DataRange As Range

' non_empty_row = lst non-empty row for specific list(box)
non_empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row
Set DataRange = Range(Worksheets("Data").Cells(2, 2), _
                Worksheets("Data").Cells(non_empty_row, 2))

' Clear items to avoid them being re-added
SearchNameListBox.Clear
SearchDateListBox.Clear
SearchTypeListBox.Clear
HiddenEventIDListBox.Clear

' Use "Union(Range1, Range2)" to combine ranges

' Search for events and add them to the listbox
Call funcs.AddSomeToListBox(NewSearchTextBox.Text, DataRange, Array(2, 3, 29, 1), SearchNameListBox, _
                           SearchDateListBox, SearchTypeListBox, HiddenEventIDListBox, "Data", 73, _
                           True)
End Sub

Private Sub GroupSearchBox_Change()
Dim non_empty_row As Long
Dim DataRange As Range

' non_empty_row = lst non-empty row for specific list(box)
non_empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row
Set DataRange = Range(Worksheets("Data").Cells(2, 2), _
                Worksheets("Data").Cells(non_empty_row, 2))

' Clear items to avoid them being re-added
GroupNameListBox.Clear
StartDateListBox.Clear
EndDateListBox.Clear
GroupTypeListBox.Clear
HiddenGroupIDListBox.Clear

' Use "Union(Range1, Range2)" to combine ranges

' Search for events and add them to the listbox
Call funcs.AddSomeToListBox(GroupSearchBox.Text, DataRange, Array(2, 3, 29, 72, 74), _
                            GroupNameListBox, StartDateListBox, GroupTypeListBox, _
                            HiddenGroupIDListBox, "Data", 73, False, EndDateListBox)
End Sub

'' LIST BOXES===============================================================

Private Sub TypeListBox_Change()
Dim TypeDefaultsSheet As String
TypeDefaultsSheet = "Type-Specific Defaults"

' What row we're looking at
Dim row As Integer
row = TypeListBox.ListIndex + 2

' Show the time page
Me.MultiPage1.Pages("Page4").Visible = True

' Change some timing boxes
' Routine is different if type is "Film"
If TypeListBox.value = "Film" And StartDateTextBox.Text <> "" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 2, "Film")
ElseIf TypeListBox.value = "Live Music" And StartDateTextBox.Text <> "" _
                            And GenreListBox.value = "Jazz" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 10, "Jazz")
Else
    ' Event
    EventDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 10)
    SetupToTakedownEndDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 13)
    SetupTakedownTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 9) + _
                                Worksheets(TypeDefaultsSheet).Cells(row, 12)
    ' Bar
    BarOpenDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 15)
    BarSetupToTakedownEndDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 15) + _
        Worksheets(TypeDefaultsSheet).Cells(row, 16)
    BarSetupTakedownTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 16)
End If

' Show the genres if applicable
If TypeListBox.value = "Live Music" Then
    GenreListBox.Visible = True
    GenreLabel.Visible = True
Else
    GenreListBox.Visible = False
    GenreLabel.Visible = False
End If

' Set volunteer minutes to defaults found in Type-Specific Defaults
FoHTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 3)
DMTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 4)
TechTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 5)
BarTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 7)
AoWVolTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 8)
MiscVolTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 14)
End Sub

Private Sub LocationListBox_Change()
' Don't do this validation if user is editing
If AutofillCheck = True Then
    ' Do nothing
' Throw an error if the location doesn't match the room selected (external should be selected for anything not in Kirkgate)
ElseIf LocationListBox.value <> "Kirkgate" And RoomListBox.value <> "External" And RoomListBox.ListIndex <> -1 Then
    MsgBox ("'External' should refer to an Arts out West venue.")
ElseIf LocationListBox.value = "Kirkgate" And RoomListBox.value = "External" Then
    MsgBox ("'External' should refer to an Arts out West venue.")
End If
End Sub

Private Sub RoomListBox_Change()
' Don't do this validation if user is editing
If AutofillCheck = True Then
ElseIf LocationListBox.value <> "Kirkgate" And RoomListBox.value <> "External" And LocationListBox.ListIndex <> -1 Then
    MsgBox ("'External' should refer to an Arts out West venue.")
ElseIf LocationListBox.value = "Kirkgate" And RoomListBox.value = "External" Then
    MsgBox ("'External' should refer to an Arts out West venue.")
End If

' Decide what to show regarding capacity options
CapacityListBoxDecider
End Sub

Private Sub AuditoriumLayoutListBox_Change()
AuditoriumCapacityTextBox.Text = Worksheets("Non-Specific Defaults").Cells(AuditoriumLayoutListBox.ListIndex + 2, 7)
End Sub

Private Sub EgremontLayoutListBox_Change()
EgremontCapacityTextBox.Text = Worksheets("Non-Specific Defaults").Cells(EgremontLayoutListBox.ListIndex + 2, 9)
End Sub

Private Sub AudienceListBox_Change()
' Decide what to show regarding capacity options
Call CapacityListBoxDecider
End Sub

Private Sub EventIDListBox_Change()
' Stop it freaking out when an item is deleted
If EventDeleteIndicator = 1 Then
    EventDeleteIndicator = 2
    Exit Sub
ElseIf EventDeleteIndicator = 2 Then
    ' Reset EventDeleteIndicator on second time of asking
    EventDeleteIndicator = 0
    Exit Sub
Else
    ' Carry on with the sub
End If

' start at 0
' increase by 1 when deletion called
' now 1
' increase by 1 one first change
' now 2
' reset to 0, then exit sub

' Store name of data sheet
Dim sheet As String
sheet = "Data"

' Store row of selected event.
Dim row As Long
row = EventIDListBox.ListIndex + 2

' So the user always knows which event has been selected.
' Might want to hide that if they're not in edit mode, idk.
EventIDUpdaterLabel1.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel2.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel3.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel4.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel5.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel6.Caption = "Selected Event ID: " & EventIDListBox.value
'EventIDUpdaterLabel7.Caption = "Selected Event ID: " & EventIDListBox.value

If AutofillCheckBox.value = True Then
    ' Fill in everything with values taken from this event.
    Call AutofillEventFromSelected(row)
End If
End Sub

Private Sub GroupIdListbox_Change()
' Stop it freaking out when an item is deleted
If GroupDeleteIndicator = 1 Then
    GroupDeleteIndicator = 2
    Exit Sub
ElseIf GroupDeleteIndicator = 2 Then
    ' Reset GroupDeleteIndicator on second time of asking
    GroupDeleteIndicator = 0
    Exit Sub
Else
    ' Carry on with the sub
End If

' start at 0
' increase by 1 when deletion called
' now 1
' increase by 1 one first change
' now 2
' reset to 0, then exit sub

GroupIDUpdater1.Caption = "Selected Group ID: " & GroupIDListBox.value
GroupManagementForm.GroupIDUpdaterLabel2.Caption = "Selected Group ID: " & GroupIDListBox.value

End Sub

'' Search for events tab
Private Sub SearchNameListBox_Click()
' Keep other listboxes in lockstep
SearchDateListBox.ListIndex = SearchNameListBox.ListIndex
SearchTypeListBox.ListIndex = SearchNameListBox.ListIndex
End Sub

Private Sub SearchDateListBox_Click()
' These click subs act like change subs, so calling one
'   calls the others because calling one changes the others.

' Keep other listboxes in lockstep
SearchNameListBox.ListIndex = SearchDateListBox.ListIndex
SearchTypeListBox.ListIndex = SearchDateListBox.ListIndex
HiddenEventIDListBox.ListIndex = SearchDateListBox.ListIndex

' Select event in EventIDListBox so that we know which row it is on
EventIDListBox.value = HiddenEventIDListBox.value
End Sub

Private Sub SearchTypeListBox_Click()
' Keep other listboxes in lockstep
SearchNameListBox.ListIndex = SearchTypeListBox.ListIndex
SearchDateListBox.ListIndex = SearchTypeListBox.ListIndex
End Sub

'' Manage groups tab
Private Sub GroupNameListBox_Click()
' Keep other listboxes in lockstep
StartDateListBox.ListIndex = GroupNameListBox.ListIndex
EndDateListBox.ListIndex = GroupNameListBox.ListIndex
GroupTypeListBox.ListIndex = GroupNameListBox.ListIndex
HiddenGroupIDListBox.ListIndex = GroupNameListBox.ListIndex

' Select event in GroupIDListBox so that we know which row it is on
GroupIDListBox.value = HiddenGroupIDListBox.value
End Sub

Private Sub StartDateListBox_Click()
' Keep other listboxes in lockstep
GroupNameListBox.ListIndex = StartDateListBox.ListIndex
EndDateListBox.ListIndex = StartDateListBox.ListIndex
GroupTypeListBox.ListIndex = StartDateListBox.ListIndex
End Sub

Private Sub EndDateListBox_Click()
' Keep other listboxes in lockstep
GroupNameListBox.ListIndex = EndDateListBox.ListIndex
StartDateListBox.ListIndex = EndDateListBox.ListIndex
GroupTypeListBox.ListIndex = EndDateListBox.ListIndex
End Sub

Private Sub GroupTypeListBox_Click()
' Keep other listboxes in lockstep
GroupNameListBox.ListIndex = GroupTypeListBox.ListIndex
StartDateListBox.ListIndex = GroupTypeListBox.ListIndex
EndDateListBox.ListIndex = GroupTypeListBox.ListIndex
End Sub

'' USERFORM/MULTIPAGE===============================================================

Private Sub Userform_Initialize()
' Add items into listboxes based on cells in specified worksheets

' Name of non-specific defaults sheet
Dim non As String
non = "Non-Specific Defaults"
' Name of type-specific defaults string
Dim spec As String
spec = "Type-Specific Defaults"

' EVENT ID
Call funcs.RefreshListBox("Data", 1, EventIDListBox)
' GROUP ID
Call funcs.RefreshListBox("Data", 72, GroupIDListBox)
' CATEGORY
Call funcs.RefreshListBox(non, 4, CategoryListBox)
CategoryListBox.ListIndex = 0
' TYPE
Call funcs.RefreshListBox(spec, 1, TypeListBox)
' LOCATION
Call funcs.RefreshListBox(non, 1, LocationListBox)
LocationListBox.ListIndex = 0
' ROOM
Call funcs.RefreshListBox(non, 2, RoomListBox)
' GENRE
Call funcs.RefreshListBox(non, 13, GenreListBox)
' AUDIENCE
Call funcs.RefreshListBox(non, 5, AudienceListBox)
AudienceListBox.ListIndex = 0
' AUDITORIUM LAYOUT
Call funcs.RefreshListBox(non, 6, AuditoriumLayoutListBox)
' EGREMONT LAYOUT
Call funcs.RefreshListBox(non, 8, EgremontLayoutListBox)
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

Private Sub EditToggleCheckBox4_Click()
' Change some button captions and checkbox values
Call ToggleEditMode(EditToggleCheckBox4.value)
End Sub

Private Sub EditToggleCheckBox5_Click()
' Change some button captions and checkbox values
Call ToggleEditMode(EditToggleCheckBox5.value)
End Sub

'' FUNCTIONS===============================================================
' Should move these into their own modules

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
' mode = False = Add event
' mode = True = Edit event

' Store which mode we're in publically
If mode = True Then
    EditingCheck = True
Else
    EditingCheck = False
End If

' Check whether the required information has been completed or not
' Basic Info
If mode = True And EventIDListBox.ListIndex = -1 Then
    MsgBox ("You are in edit mode. Please select an event to edit on the Home page")
    Exit Function
ElseIf NameTextBox.Text = "" Then
    MsgBox ("Please enter an event name")
    MultiPage1.value = 1
    Exit Function
ElseIf StartDateTextBox.Text = "" Then
    MsgBox ("Please select a date using the calendar. Double click on the text box to show the calendar.")
    MultiPage1.value = 1
    Exit Function
ElseIf IsNull(CategoryListBox.value) Then 'CategoryListBox.ListIndex = -1 Then
    MsgBox ("Please select a category")
    MultiPage1.value = 1
    Exit Function
ElseIf IsNull(TypeListBox.value) Then
    MsgBox ("Please enter a type")
    MultiPage1.value = 1
    Exit Function
ElseIf IsNull(LocationListBox.value) Then
    MsgBox ("Please select a location")
    MultiPage1.value = 1
    Exit Function
ElseIf IsNull(RoomListBox.value) Then
    MsgBox ("Please select a room")
    MultiPage1.value = 1
    Exit Function
ElseIf IsNull(AudienceListBox.value) Then
    MsgBox ("Please enter an audience type")
    MultiPage1.value = 1
    Exit Function
ElseIf MorningCheckBox.value = 0 And AfternoonCheckBox.value = 0 And EveningCheckBox.value = 0 Then
    MsgBox ("Please select a time")
    MultiPage1.value = 1
    Exit Function
ElseIf TicketedOptionButton.value = False And NonTicketedOptionButton.value = False Then
    MsgBox ("Please select 'Ticketed' or 'Non-Ticketed'")
    MultiPage1.value = 1
    Exit Function
ElseIf GenreListBox.Visible = True And IsNull(GenreListBox.value) Then
    MsgBox ("Please enter a genre")
    MultiPage1.value = 1
    Exit Function
' Layout & Capacity
ElseIf IsNull(AuditoriumLayoutListBox.value) Then
    MsgBox ("Please enter a layout for the Auditorium")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
ElseIf IsNull(EgremontLayoutListBox.value) Then
    MsgBox ("Please enter a layout for the Egremont Room")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
' Event Time
ElseIf SetupStartTimeTextBox.Text = "" Then
    MsgBox ("Please enter a setup start time")
    MultiPage1.value = 3
    Exit Function
ElseIf DoorsTimeTextBox.Text = "" Then
    MsgBox ("Please enter the time the doors open")
    MultiPage1.value = 3
    Exit Function
ElseIf EventStartTimeTextBox.Text = "" Then
    MsgBox ("Please enter the time that the event starts")
    MultiPage1.value = 3
    Exit Function
ElseIf EventEndTimeTextBox.Text = "" Then
    MsgBox ("Please enter the time that the event ends")
    MultiPage1.value = 3
    Exit Function
ElseIf TakedownEndTimeTextBox.Text = "" Then
    MsgBox ("Please enter the time that the takedown for the event ends")
    MultiPage1.value = 3
    Exit Function
ElseIf EventDurationTextBox.Text = "" Then
    MsgBox ("Please enter the duration of the event in minutes")
    MultiPage1.value = 3
    Exit Function
ElseIf SetupToTakedownEndDurationTextBox.Text = "" Then
    MsgBox ("Please enter the duration of the event from the start " _
            & "of setup to when takedown finishes in minutes")
    MultiPage1.value = 3
    Exit Function
ElseIf BarOpenOptionButton = True And BarOpenDurationTextBox.Text = "" Then
    MsgBox ("You have selected that the bar is/was open. " & _
            "Please enter the number of minutes it is/was open for.")
    Exit Function
Else ' The user is allowed to create a new event
End If

Dim my_row As Long
If mode = False Then
    ' Next available row
    my_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row + 1
ElseIf mode = True Then
    ' Row corresponds to the row of the event selected by the user
    my_row = EventIDListBox.ListIndex + 2
End If

' Add default data into spreadsheet can be overridden by user in the future
Dim i As Integer
For i = 0 To 5
    ' Add default minutes worked by each volunteer category, depending on event category selected
    Worksheets("Data").Cells(my_row, i + 18) = Worksheets("Type-Specific Defaults").Cells(CategoryListBox.ListIndex + 2, i + 3)
Next

Dim sheet As String
sheet = "Data"

' Bar gross profit bit
Worksheets(sheet).Cells(my_row, 26) = Worksheets("Non-Specific Defaults").Cells(2, 3)
Worksheets(sheet).Cells(my_row, 27) = "=RC[-2]*RC[-1]"

' Add data given by user into spreadsheet
' Basic Info
' E stands for "event" meaning this is an event, not a group
Worksheets(sheet).Cells(my_row, 1) = "E" & funcs.UUIDGenerator(CategoryListBox.value, _
                                        StartDateTextBox.Text, NameTextBox.Text)
' S stands for "single" meaning this event isn't in a group
Worksheets(sheet).Cells(my_row, 72) = "S" & funcs.UUIDGenerator(CategoryListBox.value, _
                                        StartDateTextBox.Text, NameTextBox.Text)
Worksheets(sheet).Cells(my_row, 2) = NameTextBox.Text
Worksheets(sheet).Cells(my_row, 3) = StrManip.ConvertDate(StartDateTextBox.Text)
Worksheets(sheet).Cells(my_row, 4) = LocationListBox.value
Worksheets(sheet).Cells(my_row, 24) = CategoryListBox.value
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
Worksheets(sheet).Cells(my_row, 48) = GenreListBox.value
Worksheets(sheet).Cells(my_row, 61) = BarOpenOptionButton.value
' Layout & Capacity
Worksheets(sheet).Cells(my_row, 31) = EgremontLayoutListBox.value
Worksheets(sheet).Cells(my_row, 32) = AuditoriumLayoutListBox.value
Worksheets(sheet).Cells(my_row, 33) = TotalCapacityTextBox.Text
Worksheets(sheet).Cells(my_row, 34) = BlockedSeatsTextBox.Text
Worksheets(sheet).Cells(my_row, 45) = TotalCapacityTextBox.Text - BlockedSeatsTextBox.Text

' Time
' Event
Worksheets(sheet).Cells(my_row, 8) = TimeValue(SetupStartTimeTextBox)
Worksheets(sheet).Cells(my_row, 47) = TimeValue(DoorsTimeTextBox)
Worksheets(sheet).Cells(my_row, 9) = TimeValue(EventStartTimeTextBox)
Worksheets(sheet).Cells(my_row, 10) = TimeValue(EventEndTimeTextBox)
Worksheets(sheet).Cells(my_row, 11) = TimeValue(TakedownEndTimeTextBox)
Worksheets(sheet).Cells(my_row, 13) = EventDurationTextBox.Text
Worksheets(sheet).Cells(my_row, 12) = SetupToTakedownEndDurationTextBox.Text
Worksheets(sheet).Cells(my_row, 15) = SetupTakedownTextBox.Text
' Bar
If BarOpenOptionButton = True Then
    ' TimeValue throws an error if text box is empty
    If BarSetupTimeTextBox.Text <> "" Then
        Worksheets(sheet).Cells(my_row, 54) = TimeValue(BarSetupTimeTextBox.Text)
    End If
    If BarOpenTimeTextBox.Text <> "" Then
        Worksheets(sheet).Cells(my_row, 55) = TimeValue(BarOpenTimeTextBox.Text)
    End If
    If BarCloseTimeTextBox.Text <> "" Then
        Worksheets(sheet).Cells(my_row, 56) = TimeValue(BarCloseTimeTextBox.Text)
    End If
        
    Worksheets(sheet).Cells(my_row, 57) = BarOpenDurationTextBox.Text
    Worksheets(sheet).Cells(my_row, 58) = BarSetupToTakedownEndDurationTextBox.Text
    Worksheets(sheet).Cells(my_row, 59) = BarSetupTakedownTextBox.Text
    
    ' Bar profit per hour
    If Worksheets(sheet).Cells(my_row, 55) > 0 Then
        Worksheets(sheet).Cells(my_row, 60) = "=RC[-33]*60/RC[-3]"
    End If
Else
    ' the bar isn't/wasn't open, so enter nothing
End If

' Costs & Income
Worksheets(sheet).Cells(my_row, 14) = NumTicketsSoldTextBox.Text
Worksheets(sheet).Cells(my_row, 42) = BoxOfficeRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 44) = SupportRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 70) = RoomHireRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 71) = MiscRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 35) = FilmCostTextBox.Text
Worksheets(sheet).Cells(my_row, 36) = FilmTransportTextBox.Text
Worksheets(sheet).Cells(my_row, 37) = AccommodationTextBox.Text
Worksheets(sheet).Cells(my_row, 38) = ArtistFoodTextBox.Text
Worksheets(sheet).Cells(my_row, 43) = HiredPersonnelTextBox.Text
Worksheets(sheet).Cells(my_row, 39) = HeatingTextBox.Text
Worksheets(sheet).Cells(my_row, 40) = LightingTextBox.Text
Worksheets(sheet).Cells(my_row, 41) = MiscCostTextBox.Text
Worksheets(sheet).Cells(my_row, 25) = BarRevenueTextBox.Text

Dim TotalCosts As Double ' Total costs
Dim TotalRevenueExcBar As Double ' Total revenue minus bar
Dim TotalRevenueIncBar As Double ' Total revenue including bar

TotalCosts = BoxOfficeRevenueTextBox.Text + 0
' We need to be selling tickets ourselves and have positive box office revenue
'   and know how many tickets we've sold
If TicketedOptionButton.value = True And BoxOfficeRevenueTextBox.Text > 0 _
    And NumTicketsSoldTextBox.Text <> "" Then
    
    ' Calculate Ticketsolve fee estimate
    ' 80p per ticket sold
    Worksheets(sheet).Cells(my_row, 51) = 0.8 * NumTicketsSoldTextBox.Text
    Worksheets(sheet).Cells(my_row, 50) = BoxOfficeRevenueTextBox.Text - _
        0.8 * NumTicketsSoldTextBox.Text
End If

If BarMarginTextBox.Text <> "" Then
    Worksheets(sheet).Cells(my_row, 26) = CDbl(BarMarginTextBox.Text) * 0.01 ' convert percentage into decimal
End If

' Calculate contribution to overheads with and without bar
' Uneccessary because of pivot table formulae
Worksheets(sheet).Cells(my_row, 52) = _
        "=RC[-8]+RC[-10]+RC[18]-RC[-1]-RC[-11]-RC[-12]-RC[-13]-RC[-14]-RC[-15]-RC[-16]-RC[-17]"
Worksheets(sheet).Cells(my_row, 53) = "=RC[-1]+RC[-26]"

' Volunteer Minutes
Worksheets(sheet).Cells(my_row, 18) = FoHTextBox.Text
Worksheets(sheet).Cells(my_row, 19) = DMTextBox.Text
Worksheets(sheet).Cells(my_row, 20) = TechTextBox.Text
Worksheets(sheet).Cells(my_row, 22) = BarTextBox.Text
Worksheets(sheet).Cells(my_row, 23) = AoWVolTextBox.Text
Worksheets(sheet).Cells(my_row, 49) = MiscVolTextBox.Text
' Volunteer Nominal Pay
Worksheets(sheet).Cells(my_row, 62) = FoHPayTextBox.Text
Worksheets(sheet).Cells(my_row, 63) = DMPayTextBox.Text
Worksheets(sheet).Cells(my_row, 64) = TechPayTextBox.Text
Worksheets(sheet).Cells(my_row, 66) = BarPayTextBox.Text
Worksheets(sheet).Cells(my_row, 67) = AoWVolPayTextBox.Text
Worksheets(sheet).Cells(my_row, 68) = MiscVolPayTextBox.Text

'Add a 1 in the column counting the number of events.
' Need to do this because pivot tables can't be added to the data model.
Worksheets(sheet).Cells(my_row, 69) = 1

' Declare that this is an event, not a dummy event representing a group
Worksheets(sheet).Cells(my_row, 73) = False

' Update pivot table(s)
Call funcs.ChangeSource(sheet, "Analysis", "PivotTable1")

' Change message depending on mode
Dim response As Variant
If mode = False Then
    MsgBox ("Event Added")
    response = MsgBox("Would you like to add this to a group?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Question")
ElseIf mode = True Then
    MsgBox ("Event Edited")
Else
    MsgBox ("mode = Null when adding event. Contact support")
End If

If response = vbYes Then
    MultiPage1.value = 6
    GroupSearchBox.Text = NameTextBox.Text
End If

EditingCheck = False

' Update listboxes
Call funcs.RefreshListBox("Data", 1, EventIDListBox)
Call funcs.RefreshListBox("Data", 72, GroupIDListBox)

Call NewSearchTextBox_Change

' Change selected event to one just added, if one was added (and not edited)
If EventIDListBox.ListCount > 0 And mode = False Then
    EventIDListBox.ListIndex = EventIDListBox.ListCount - 1
End If

End Function

Private Function AuditoriumUsed()
' To be called when Auditorium is being used

Dim empty_row As Long
Dim ListBoxRange As Range

' Find last non-empty row for auditorium layout list
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 6).End(xlUp).row

Set ListBoxRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 6), _
                    Worksheets("Non-Specific Defaults").Cells(empty_row, 6))
AuditoriumLayoutListBox.RowSource = ListBoxRange.address(External:=True)

' Store current multipage1 index
Dim currentIndex As Integer
currentIndex = MultiPage1.value

' Change page we're on so that everything loads
MultiPage1.value = 2

' Set default value
AuditoriumLayoutListBox.ListIndex = 1
MultiPage1.value = currentIndex ' reset index

AuditoriumCapacityTextBox.Locked = False
TotalCapacityTextBox.Locked = False
End Function

Private Function EgremontUsed()
' To be called when Egremont room is being used

Dim empty_row As Long
Dim ListBoxRange As Range

' Find last non-empty row for auditorium layout list
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 8).End(xlUp).row

Set ListBoxRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 8), _
                    Worksheets("Non-Specific Defaults").Cells(empty_row, 8))
EgremontLayoutListBox.RowSource = ListBoxRange.address(External:=True)

' Store current multipage1 index
Dim currentIndex As Integer
currentIndex = MultiPage1.value

' Change page we're on so that everything loads
MultiPage1.value = 2

' Set default value
EgremontLayoutListBox.ListIndex = 1
MultiPage1.value = currentIndex ' reset index

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

Dim ListBoxRange As Range
Set ListBoxRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 6), _
                    Worksheets("Non-Specific Defaults").Cells(2, 6))
AuditoriumLayoutListBox.RowSource = ListBoxRange.address(External:=True)

' Store current multipage1 index
Dim currentIndex As Integer
currentIndex = MultiPage1.value

' Change page we're on so that everything loads
MultiPage1.value = 2

' Set default value
AuditoriumLayoutListBox.ListIndex = 0
MultiPage1.value = currentIndex ' reset index

AuditoriumCapacityTextBox.Locked = True
AuditoriumCapacityTextBox.value = "0"
End Function

Private Function EgremontNotUsed()
' To be called when Egremont room is not being used

Dim ListBoxRange As Range
Set ListBoxRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 8), _
                    Worksheets("Non-Specific Defaults").Cells(2, 8))
EgremontLayoutListBox.RowSource = ListBoxRange.address(External:=True)

' Store current multipage1 index
Dim currentIndex As Integer
currentIndex = MultiPage1.value

' Change page we're on so that everything loads
MultiPage1.value = 2
' Set default value
EgremontLayoutListBox.ListIndex = 0
MultiPage1.value = currentIndex ' reset index

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

Private Sub ToggleEditMode(state As Boolean)
If state = True Then
    ' Update captions
    EventButton1.Caption = "Edit Selected Event"
    EventButton2.Caption = "Edit Selected Event"
    EventButton3.Caption = "Edit Selected Event"
    EventButton4.Caption = "Edit Selected Event"
    EventButton5.Caption = "Edit Selected Event"
    
    ' Tick all of the other edit checkboxes
    EditToggleCheckBox1.value = True
    EditToggleCheckBox2.value = True
    EditToggleCheckBox3.value = True
    EditToggleCheckBox4.value = True
    EditToggleCheckBox5.value = True
ElseIf state = False Then
    ' Update captions
    EventButton1.Caption = "Add New Event"
    EventButton2.Caption = "Add New Event"
    EventButton3.Caption = "Add New Event"
    EventButton4.Caption = "Add New Event"
    EventButton5.Caption = "Add New Event"

    ' Untick all of the other edit checkboxes
    EditToggleCheckBox1.value = False
    EditToggleCheckBox2.value = False
    EditToggleCheckBox3.value = False
    EditToggleCheckBox4.value = False
    EditToggleCheckBox5.value = False
End If
End Sub

Private Sub AutofillEventFromSelected(row As Variant)
' Don't autofill if user is trying to edit
If EditingCheck = True Then
    Exit Sub
End If

' Tell programme if autofill is happening
AutofillCheck = True
' Fill in everything with values taken from this event.

' Sheet data is stored on. Makes things easier to read.
Dim sheet As String
sheet = "Data"

' Store current page user is on
Dim currentPage As Integer
currentPage = MultiPage1.value

' Basic Info
' Go to page so that everything loads
MultiPage1.value = 1

NameTextBox.Text = Worksheets(sheet).Cells(row, 2)
StartDateTextBox.Text = Worksheets(sheet).Cells(row, 3)
If Worksheets(sheet).Cells(row, 24) <> "" Then
    CategoryListBox.value = Worksheets(sheet).Cells(row, 24)
Else
    ' Unselect item if no info provided
    CategoryListBox.ListIndex = -1
    MsgBox ("Nothing in category")
End If

If Worksheets(sheet).Cells(row, 29) <> "" Then
    TypeListBox.value = Worksheets(sheet).Cells(row, 29)
Else
    ' Unselect item if no info provided
    TypeListBox.ListIndex = -1
    MsgBox ("Nothing in type")
End If

If Worksheets(sheet).Cells(row, 4) <> "" Then
    LocationListBox.value = Worksheets(sheet).Cells(row, 4)
Else
    ' Unselect item if no info provided
    LocationListBox.ListIndex = -1
    MsgBox ("nothing in location")
End If

If Worksheets(sheet).Cells(row, 28) <> "" Then
    RoomListBox.value = Worksheets(sheet).Cells(row, 28)
Else
    RoomListBox.ListIndex = -1
    MsgBox ("nothing in room")
End If

If Worksheets(sheet).Cells(row, 30) <> "" Then
    AudienceListBox.value = Worksheets(sheet).Cells(row, 30)
Else
    AudienceListBox.ListIndex = -1
    MsgBox ("nothing in audience")
End If

MorningCheckBox.value = Worksheets(sheet).Cells(row, 5)
AfternoonCheckBox.value = Worksheets(sheet).Cells(row, 6)
EveningCheckBox.value = Worksheets(sheet).Cells(row, 7)

If Worksheets(sheet).Cells(row, 46) = True Then
    TicketedOptionButton.value = True
    NonTicketedOptionButton.value = False
Else
    TicketedOptionButton.value = False
    NonTicketedOptionButton.value = True
End If

If Worksheets(sheet).Cells(row, 61) = True Then
    BarOpenOptionButton.value = True
    BarNotOpenOptionButton.value = False
Else
    BarOpenOptionButton.value = False
    BarNotOpenOptionButton.value = True
End If

If Worksheets(sheet).Cells(row, 48) <> "" Then
    GenreListBox.value = Worksheets(sheet).Cells(row, 48)
Else
    GenreListBox.ListIndex = -1
End If

' Layout & Capacity
' Go to page so that everything loads
MultiPage1.value = 2

If Worksheets(sheet).Cells(row, 32) <> "" Then
    AuditoriumLayoutListBox.value = Worksheets(sheet).Cells(row, 32)
Else
    AuditoriumLayoutListBox.ListIndex = -1
    MsgBox ("Nothing in auditorium")
End If

If Worksheets(sheet).Cells(row, 31) <> "" Then
    EgremontLayoutListBox.value = Worksheets(sheet).Cells(row, 31)
Else
    EgremontLayoutListBox.ListIndex = -1
    MsgBox ("Nothing in egremont")
End If

TotalCapacityTextBox.Text = Worksheets(sheet).Cells(row, 33)
BlockedSeatsTextBox.Text = Worksheets(sheet).Cells(row, 34)

' Event Time
' Go to page so that everything loads
MultiPage1.value = 3

SetupStartTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 8), "hh:mm")
DoorsTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 47), "hh:mm")
EventStartTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 9), "hh:mm")
EventEndTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 10), "hh:mm")
TakedownEndTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 11), "hh:mm")
EventDurationTextBox.Text = Worksheets(sheet).Cells(row, 13)
SetupToTakedownEndDurationTextBox.Text = Worksheets(sheet).Cells(row, 12)
SetupTakedownTextBox.Text = Worksheets(sheet).Cells(row, 15)
' Bar Time
' Show the bar if needed
If BarOpenOptionButton = True Then
    BarTimeFrame.Visible = True
End If
BarSetupTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 54), "hh:mm")
BarOpenTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 55), "hh:mm")
BarCloseTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 56), "hh:mm")
BarOpenDurationTextBox.Text = Worksheets(sheet).Cells(row, 57)
BarSetupToTakedownEndDurationTextBox.Text = Worksheets(sheet).Cells(row, 58)
BarSetupTakedownTextBox.Text = Worksheets(sheet).Cells(row, 59)

' Income
' Go to page so that everything loads
MultiPage1.value = 4
NumTicketsSoldTextBox.Text = Worksheets(sheet).Cells(row, 14)
BoxOfficeRevenueTextBox.Text = Worksheets(sheet).Cells(row, 42)
SupportRevenueTextBox.Text = Worksheets(sheet).Cells(row, 44)
BarRevenueTextBox.Text = Worksheets(sheet).Cells(row, 25)
BarMarginTextBox.Text = Worksheets(sheet).Cells(row, 26) * 100
' Costs
FilmCostTextBox.Text = Worksheets(sheet).Cells(row, 35)
FilmTransportTextBox.Text = Worksheets(sheet).Cells(row, 36)
AccommodationTextBox.Text = Worksheets(sheet).Cells(row, 37)
ArtistFoodTextBox.Text = Worksheets(sheet).Cells(row, 38)
HiredPersonnelTextBox.Text = Worksheets(sheet).Cells(row, 43)
HeatingTextBox.Text = Worksheets(sheet).Cells(row, 39)
LightingTextBox.Text = Worksheets(sheet).Cells(row, 40)
MiscCostTextBox.Text = Worksheets(sheet).Cells(row, 41)

' Volunteer minutes
' Go to page so that everything loads
MultiPage1.value = 5
FoHTextBox.Text = Worksheets(sheet).Cells(row, 18)
DMTextBox.Text = Worksheets(sheet).Cells(row, 19)
TechTextBox.Text = Worksheets(sheet).Cells(row, 19)
BarTextBox.Text = Worksheets(sheet).Cells(row, 22)
AoWVolTextBox.Text = Worksheets(sheet).Cells(row, 23)
MiscVolTextBox.Text = Worksheets(sheet).Cells(row, 49)


' Go to original page
MultiPage1.value = currentPage
' Tell the programme we're done autofilling
AutofillCheck = False
End Sub

Private Sub AutofillGroupFromSelected(row As Variant)
' Sheet data is stored on. Makes things easier to read.
Dim sheet As String
sheet = "Data"

' Do the autofilling
GroupManagementForm.GroupNameTextBox.Text = Worksheets(sheet).Cells(row, 2)
GroupManagementForm.StartDateTextBox.Text = Worksheets(sheet).Cells(row, 3)
GroupManagementForm.EndDateTextBox.Text = Worksheets(sheet).Cells(row, 74)
GroupManagementForm.CategoryListBox.Text = Worksheets(sheet).Cells(row, 24)
GroupManagementForm.TypeListBox.Text = Worksheets(sheet).Cells(row, 29)
End Sub

Private Sub DaySpecificDefaults(sheet As String, startRow As Integer, eventType As String)
' Purpose:
' Set event and bar times based on the day and event type.
' Only works for specific boxes and setup we are using.
' Assumes we start at Monday and go through to Sunday in usual order.
'
' Input:
' sheet = sheet where defaults are found
' startRow = row where the Monday values for event type are found.
' eventType = name of event type. Used only for double checking startRow is correct.
'
' Output:
' event and bar time boxes will be filled out

' Double check startRow is correct
If Worksheets(sheet).Cells(startRow - 1, 1) <> eventType Then
    MsgBox ("in DaySpecificDefaults, eventType does not match what's on the sheet" & _
            "eventType = " & eventType & ", and on the sheet: " _
            & Worksheets(sheet).Cells(startRow - 1, 1))
End If

' Check that a date has been entered
If StartDateTextBox = "" Then
    Exit Sub
End If

' Find out day event is on
Dim day As String
day = Format(StartDateTextBox.Text, "dddd")

' Store AutoTimeCheckbox value
Dim autoTime As Boolean
autoTime = AutoTimeCheckBox.value

' Disable time autofill
AutoTimeCheckBox.value = False

' Store row to look on for defaults
Dim row As Integer
If day = "Monday" Then
    row = startRow + 0
ElseIf day = "Tuesday" Then
    row = startRow + 1
ElseIf day = "Wednesday" Then
    row = startRow + 2
ElseIf day = "Thursday" Then
    row = startRow + 3
ElseIf day = "Friday" Then
    row = startRow + 4
ElseIf day = "Saturday" Then
    row = startRow + 5
ElseIf day = "Sunday" Then
    row = startRow + 6
Else
    MsgBox ("Day not found. Error in TypeListBox_Change(). Contact support.")
    Exit Sub
End If

' Find out if first data cell in row is blank
If Worksheets(sheet).Cells(row, 2) = "" Then
    ' no point continuing because all of the times are blank
    Exit Sub
End If

' Start filling in data

' Rough time checkbox on basic info page
' Reset checkboxes
MorningCheckBox.value = False
AfternoonCheckBox.value = False
EveningCheckBox.value = False
' Set correct checkbox
If Worksheets(sheet).Cells(row, 10) = "Morning" Then
    MorningCheckBox.value = True
ElseIf Worksheets(sheet).Cells(row, 10) = "Afternoon" Then
    AfternoonCheckBox.value = True
ElseIf Worksheets(sheet).Cells(row, 10) = "Evening" Then
    EveningCheckBox.value = True
Else
    ' Do nothing because we don't know what to do
End If

' Event Times
SetupStartTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 2), "hh:mm")
DoorsTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 3), "hh:mm")
EventStartTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 4), "hh:mm")
EventEndTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 5), "hh:mm")
TakedownEndTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 6), "hh:mm")
' Event Durations
' event end time - event start time
EventDurationTextBox.Text = CInt((Worksheets(sheet).Cells(row, 5) _
                            - Worksheets(sheet).Cells(row, 4)) * 24 * 60)
' event takedown end time - event setup start time
SetupToTakedownEndDurationTextBox.Text = CInt((Worksheets(sheet).Cells(row, 6) _
                                        - Worksheets(sheet).Cells(row, 2)) * 24 * 60)
                                
' (event takedown end time - event end time) + (event start time - event setup start time)
SetupTakedownTextBox.Text = CInt((Worksheets(sheet).Cells(row, 6) _
                            - Worksheets(sheet).Cells(row, 5) _
                            + Worksheets(sheet).Cells(row, 4) _
                            - Worksheets(sheet).Cells(row, 2)) * 24 * 60)

' Bar
BarSetupTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 7), "hh:mm")
BarOpenTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 8), "hh:mm")
BarCloseTimeTextBox.Text = Format(Worksheets(sheet).Cells(row, 9), "hh:mm")
' Bar Durations
' bar close time - bar open time
BarOpenDurationTextBox.Text = CInt((Worksheets(sheet).Cells(row, 9) _
                                - Worksheets(sheet).Cells(row, 8)) * 24 * 60)

' bar close time - bar setup time
BarSetupToTakedownEndDurationTextBox.Text = CInt((Worksheets(sheet).Cells(row, 9) _
                                - Worksheets(sheet).Cells(row, 7)) * 24 * 60)

' bar open time - bar steup time
BarSetupTakedownTextBox.Text = CInt((Worksheets(sheet).Cells(row, 8) _
                                - Worksheets(sheet).Cells(row, 7)) * 24 * 60)

' Reset time autofill to previous state
AutoTimeCheckBox.value = autoTime

' Refresh volunteer hours based on new times
Call UpdateVolunteerMinutes
End Sub

Private Function UpdateVolunteerMinutes() As Boolean
' use event and bar times to update minutes worked by volunteers

Dim Ctrl As Control
' Check everything is filled in. Exit sub if not.
For Each Ctrl In Me.MultiPage1.Pages(3).Controls
    If TypeOf Ctrl Is MSForms.TextBox Then
        If Ctrl.Text = "" Then
            UpdateVolunteerMinutes = False
            Exit Function
        End If
    End If
Next

Dim FoH1 As Integer
Dim FoH2 As Integer
Dim DM As Integer
Dim Tech As Integer
Dim Bar As Integer

' 15 minutes before doors, leave as event starts
FoH1 = CInt(15 + DateDiff("n", DoorsTimeTextBox, EventStartTimeTextBox))
' 15 minutes before event starts, 5 minutes after event ends
FoH2 = CInt(15 + DateDiff("n", EventStartTimeTextBox, EventEndTimeTextBox) + 5)
' 15 minutes before doors, leave 15 minutes after event ends
DM = CInt(15 + DateDiff("n", DoorsTimeTextBox, EventEndTimeTextBox) + 15)
' 15 minutes before doors, leave 15 minutes after event ends
Tech = CInt(15 + DateDiff("n", DoorsTimeTextBox, EventEndTimeTextBox) + 15)
' 15 minutes before doors, leave after bar closes
Bar = CInt(DateDiff("n", BarSetupTimeTextBox, BarCloseTimeTextBox))

' Change page to volunteer page
Dim currentPage As Integer
currentPage = MultiPage1.value
MultiPage1.value = 5

' Fill in volunteer minutes
FoHTextBox.Text = FoH1 + FoH2
DMTextBox.Text = DM
TechTextBox.Text = Tech
BarTextBox.Text = Bar

' Go back to page user was on
MultiPage1.value = currentPage

UpdateVolunteerMinutes = True
End Function
