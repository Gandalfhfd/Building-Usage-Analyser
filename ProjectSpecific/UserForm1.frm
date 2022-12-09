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
' Dumb, but why not
Public DeleteIndicator As Integer

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

Private Sub AutofillCheckBox_Click()
If EventIDListBox.ListIndex = -1 Then
    ' No event has been selected, so do nothing
    Exit Sub
End If

' Store row we're autofilling into
Dim nameLocation As Integer
nameLocation = EventIDListBox.ListIndex + 2

' Autofill data into form
If AutofillCheckBox.value = True Then
    Call AutofillFromSelected(nameLocation)
End If
End Sub

Private Sub BarOpenOptionButton_Change()
' Hide the bar stuff if it isn't needed.
If BarOpenOptionButton = True Then
    BarTimeFrame.Visible = True
Else
    BarTimeFrame.Visible = False
End If
End Sub

'' BUTTON CLICKING===============================================================

Private Sub EventButton1_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton1_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton2_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton3_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton3_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
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
    DeleteIndicator = 1
    ' Delete entire row corresponding to selected event
    Sheets("Data").Rows(row).Delete
ElseIf EventIDListBox.ListCount = my_index + 1 Then
    ' We are at end of list, so go up one
    EventIDListBox.ListIndex = my_index - 1
    DeleteIndicator = 1
    ' Delete entire row corresponding to selected event
    Sheets("Data").Rows(row).Delete
Else
    ' Delete entire row corresponding to selected event
    DeleteIndicator = 1
    Sheets("Data").Rows(row).Delete
End If

' Update EventIDListBox
Call RefreshListBox("Data", 1, EventIDListBox)

' Update pivot table(s)
Call ChangeSource("Data", "Analysis", "PivotTable1")
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
Call ImportFromTicketsolve("Selected")
End Sub

Private Sub TicketsolveImportPreviousButton_Click()
' Import data into the event which was most recently added
Call ImportFromTicketsolve("Previous")
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
Call ImportFromZettle("Selected")
End Sub
Private Sub ZettleImportPreviousButton_Click()
' Import info from Zettle
Call ImportFromZettle("Previous")
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

'' TEXT BOXES===============================================================

' Basic Info============================================================================
Private Sub EventDateTextBox_DblClick(ByVal Cancel As MsForms.ReturnBoolean)
' Re-show Date Picker if box has already been entered
Call GetCalendar ' Show Date Picker
End Sub

Private Sub EventDateTextBox_Enter()
Call GetCalendar ' Show Date Picker
End Sub

Private Sub EventDateTextBox_change()
' Update times if date is changed and type matches what we want
If TypeListBox.value = "Film" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 2, "Film")
ElseIf TypeListBox.value = "Live Music" And GenreListBox.value = "Jazz" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 10, "Jazz")
End If
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

Private Sub TestButton_Click()
Call UpdateVolunteerMinutes
End Sub

Private Sub TotalCapacityTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(TotalCapacityTextBox, TotalCapacity)
End Sub

Private Sub BlockedSeatsTextBox_Change()
' Sanitise input to ensure only real numbers <= 100 are input
Call InptValid.SanitiseNonNegInt(BlockedSeatsTextBox, BlockedSeats)
End Sub

Private Sub BlockedSeatsTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Stop it from being left blank
If BlockedSeatsTextBox.Text = "" Then
    BlockedSeatsTextBox.Text = 0
End If
End Sub

' Time============================================================================

Private Sub SetupStartTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub DoorsTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub EventStartTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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
BarSetupTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 15) - _
    Worksheets(sheet).Cells(row, 16), EventStartTimeTextBox.Text), "hh:mm")

' Change bar open time
BarOpenTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 15), _
    EventStartTimeTextBox.Text), "hh:mm")
    
' Change bar close time
BarCloseTimeTextBox.Text = EventStartTimeTextBox.Text

' Update volunteer hours
Call UpdateVolunteerMinutes
End Sub

Private Sub EventEndTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub TakedownEndTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub BarSetupTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarSetupTimeTextBox, BarSetupTime)
End Sub

Private Sub BarOpenTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarOpenTimeTextBox, BarOpenTime)
End Sub

Private Sub BarCloseTimeTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub BoxOfficeRevenueTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
BoxOfficeRevenueTextBox.Text = StrManip.Convert2Currency(BoxOfficeRevenueTextBox)
End Sub

Private Sub SupportRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(SupportRevenueTextBox, SupportRevenue)
End Sub

Private Sub SupportRevenueTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
SupportRevenueTextBox.Text = StrManip.Convert2Currency(SupportRevenueTextBox)
End Sub

Private Sub BarRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BarRevenueTextBox, BarRevenue)
End Sub

Private Sub BarRevenueTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub FilmCostTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
FilmCostTextBox.Text = StrManip.Convert2Currency(FilmCostTextBox)
End Sub

Private Sub FilmTransportTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmTransportTextBox, FilmTransport)
End Sub

Private Sub FilmTransportTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
FilmTransportTextBox.Text = StrManip.Convert2Currency(FilmTransportTextBox)
End Sub

Private Sub AccommodationTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(AccommodationTextBox, Accommodation)
End Sub

Private Sub AccommodationTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
AccommodationTextBox.Text = StrManip.Convert2Currency(AccommodationTextBox)
End Sub

Private Sub ArtistFoodTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(ArtistFoodTextBox, ArtistFood)
End Sub

Private Sub ArtistFoodTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
ArtistFoodTextBox.Text = StrManip.Convert2Currency(ArtistFoodTextBox)
End Sub

Private Sub HeatingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HeatingTextBox, Heating)
End Sub

Private Sub HeatingTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
HeatingTextBox.Text = StrManip.Convert2Currency(HeatingTextBox)
End Sub

Private Sub LightingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(LightingTextBox, Lighting)
End Sub

Private Sub LightingTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
LightingTextBox.Text = StrManip.Convert2Currency(LightingTextBox)
End Sub

Private Sub MiscCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(MiscCostTextBox, MiscCost)
End Sub

Private Sub MiscCostTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
MiscCostTextBox.Text = StrManip.Convert2Currency(MiscCostTextBox)
End Sub

Private Sub HiredPersonnelTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HiredPersonnelTextBox, HiredPersonnel)
End Sub

Private Sub HiredPersonnelTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
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

Private Sub FoHPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
FoHPayTextBox.Text = StrManip.Convert2Currency(FoHPayTextBox)
End Sub

Private Sub DMPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(DMPayTextBox, DMPay)
End Sub

Private Sub DMPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
DMPayTextBox.Text = StrManip.Convert2Currency(DMPayTextBox)
End Sub

Private Sub TechPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(TechPayTextBox, TechPay)
End Sub

Private Sub TechPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
TechPayTextBox.Text = StrManip.Convert2Currency(TechPayTextBox)
End Sub

Private Sub BoxOfficePayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(BoxOfficePayTextBox, BoxOfficePay)
End Sub

Private Sub BoxOfficePayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
BoxOfficePayTextBox.Text = StrManip.Convert2Currency(BoxOfficePayTextBox)
End Sub

Private Sub BarPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(BarPayTextBox, BarPay)
End Sub

Private Sub BarPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
BarPayTextBox.Text = StrManip.Convert2Currency(BarPayTextBox)
End Sub

Private Sub AoWVolPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(AoWVolPayTextBox, AoWVolPay)
End Sub

Private Sub AoWVolPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
AoWVolPayTextBox.Text = StrManip.Convert2Currency(AoWVolPayTextBox)
End Sub

Private Sub MiscVolPayTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseReal(MiscVolPayTextBox, MiscVolPay)
End Sub

Private Sub MiscVolPayTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
' Sanitise input to ensure only non-negative integers are input
MiscVolPayTextBox.Text = StrManip.Convert2Currency(MiscVolPayTextBox)
End Sub

Private Sub NewSearchTextBox_Change()

Dim empty_row As Long ' Store number of items in list box
Dim DataRange As Range

' empty_row = lst non-empty row for specific list(box)
empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row
Set DataRange = Union(Range(Worksheets("Data").Cells(2, 1), _
                Worksheets("Data").Cells(empty_row, 1)), _
                Range(Worksheets("Data").Cells(2, 2), _
                Worksheets("Data").Cells(empty_row, 2)))

' Clear items to avoid them being re-added
ListBox1.ListIndex = -1
ListBox1.Clear

' Use "Union(Range1, Range2)"

Call funcs.AddAllToListBox(NewSearchTextBox.Text, DataRange, Array(1), ListBox1, False)
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
If TypeListBox.value = "Film" And EventDateTextBox.Text <> "" Then
    Call DaySpecificDefaults("Day & Type-Specific Defaults", 2, "Film")
ElseIf TypeListBox.value = "Live Music" And EventDateTextBox.Text <> "" _
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
'MsgBox ("EventIDListBox_Change was called")

' Stop it freaking out when an item is deleted
If DeleteIndicator = 1 Then
    DeleteIndicator = 2
    Exit Sub
ElseIf DeleteIndicator = 2 Then
    ' Reset DeleteIndicator on second time of asking
    DeleteIndicator = 0
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
' Display Event ID in search box
If EventIDListBox.ListIndex <> -1 Then
    SearchBox.Text = EventIDListBox.value
End If

' Store row of selected event.
Dim row As Long
row = EventIDListBox.ListIndex + 2

' Display the event name in search box.
' It will always be found in the second column, so hard coding is ok.
NameSearchTextBox.Text = Worksheets(sheet).Cells(row, 2)

' So the user always knows which event has been selected.
' Might want to hide that if they're not in edit mode, idk.
EventIDUpdaterLabel1.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel2.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel3.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel4.Caption = "Selected Event ID: " & EventIDListBox.value
EventIDUpdaterLabel5.Caption = "Selected Event ID: " & EventIDListBox.value

If AutofillCheckBox.value = True Then
    ' Fill in everything with values taken from this event.
    Call AutofillFromSelected(row)
End If
End Sub

'' USERFORM/MULTIPAGE===============================================================

Private Sub UserForm_Initialize()
' Add items into listboxes based on cells in specified worksheets

' Name of non-specific defaults sheet
Dim non As String
non = "Non-Specific Defaults"
' Name of type-specific defaults string
Dim spec As String
spec = "Type-Specific Defaults"

' EVENT ID
Call RefreshListBox("Data", 1, EventIDListBox)
' CATEGORY
Call RefreshListBox(non, 4, CategoryListBox)
' TYPE
Call RefreshListBox(spec, 1, TypeListBox)
' LOCATION
Call RefreshListBox(non, 1, LocationListBox)
' ROOM
Call RefreshListBox(non, 2, RoomListBox)
' GENRE
Call RefreshListBox(non, 13, GenreListBox)
' AUDIENCE
Call RefreshListBox(non, 5, AudienceListBox)
' AUDITORIUM LAYOUT
Call RefreshListBox(non, 6, AuditoriumLayoutListBox)
' EGREMONT LAYOUT
Call RefreshListBox(non, 8, EgremontLayoutListBox)
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
' Generate uniqueish UUID.
' If name, category and date are all the same, there is a 1 in 844,596,301 change of a collision.
UUIDGenerator = InptValid.RmSpecialChars(name) & InptValid.RmSpecialChars(category) _
                & InptValid.RmSpecialChars(eventDate) & funcs.GenerateRandomAlphaNumericStr(5)
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
ElseIf EventDateTextBox.Text = "" Then
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
Worksheets(sheet).Cells(my_row, 1) = UUIDGenerator(CategoryListBox.value, EventDateTextBox.Text, NameTextBox.Text)
Worksheets(sheet).Cells(my_row, 2) = NameTextBox.Text
Worksheets(sheet).Cells(my_row, 3) = StrManip.ConvertDate(EventDateTextBox.Text)
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
        "=RC[-8]+RC[-10]-RC[-1]-RC[-11]-RC[-12]-RC[-13]-RC[-14]-RC[-15]-RC[-16]-RC[-17]"
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

' Update pivot table(s)
Call ChangeSource(sheet, "Analysis", "PivotTable1")

' Change message depending on mode
If mode = False Then
    MsgBox ("Event Added")
ElseIf mode = True Then
    MsgBox ("Event Edited")
Else
    MsgBox ("mode = Null when adding event. Contact support")
End If

EditingCheck = False
Call RefreshListBox("Data", 1, EventIDListBox)

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

Dim value As String ' Store target cell's value
myAddress = funcs.search(keyword, importSheet)
If myAddress(0) = "0" Then
    value = substitute
    MsgBox (errorMsg)
    succeeded = False
Else
    ' Store value of cell we're interested in importing
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

Private Sub ImportFromTicketsolve(mode As String)
' Highly non-general function/sub which imports from the "Events Summary" csv from Ticketsolve

' Store which row we're working on. Depends on what we're after.
Dim my_row As Long

' Not strictly necessary. Used to avoid too much hard-coding.
Dim dataSheetName As String
dataSheetName = "Data"

' Decide which event to import info into
If mode = "Selected" Then
    my_row = funcs.search(SearchBox.Text, dataSheetName)(0)
ElseIf mode = "Previous" Then
    my_row = Worksheets(dataSheetName).Cells(Rows.Count, 1).End(xlUp).row
Else
    MsgBox ("Programmer error in Private Sub ImportFromTicketsolve")
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
sheetName = "TicketsolveImport"

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


Dim tempSuccessCheck As Boolean
Dim exportCell() As Variant ' stores the cell we're working on
Dim offset() As Variant

'Find total number of ticket sales
exportCell = Array(my_row, 14)
offset = Array(0, 1)

tempSuccessCheck = ImportCell("Sold", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find event capacity
exportCell(1) = 33
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

' Find total Support the Kirkgate revenue
exportCell(1) = 44
offset = Array(0, 3)
tempSuccessCheck = ImportCell("Support the Kirkgate", sheetName, dataSheetName, exportCell, offset)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Find total ticket revenue
exportCell(1) = 42
offset = Array(1, 1)
' Find net sales
tempSuccessCheck = ImportCell("Tax", sheetName, dataSheetName, exportCell, offset)

' Find ticket sale revenue by getting net
Worksheets(dataSheetName).Cells(my_row, 42) = Worksheets(dataSheetName).Cells(my_row, 42)

If tempSuccessCheck = False Then
    succeeded = False
End If

' Must go at the bottom
If succeeded = True Then
    MsgBox ("Import successful")
End If

' Determine actual capacity and write it to a cell
Dim trueCapacity As Integer
trueCapacity = Worksheets(dataSheetName).Cells(my_row, 33) _
                - Worksheets(dataSheetName).Cells(exportCell(0), 34)
Worksheets(dataSheetName).Cells(my_row, 45) = trueCapacity

' Find Ticketsolve fees and revenue after fees

' We need to be selling tickets ourselves and have positive box office revenue
'   and know how many tickets we've sold
If Worksheets(dataSheetName).Cells(my_row, 46) = True And _
    Worksheets(dataSheetName).Cells(my_row, 42) > 0 And _
    Worksheets(dataSheetName).Cells(my_row, 14) <> "" Then
    
    ' Calculate Ticketsolve fee estimate
    ' 80p per ticket sold
    Worksheets(dataSheetName).Cells(my_row, 51) = 0.8 * Worksheets(dataSheetName).Cells(my_row, 14)
    Worksheets(dataSheetName).Cells(my_row, 50) = Worksheets(dataSheetName).Cells(my_row, 42) - _
        0.8 * Worksheets(dataSheetName).Cells(my_row, 14)
End If

' Update pivot table(s)
Call ChangeSource(dataSheetName, "Analysis", "PivotTable1")
End Sub

Private Sub ImportFromZettle(mode As String)
' FUNCTION NOT COMPLETED YET

' Highly non-general function/sub which imports from the "Raw data Excel" file from Zettle

' Store which row we're working on. Depends on what we're after.
Dim my_row As Long

' Not strictly necessary. Used to avoid too much hard-coding.
Dim dataSheetName As String
dataSheetName = "Data"

' Decide which event to import info into
If mode = "Selected" Then
    my_row = funcs.search(SearchBox.Text, dataSheetName)(0)
ElseIf mode = "Previous" Then
    my_row = Worksheets(dataSheetName).Cells(Rows.Count, 1).End(xlUp).row
Else
    MsgBox ("Programm   er error in Private Sub ImportFromZettle")
    Exit Sub
End If

' Store whether import is successful or not.
Dim succeeded As Boolean
succeeded = True

' Check if the bar is/was open selling. If not, prompt the user to update
'   the event before importing
If Worksheets(dataSheetName).Cells(my_row, 46) = False Then
    MsgBox ("The bar is/was not open for this event. If you wish to import " _
            & "data from Zettle, please mark the event as 'Bar Open'.")
    Exit Sub
End If

' Not strictly necessary. Used to avoid too much hard-coding.
Dim sheetName As String
sheetName = "ZettleImport"

Dim xlsxImportSuccessCheck As Boolean
' Import the xlsx selected by the user into sheet "sheetName"
xlsxImportSuccessCheck = funcs.xlsx_Import(sheetName)
If xlsxImportSuccessCheck = True Then
    ' Continue with sub
Else
    ' Exit sub, since a file wasn't selected
    Exit Sub
End If

' Find the first row of data
' First row = dateAddress(0) + 1
' Store address of various things
Dim dateAddress As Variant

' Find address of "Date" in "ZettleImport" sheet.
dateAddress = funcs.search("Date", sheetName)
If dateAddress(0) = "0" Then
    MsgBox ("Import failed. Did you select the correct file?")
    succeeded = False
    Exit Sub
End If

' Find the last row of data
Dim FinalRow As Integer
FinalRow = Worksheets(sheetName).Cells(Rows.Count, 1).End(xlUp).row

' Find the bar open time
Dim barOpen As String
barOpen = Worksheets(dataSheetName).Cells(my_row, 55)
If barOpen = "" Then
    MsgBox ("Import failed. Enter the bar opening time " _
            & "before attempting to import from Zettle")
    succeeded = False
    Exit Sub
Else
    ' Convert into a format VBA finds acceptable
    barOpen = Format(barOpen, "hh:mm:ss")
End If

' Find the bar close time
Dim barClosed As String
barClosed = Worksheets(dataSheetName).Cells(my_row, 56)
If barClosed = "" Then
    MsgBox ("Import failed. Enter the bar close time " _
            & "before attempting to import from Zettle")
    succeeded = False
    Exit Sub
Else
    ' Convert into a format VBA finds acceptable
    ' And add 10 minutes as a fudge-factor
    barClosed = DateAdd("n", 10, Format(barClosed, "hh:mm:ss"))
End If

' Find column with final price in
Dim priceAddress As Variant
priceAddress = funcs.search("Final price (GBP)", sheetName)
If priceAddress(0) = "0" Then
    MsgBox ("Import failed. Did you select the correct file?")
    succeeded = False
    Exit Sub
End If

' Store revenue
Dim revenue As Double
revenue = 0

' Run down the date column to find the first and last row which
'   are within the time range
Dim i As Integer ' index
Dim iTime As String ' time we're currently inspecting
Dim iDate As String ' date we're currently inspecting

For i = dateAddress(0) + 1 To FinalRow
    ' The value from the worksheet comes in as a date followed by a space, then a time
    ' First split the value using a space as a delimiter.
    ' Then take the second item in the array generated.
    ' Now format the resulting time to something acceptable to be turned into a TimeValue
    iTime = Format(Split(Worksheets(sheetName).Cells(i, 1), " ")(1), "hh:mm:ss")
    iDate = Split(Worksheets(sheetName).Cells(i, 1), " ")(0)
    
    ' Check if the transaction is within the time range and on the right date.
    If TimeValue(iTime) >= TimeValue(barOpen) _
        And TimeValue(iTime) <= TimeValue(barClosed) And _
        Worksheets(dataSheetName).Cells(my_row, 3) = iDate Then
        revenue = revenue + CDbl(Worksheets(sheetName).Cells(i, CInt(priceAddress(1))))
    Else
    End If
Next

' Put revenue into data sheet
Worksheets(dataSheetName).Cells(my_row, 25) = revenue

' Update pivot table(s)
Call ChangeSource("Data", "Analysis", "PivotTable1")
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

Private Function ChangeSource(dataSheetName As String, pivotSheetName As String, pivotName As String) As Boolean
'PURPOSE: Automatically readjust a Pivot Table's data source range
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault
' NOTE: do NOT select "Add this data to the data model" when creating the pivot table.

Dim Data_Sheet As Worksheet
Dim Pivot_Sheet As Worksheet
Dim StartPoint As Range
Dim DataRange As Range
Dim newRange As String
Dim LastCol As Long
Dim lastRow As Long
Dim DownCell As Long

'Set Pivot Table & Source Worksheet
Set Data_Sheet = ThisWorkbook.Worksheets(dataSheetName)
Set Pivot_Sheet = ThisWorkbook.Worksheets(pivotSheetName)

'Dynamically Retrieve Range Address of Data
Set StartPoint = Data_Sheet.Range("A1")
LastCol = StartPoint.End(xlToRight).Column
DownCell = StartPoint.End(xlDown).row
Set DataRange = Data_Sheet.Range(StartPoint, Data_Sheet.Cells(DownCell, LastCol))
'Set DataRange = Data_Sheet.Range(StartPoint, Cells(42, 46))

newRange = Data_Sheet.name & "!" & DataRange.address(ReferenceStyle:=xlR1C1)

'Change Pivot Table Data Source Range Address
Pivot_Sheet.PivotTables(pivotName). _
ChangePivotCache ActiveWorkbook. _
PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange)

 'Ensure Pivot Table is Refreshed
Pivot_Sheet.PivotTables(pivotName).RefreshTable

End Function

Private Sub AutofillFromSelected(row As Variant)
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
EventDateTextBox.Text = Worksheets(sheet).Cells(row, 3)
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

Private Sub RefreshListBox(sourceSheet As String, sourceColumn As Integer, list As Control)
' Will show column header if column is empty (apart from the header)
' Not fixing it for now because it doesn't seem to matter.

Dim empty_row As Long ' Store number of items in list box
Dim DataRange As Range
Dim myIndex As Long
myIndex = list.ListIndex

' empty_row = lst non-empty row for specific list(box)
empty_row = Worksheets(sourceSheet).Cells(Rows.Count, 1).End(xlUp).row
Set DataRange = Range(Worksheets(sourceSheet).Cells(2, sourceColumn), _
                Worksheets(sourceSheet).Cells(empty_row, sourceColumn))
list.RowSource = DataRange.address(External:=True)
list.ListIndex = myIndex
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
If EventDateTextBox = "" Then
    Exit Sub
End If

' Find out day event is on
Dim day As String
day = Format(EventDateTextBox.Text, "dddd")

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
    If TypeOf Ctrl Is MsForms.TextBox Then
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
