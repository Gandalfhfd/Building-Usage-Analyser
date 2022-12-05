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
' Capacity
Public AuditoriumCapacity As String
Public EgremontCapacity As String
Public TotalCapacity As String
Public BlockedSeats As String

' Revenue & Costs
Public NumTicketsSold As String
Public BoxOfficeRevenue As String
Public SupportRevenue As String ' Revenue from Support the Kirkgate donations
' Film
Public FilmCost As String
Public FilmTransport As String
Public Accommodation As String
Public ArtistFood As String
Public Heating As String
Public Lighting As String
Public MiscCost As String
' Bar
Public BarRevenue As String
Public BarMargin As String

' Time
' Event Time
Public SetupStartTime As String
Public DoorsTime As String
Public EventStartTime As String
Public EventEndTime As String
Public TakedownEndTime As String
Public EventDuration As String
Public SetupAvailableDuration As String
Public SetupTakedown As String
'Bar Time
Public BarSetupTime As String
Public BarOpenTime As String
Public BarCloseTime As String
Public BarOpenDuration As String
Public BarSetupToTakedownEndDuration As String
Public BarSetupTakedown As String

' Volunteer minutes
Public FoH As String
Public DM As String
Public Proj As String
Public BoxOffice As String
Public Bar As String
Public AoWVol As String
Public MiscVol As String

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

Private Sub EventButton3_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Needed because a double click is counted differently to a single click
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_Click()
Call AddEvent(EditToggleCheckBox1.value)
End Sub

Private Sub EventButton5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
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

' Update pivot table(s)
Call ChangeSource("Data", "Analysis", "PivotTable1")
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

Private Sub MultiPage1_Change()

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
SetupStartTimeTextBox.Text = ""
DoorsTimeTextBox.Text = ""
EventStartTimeTextBox.Text = ""
EventEndTimeTextBox.Text = ""
TakedownEndTimeTextBox.Text = ""
EventDurationTextBox.Text = ""
SetupToTakedownEndDurationTextBox.Text = ""
SetupTakedownTextBox.Text = ""
End Sub

'' TEXT BOXES===============================================================

' Basic Info============================================================================
Private Sub EventDateTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean) ' Re-show Date Picker if box has already been entered
Call GetCalendar ' Show Date Picker
End Sub

Private Sub EventDateTextBox_Enter()
Call GetCalendar ' Show Date Picker
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

Private Sub BlockedSeatsTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Stop it from being left blank
If BlockedSeatsTextBox.Text = "" Then
    BlockedSeatsTextBox.Text = 0
End If
End Sub

' Time============================================================================

Private Sub SetupStartTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(SetupStartTimeTextBox, SetupStartTime)

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

Private Sub DoorsTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(DoorsTimeTextBox, DoorsTime)

If DoorsTimeTextBox.Text = "" Or EventStartTimeTextBox.Text = "" Then
    Exit Sub
' Sanity check doors time
ElseIf TimeValue(EventStartTimeTextBox.Text) < TimeValue(DoorsTimeTextBox.Text) Then
    MsgBox ("Doors cannot open after event starts")
    DoorsTimeTextBox.Text = ""
End If

End Sub

Private Sub EventStartTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(EventStartTimeTextBox, EventStartTime)

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
EventEndTimeTextBox.Text = Format(DateAdd("n", EventDurationTextBox.Text, _
    EventStartTimeTextBox.Text), "hh:mm")

' Change takedown time
TakedownEndTimeTextBox.Text = Format(DateAdd("n", EventDurationTextBox.Text + _
    Worksheets(sheet).Cells(row, 12), EventStartTimeTextBox.Text), "hh:mm")
    
' BAR
' Change bar setup start time
BarSetupTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 15) - _
    Worksheets(sheet).Cells(row, 16), EventStartTimeTextBox.Text), "hh:mm")

' Change bar open time
BarOpenTimeTextBox.Text = Format(DateAdd("n", -Worksheets(sheet).Cells(row, 15), _
    EventStartTimeTextBox.Text), "hh:mm")
End Sub

Private Sub EventEndTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(EventEndTimeTextBox, EventEndTime)

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

Private Sub TakedownEndTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(TakedownEndTimeTextBox, TakedownEndTime)

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

Private Sub BarSetupTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarSetupTimeTextBox, BarSetupTime)
End Sub

Private Sub BarOpenTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
' Sanitise input to ensure only valid 24hr times are input. Format hh:mm
Call InptValid.Sanitise24Hr(BarOpenTimeTextBox, BarOpenTime)
End Sub

Private Sub BarCloseTimeTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
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

Private Sub BoxOfficeRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
BoxOfficeRevenueTextBox.Text = StrManip.Convert2Currency(BoxOfficeRevenueTextBox)
End Sub

Private Sub SupportRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(SupportRevenueTextBox, SupportRevenue)
End Sub

Private Sub SupportRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
SupportRevenueTextBox.Text = StrManip.Convert2Currency(SupportRevenueTextBox)
End Sub

Private Sub FilmCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmCostTextBox, FilmCost)
End Sub

Private Sub FilmCostTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
FilmCostTextBox.Text = StrManip.Convert2Currency(FilmCostTextBox)
End Sub

Private Sub FilmTransportTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(FilmTransportTextBox, FilmTransport)
End Sub

Private Sub FilmTransportTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
FilmTransportTextBox.Text = StrManip.Convert2Currency(FilmTransportTextBox)
End Sub

Private Sub AccommodationTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(AccommodationTextBox, Accommodation)
End Sub

Private Sub AccommodationTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
AccommodationTextBox.Text = StrManip.Convert2Currency(AccommodationTextBox)
End Sub

Private Sub ArtistFoodTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(ArtistFoodTextBox, ArtistFood)
End Sub

Private Sub ArtistFoodTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ArtistFoodTextBox.Text = StrManip.Convert2Currency(ArtistFoodTextBox)
End Sub

Private Sub HeatingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(HeatingTextBox, Heating)
End Sub

Private Sub HeatingTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
HeatingTextBox.Text = StrManip.Convert2Currency(HeatingTextBox)
End Sub

Private Sub LightingTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(LightingTextBox, Lighting)
End Sub

Private Sub LightingTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
LightingTextBox.Text = StrManip.Convert2Currency(LightingTextBox)
End Sub

Private Sub MiscCostTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(MiscCostTextBox, MiscCost)
End Sub

Private Sub MiscCostTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
MiscCostTextBox.Text = StrManip.Convert2Currency(MiscCostTextBox)
End Sub

Private Sub BarRevenueTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitiseReal(BarRevenueTextBox, BarRevenue)
End Sub

Private Sub BarRevenueTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
BarRevenueTextBox.Text = StrManip.Convert2Currency(BarRevenueTextBox)
End Sub

Private Sub BarMarginTextBox_Change()
' Sanitise input to ensure only real numbers are input
Call InptValid.SanitisePercentage(BarMarginTextBox, BarMargin)
End Sub

' Volunteer Minutes==================================================================
Private Sub FoHTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(FoHTextBox, FoH)
End Sub

Private Sub DMTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(DMTextBox, DM)
End Sub

Private Sub ProjTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(ProjTextBox, Proj)
End Sub

Private Sub BoxOfficeTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BoxOfficeTextBox, BoxOffice)
End Sub

Private Sub BarTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(BarTextBox, Bar)
End Sub

Private Sub AoWVolTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(AoWVolTextBox, AoWVol)
End Sub

Private Sub MiscVolTextBox_Change()
' Sanitise input to ensure only non-negative integers are input
Call InptValid.SanitiseNonNegInt(MiscVolTextBox, MiscVol)
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
' Event
EventDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 10)
SetupToTakedownEndDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 13)
SetupTakedownTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 9) + _
                            Worksheets(TypeDefaultsSheet).Cells(row, 12)
' Bar
BarOpenDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 15)
BarSetupToTakedownEndDurationTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 15) + _
    Worksheets(TypeDefaultsSheet).Cells(row, 16) + Worksheets(TypeDefaultsSheet).Cells(row, 17)
BarSetupTakedownTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 16) + _
                                Worksheets(TypeDefaultsSheet).Cells(row, 17)

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
ProjTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 5)
BoxOfficeTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 6)
BarTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 7)
AoWVolTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 8)
MiscVolTextBox.Text = Worksheets(TypeDefaultsSheet).Cells(row, 14)
End Sub

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

'' USERFORM/MULTIPAGE===============================================================

Private Sub UserForm_Initialize()

' Add items into listboxes based on cells in specified worksheets

Dim empty_row As Long ' Store number of items in list box
Dim DataRange As Range
Dim TypeSpecificDefaultsRange As Range
Dim NonSpecificDefaultsRange As Range

' empty_row = lst non-empty row for specific list(box)

' EVENT ID
empty_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row
Set DataRange = Range(Worksheets("Data").Cells(2, 1), Worksheets("Data").Cells(empty_row, 1))
EventIDListBox.RowSource = DataRange.address(External:=True)

'' CATEGORY
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 4).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 4), Worksheets("Non-Specific Defaults").Cells(empty_row, 4))
CategoryListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

'' TYPE
empty_row = Worksheets("Type-Specific Defaults").Cells(Rows.Count, 1).End(xlUp).row
Set TypeSpecificDefaultsRange = Range(Worksheets("Type-Specific Defaults").Cells(2, 1), Worksheets("Type-Specific Defaults").Cells(empty_row, 1))
TypeListBox.RowSource = TypeSpecificDefaultsRange.address(External:=True)

'' LOCATION
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 1).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 1), Worksheets("Non-Specific Defaults").Cells(empty_row, 1))
LocationListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

'' ROOM
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 2).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 2), Worksheets("Non-Specific Defaults").Cells(empty_row, 2))
RoomListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

'' GENRE
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 13).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 13), Worksheets("Non-Specific Defaults").Cells(empty_row, 13))
GenreListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

'' AUDIENCE
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 5).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 5), Worksheets("Non-Specific Defaults").Cells(empty_row, 5))
AudienceListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)
   
'' AUDITORIUM LAYOUT
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 6).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 6), Worksheets("Non-Specific Defaults").Cells(empty_row, 6))
AuditoriumLayoutListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

'' EGREMONT LAYOUT
empty_row = Worksheets("Non-Specific Defaults").Cells(Rows.Count, 8).End(xlUp).row
Set NonSpecificDefaultsRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 8), Worksheets("Non-Specific Defaults").Cells(empty_row, 8))
EgremontLayoutListBox.RowSource = NonSpecificDefaultsRange.address(External:=True)

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
' Generate uniqueish UUID. Not unique if the same event is added twice within a second
UUIDGenerator = InptValid.RmSpecialChars(name) & InptValid.RmSpecialChars(category) _
                & InptValid.RmSpecialChars(eventDate) & Format(Now, "ss")
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
    MultiPage1.value = 1
    Exit Function
ElseIf EventDateTextBox.Text = "" Then
    MsgBox ("Please select a date using the calendar. Double click on the text box to show the calendar.")
    MultiPage1.value = 1
    Exit Function
ElseIf CategoryListBox.ListIndex = -1 Then ' .ListIndex = -1 means nothing has been selected yet
    MsgBox ("Please select a category")
    MultiPage1.value = 1
    Exit Function
ElseIf TypeListBox.ListIndex = -1 Then
    MsgBox ("Please enter a type")
    MultiPage1.value = 1
    Exit Function
ElseIf LocationListBox.ListIndex = -1 Then
    MsgBox ("Please select a location")
    MultiPage1.value = 1
    Exit Function
ElseIf RoomListBox.ListIndex = -1 Then
    MsgBox ("Please select a room")
    MultiPage1.value = 1
    Exit Function
ElseIf AudienceListBox.ListIndex = -1 Then
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
ElseIf GenreListBox.Visible = True And GenreListBox.ListIndex = -1 Then
    MsgBox ("Please enter a genre")
    MultiPage1.value = 1
    Exit Function
ElseIf AuditoriumLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Auditorium")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
ElseIf EgremontLayoutListBox.ListIndex = -1 Then
    MsgBox ("Please enter a layout for the Egremont Room")
    MultiPage1.value = 2 ' Take the user to the layouts page
    Exit Function
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
Else ' The user is allowed to create a new event
End If

Dim my_row As Long
If mode = False Then
    ' Next available row
    my_row = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).row + 1
ElseIf mode = True Then
    ' Row corresponds to the row of the event selected by the user
    my_row = funcs.search(SearchBox.Text, "Data")(0)
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

' Layout & Capacity
Worksheets(sheet).Cells(my_row, 31) = EgremontLayoutListBox.value
Worksheets(sheet).Cells(my_row, 32) = AuditoriumLayoutListBox.value
Worksheets(sheet).Cells(my_row, 33) = TotalCapacityTextBox.Text
Worksheets(sheet).Cells(my_row, 34) = BlockedSeatsTextBox.Text
Worksheets(sheet).Cells(my_row, 45) = TotalCapacityTextBox.Text - BlockedSeatsTextBox.Text

' Time
Worksheets(sheet).Cells(my_row, 8) = TimeValue(SetupStartTimeTextBox)
Worksheets(sheet).Cells(my_row, 47) = TimeValue(DoorsTimeTextBox)
Worksheets(sheet).Cells(my_row, 9) = TimeValue(EventStartTimeTextBox)
Worksheets(sheet).Cells(my_row, 10) = TimeValue(EventEndTimeTextBox)
Worksheets(sheet).Cells(my_row, 11) = TimeValue(TakedownEndTimeTextBox)
Worksheets(sheet).Cells(my_row, 13) = EventDurationTextBox.Text
Worksheets(sheet).Cells(my_row, 12) = SetupToTakedownEndDurationTextBox.Text
Worksheets(sheet).Cells(my_row, 15) = SetupTakedownTextBox.Text

' Costs & Income
Worksheets(sheet).Cells(my_row, 14) = NumTicketsSoldTextBox.Text
Worksheets(sheet).Cells(my_row, 42) = BoxOfficeRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 44) = SupportRevenueTextBox.Text
Worksheets(sheet).Cells(my_row, 35) = FilmCostTextBox.Text
Worksheets(sheet).Cells(my_row, 36) = FilmTransportTextBox.Text
Worksheets(sheet).Cells(my_row, 37) = AccommodationTextBox.Text
Worksheets(sheet).Cells(my_row, 38) = ArtistFoodTextBox.Text
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
Worksheets(sheet).Cells(my_row, 52) = _
        "=RC[-8]+RC[-10]-RC[-1]-RC[-11]-RC[-12]-RC[-13]-RC[-14]-RC[-15]-RC[-16]-RC[-17]"
Worksheets(sheet).Cells(my_row, 53) = "=RC[-1]+RC[-26]"

' Volunteer Minutes
Worksheets(sheet).Cells(my_row, 18) = FoHTextBox.Text
Worksheets(sheet).Cells(my_row, 19) = DMTextBox.Text
Worksheets(sheet).Cells(my_row, 20) = ProjTextBox.Text
Worksheets(sheet).Cells(my_row, 21) = BoxOfficeTextBox.Text
Worksheets(sheet).Cells(my_row, 22) = BarTextBox.Text
Worksheets(sheet).Cells(my_row, 23) = AoWVolTextBox.Text
Worksheets(sheet).Cells(my_row, 49) = MiscVolTextBox.Text


' Update pivot table(s)
Call ChangeSource(sheet, "Analysis", "PivotTable1")

MsgBox ("Event Added")
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

AuditoriumCapacityTextBox.Locked = True
AuditoriumCapacityTextBox.value = "0"
End Function

Private Function EgremontNotUsed()
' To be called when Egremont room is not being used

Dim ListBoxRange As Range
Set ListBoxRange = Range(Worksheets("Non-Specific Defaults").Cells(2, 8), _
                    Worksheets("Non-Specific Defaults").Cells(2, 8))
EgremontLayoutListBox.RowSource = ListBoxRange.address(External:=True)

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
trueCapacity = Worksheets(dataSheetName).Cells(exportCell(0), 33) - Worksheets(dataSheetName).Cells(exportCell(0), 34)
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

' Highly non-general function/sub which imports from the "Report Excel" file from Zettle

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
    MsgBox ("Programmer error in Private Sub ImportFromZettle")
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
    
End Sub

Private Sub ZettleImportButton_Click()
'' Function not written yet
'Call ImportFromZettle("Selected")

' Update pivot table(s)
Call ChangeSource("Data", "Analysis", "PivotTable1")
End Sub

Private Function ChangeSource(dataSheetName As String, pivotSheetName As String, pivotName As String) As Boolean
'PURPOSE: Automatically readjust a Pivot Table's data source range
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault
' NOTE: do NOT select "Add this data to the data model" when creating the pivot table.

Dim Data_Sheet As Worksheet
Dim Pivot_Sheet As Worksheet
Dim StartPoint As Range
Dim DataRange As Range
Dim NewRange As String
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

NewRange = Data_Sheet.name & "!" & DataRange.address(ReferenceStyle:=xlR1C1)

'Change Pivot Table Data Source Range Address
Pivot_Sheet.PivotTables(pivotName). _
ChangePivotCache ActiveWorkbook. _
PivotCaches.Create(SourceType:=xlDatabase, SourceData:=NewRange)

 'Ensure Pivot Table is Refreshed
Pivot_Sheet.PivotTables(pivotName).RefreshTable

End Function
