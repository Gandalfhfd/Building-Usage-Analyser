Attribute VB_Name = "Import"
Option Explicit

Public Function csv_Import(sheetName As String) As Boolean
' Declare stuff
Dim wsheet As Worksheet, file_mrf As String
Set wsheet = ActiveWorkbook.Sheets(sheetName)

' Open file explorer and let the user select a csv
file_mrf = Application.GetOpenFilename("CSV (*.csv),*.csv", , "Provide Text or CSV File:")

' Prevent it from crashing if the user doesn't select a file
If file_mrf <> "False" Then
    ' Clear sheet
    Sheets(sheetName).Cells.Clear
    ' Import file into sheet
    With wsheet.QueryTables.Add(Connection:="TEXT;" & file_mrf, Destination:=wsheet.Cells)
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
    ' Set exit code
    csv_Import = True
Else
    csv_Import = False
End If
End Function

Public Function xlsx_Import(sheetName As String) As Boolean
' Can actually import any Excel workbook
' Only imports the first sheet into a workbook called "Fantasy Spreadsheet"

' Declare stuff
Dim wsheet As Worksheet, file_mrf As String
Set wsheet = ActiveWorkbook.Sheets(sheetName)

' Open file explorer and let the user select an xlsx. This just gets the file name & path.
file_mrf = Application.GetOpenFilename("Excel Workbooks (*.xl??),*.xl??", , "Provide xlsx File:")

' If no file is selected, exit function
If file_mrf = "False" Then
    Exit Function
End If

' Find file name (not path)
Dim file_name As String
Dim pathArr As Variant
' Split file_mrf into sections delimited by "\"
pathArr = Split(file_mrf, "\")
' Find file name including extension
file_name = pathArr(UBound(pathArr))

' Detect if the workbook is open
' Not very reliable.
Dim workbook_open As Boolean
workbook_open = IsWorkBookOpen(file_mrf)

If file_mrf <> "False" Then
    ' Clear sheet
    Sheets(sheetName).Cells.Clear
    If workbook_open = False Then
        ' Open workbook
        Workbooks.Open (file_mrf)
    End If
    
    ' Import data from worksheet 1
    Workbooks(file_name).Worksheets(1).Cells.Copy _
        Destination:=Workbooks("Fantasy Spreadsheet").Worksheets(sheetName).Cells
    
    ' If it wasn't open to start with, we should close it
    If workbook_open = False Then
        Workbooks(file_name).Close
    End If
    xlsx_Import = True
Else
    xlsx_Import = False
End If

End Function

Public Sub ImportFromTicketsolve(mode As String)
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
    MsgBox ("Programmer error in Private Sub Import.ImportFromTicketsolve")
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
csvImportSuccessCheck = Import.csv_Import(sheetName)
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
Call funcs.ChangeSource(dataSheetName, "Analysis", "PivotTable1")

' Update autofilled info, if applicable
If AutofillCheckBox = True Then
    AutofillCheckBox = False
    AutofillCheckBox = True
End If

' Must go at the bottom
If succeeded = True Then
    MsgBox ("Import successful")
End If
End Sub

Public Sub ImportFromZettle(mode As String)
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
    MsgBox ("Programm   er error in Private Sub Import.ImportFromZettle")
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
xlsxImportSuccessCheck = Import.xlsx_Import(sheetName)
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

' Update autofilled info, if applicable
If AutofillCheckBox = True Then
    AutofillCheckBox = False
    AutofillCheckBox = True
End If

If revenue <> 0 Then
    MsgBox ("Import successful")
Else
    MsgBox ("Import could not find any transactions while the bar was open. " _
            & "Did you import the correct file?")
End If

' Update pivot table(s)
Call funcs.ChangeSource("Data", "Analysis", "PivotTable1")
End Sub
