Attribute VB_Name = "funcs"
Option Explicit

Function search(word As String, sheetName As String) As Variant
' Search for word in current workbook, and sheetName sheet.
' Output location, if found.
' If not found, output (0,0)

' Output: 1D array with 2 values.
' First is row where item was found. Second is column

' Search for stuff
Dim c As Range
Dim R1C1Address As String ' Address in R1C1 form
Dim myAddress As Variant ' Address as array

If word = "" Then
    ' Impossible address. Means nothing found.
    ReDim myAddress(1) As Variant
    myAddress(0) = "0"
    myAddress(1) = "0"
    ' Give function an output
    search = myAddress
    Exit Function
End If

With ActiveWorkbook.Worksheets(sheetName).Cells ' Look in worksheet
    ' This does the searching.
    '   xlValues says we're looking at the values of the cells, as opposed to comments, say.
    '   xlWhole means exact match,
    '   so a search of "e", for example, wouldn't turn up everything'
    '   in the sheet which contains an "e"
    Set c = .Find(What:=word, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then ' If anything is found, then...
        ' Give address in R1C1 form
        R1C1Address = c.address(ReferenceStyle:=xlR1C1)
        ' Convert R1C1 into array
        myAddress = StrManip.SplitR1C1(R1C1Address)
    Else
        ' Impossible address. Means nothing found.
        ReDim myAddress(1) As Variant
        myAddress(0) = "0"
        myAddress(1) = "0"
    End If
End With

' Give function an output
search = myAddress

End Function

Public Function AddAllToListBox(word As String, searchRange As Range, _
                                listColumns As Variant, nameList As Control, dateList As Control, _
                                typeList As Control, IDList As Control, dataSheet As String) _
                                As Boolean
' Find all entries matching word, then add them to the listbox called list
' Data from the row where the match is found is shown in the listbox, according to the
'   listColumns array.
' Inputs:
' word = text you're searching for
' searchRange = the range over which you're searching, including the worksheet.
' listColumns = which columns from the sheet should appear in the list
' list = the listbox we will be writing to
' dataSheet = the name of the sheet we're pulling data from

' Output:
' boolean which says whether a match was found.
' True if one was, False if one wasn't

If word = "" Then
    Exit Function
End If

Dim c As Range
Dim firstAddress As String

' Store address of cell containing word
Dim resultAddress As Variant

Dim i As Integer
With searchRange
    Set c = .Find(word, LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        firstAddress = c.address
        Do
            ' Find address
            resultAddress = StrManip.SplitR1C1(c.address(ReferenceStyle:=xlR1C1))
            ' Add items to listbox
            nameList.AddItem (Worksheets(dataSheet).Cells(resultAddress(0), listColumns(0)))
            dateList.AddItem (Format(Worksheets(dataSheet).Cells(resultAddress(0), listColumns(1)), _
                                "dd/mm/yyyy"))
            typeList.AddItem (Worksheets(dataSheet).Cells(resultAddress(0), listColumns(2)))
            ' Add item to Event ID listbox. This listbox is hidden, but links the events on
            '   this page to the data.
            IDList.AddItem (Worksheets(dataSheet).Cells(resultAddress(0), 1))
            
            Set c = .FindNext(c)
        Loop While firstAddress <> c.address
    End If
End With


End Function

Public Function ReverseArray(arr As Variant) As Variant
' Return a reversed array

' Regular index
Dim i As Integer

Dim j As Integer
j = UBound(arr)
' loop from the LBound of arr to the midpoint of arr
Dim temp As Variant
For i = LBound(arr) To ((UBound(arr) - LBound(arr) + 1) \ 2)
    'swap the elements
    temp = arr(i)
    arr(i) = arr(j)
    arr(j) = temp
    ' decrement the upper index
    j = j - 1
Next

ReverseArray = arr
End Function

Public Function ArrLen(arr As Variant) As Integer
ArrLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function max(x, y As Variant) As Variant
' Find max of two numbers
  max = IIf(x > y, x, y)
End Function

Public Function min(x, y As Variant) As Variant
' Find min of two numbers
   min = IIf(x < y, x, y)
End Function

Public Function IsWorkBookOpen(FileName As String)
    ' I don't know what most of this does. I got it from
    ' https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Function GenerateRandomInt(lowbound As Integer, upbound As Integer) As Integer
' Return a random integer between lowbound and upbound
GenerateRandomInt = Int((upbound - lowbound + 1) * Rnd + lowbound)
End Function

Function GenerateRandomAlphaNumericStr(length As Integer) As String
' Returns a string of length length of upper and lower case letters and numbers
' Will never return a zero, due to confusion between "0" and "O"
' length = length of string to be returned

' Decide what type of character to use
Dim decider As Single

' Store our string because it's shorter than the function name
Dim randomStr As String
randomStr = ""

' for loop index
Dim i As Integer

' Repeat length times
For i = 0 To length - 1
    ' Need new decision each round
    decider = Rnd
    If decider < 9 / 61 Then
        ' Generate a random integer between 1 and 9
        randomStr = randomStr & GenerateRandomInt(1, 9)
    ElseIf decider >= 9 / 61 And decider < 35 / 61 Then
        ' Generate a random upper case letter
        randomStr = randomStr & Chr(GenerateRandomInt(65, 90))
    ElseIf decider >= 35 / 61 Then
        ' Generate a random lower case letter
        randomStr = randomStr & Chr(GenerateRandomInt(97, 122))
    Else
        MsgBox ("GenerateRandomAlphaNumbericStr has failed. Contact support.")
    End If
Next

GenerateRandomAlphaNumericStr = randomStr
End Function

Public Function ChangeSource(dataSheetName As String, pivotSheetName As String, pivotName As String) As Boolean
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

Public Sub RefreshListBox(sourceSheet As String, sourceColumn As Integer, list As Control)
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

Public Function UUIDGenerator(category As String, eventDate As String, name As String) As String
' Generate uniqueish UUID.
' If name, category and date are all the same, there is a 1 in 844,596,301 change of a collision.
UUIDGenerator = InptValid.RmSpecialChars(name) & InptValid.RmSpecialChars(category) _
                & InptValid.RmSpecialChars(eventDate) & funcs.GenerateRandomAlphaNumericStr(5)
End Function

Sub GetCalendar(DateTextBox As Control) ' Calendar format
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
If dateVariable <> 0 Then DateTextBox = Format(dateVariable, "dd/mm/yyyy")
End Sub
