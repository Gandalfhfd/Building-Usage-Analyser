Attribute VB_Name = "funcs"
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

Public Function AddAllToListBox(word As String, sheet As Worksheet, searchColumns As Range, _
                                listColumns As Variant, list As Control, exact As Boolean) _
                                As Boolean
' Find all entries matching word, then add them to the listbox called list
' Data from the row where the match is found is shown in the listbox, according to the
'   listColumns array.
' Inputs:
' word = text you're searching for
' sheet = worksheet you're looking on
' searchColumns = the columns you're searching over. Blank if you want to look at the entire sheet.
' listColumns = which columns from the sheet should appear in the list
' list = the listbox we will be writing to
' exact = whether the match must be perfect or not. True if it must be perfect.
'   False if it needn't be

' Output:
' boolean which says whether a match was found.
' True if one was, False if one wasn't

If word = "" Then
    Exit Function
End If

Dim c As Range
Dim firstAddress As String

Dim newRange As Range

Dim i As Integer
With searchColumns
    Set c = .Find(word, LookIn:=xlValues)
    If Not c Is Nothing Then
        firstAddress = c.address
        Do
            For i = 0 To ArrLen(listColumns) - 1
                ' Add items to listbox
                list.AddItem (c.value)
            Next
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

Public Sub IsWorkBookOpen(FileName As String)
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
End Sub

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
