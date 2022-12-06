Attribute VB_Name = "funcs"
Option Explicit

Public Function csv_Import(sheetName As String) As Boolean

' Declare stuff
Dim wsheet As Worksheet, file_mrf As String
Set wsheet = ActiveWorkbook.Sheets(sheetName)

' Open file explorer and let the user select a csv
file_mrf = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Provide Text or CSV File:")

' Prevent it from crashing if the user doesn't select a file
If file_mrf <> "False" Then
    ' Clear "Import" sheet
    Sheets("Import").Cells.Clear
    With wsheet.QueryTables.Add(Connection:="TEXT;" & file_mrf, Destination:=wsheet.Range("B2"))
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

