Attribute VB_Name = "funcs"
Option Explicit

Sub csv_Import(sheetName As String)

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
End If

End Sub

Function SplitR1C1(address As String) As Variant
SplitR1C1 = Array("", "") ' Set up as array of strings

' Temporary array which stores address as separate bits
Dim parts As Variant

' Split the R1C1 format into an array of two strings. First is "Rx". Second is "y"
parts = Split(address, "C")
' parts(0) is "Rx". This starts parts(0) from the second character onwards
parts(0) = Mid(parts(0), 2)

' Set function output to be separated address
SplitR1C1 = parts

End Function

Function search(word As String, sheetName As String) As Variant
' Search for word in current workbook, and sheetName sheet.
' Output location, if found.
' If not found, output (0,0)

' Output: 1D array with 2 values.
' First is row where item was found. Second is column

' Search for stuff
Dim c As Range
Dim R1C1address As String ' Address in R1C1 form
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

With ActiveWorkbook.Worksheets(sheetName).Range("A:Z") ' Look in worksheet
    Set c = .Find(word, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then ' If anything is found, then...
        ' Give address in R1C1 form
        R1C1address = c.address(ReferenceStyle:=xlR1C1)
        ' Convert R1C1 into array
        myAddress = funcs.SplitR1C1(R1C1address)
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

'Function SubtractFromRange(oldRange As Range, subtrahendRange As Range) As Range
'' look up subtrahend if you don't know what it means
'' Take set oldRange and set subtrahendRange. Do SubtractFromRange = oldrange - subtrahendRange.
'' Can return empty set
'
'' FUNCTION IS VERY SLOW DON'T USE
'
'Dim rCell As Range ' Store current cell being considered
'
'For Each rCell In oldRange
'    If Not IsEmpty(Intersect(rCell, subtrahendRange)) Then
'        ' If intersection is nonempty, then we don't want this cell, so don't add it to new range.
'    ElseIf SubtractFromRange Is Nothing Then
'        ' This is the first rCell and we want to add to the new range, so we have to do
'        '   something different because SubtractFromRange is currently empty.
'        Set SubtractFromRange = rCell
'    Else
'        ' We want rCell to be in the new range, so we add it on.
'        Set SubtractFromRange = Union(SubtractFromRange, rCell)
'    End If
'Next
'End Function