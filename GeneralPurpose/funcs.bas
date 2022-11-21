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

' Search for stuff
Dim c As Range
Dim R1C1address As String ' Address in R1C1 form
Dim myAddress As Variant ' Address as array

' Find event capacity

'Find total sales

With ActiveWorkbook.Worksheets(sheetName).Range("A:Z") ' Look in worksheet
    Set c = .Find(word, LookIn:=xlValues)
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

