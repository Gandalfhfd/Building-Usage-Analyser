Attribute VB_Name = "Import_CSV"
Option Explicit

Sub csv_Import()
Dim wsheet As Worksheet, file_mrf As String
Set wsheet = ActiveWorkbook.Sheets("Single")
file_mrf = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Provide Text or CSV File:")
With wsheet.QueryTables.Add(Connection:="TEXT;" & file_mrf, Destination:=wsheet.Range("B2"))
.TextFileParseType = xlDelimited
.TextFileCommaDelimiter = True
.Refresh
End With
End Sub
 
