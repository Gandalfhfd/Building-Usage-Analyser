Attribute VB_Name = "StrManip"
Public Function ConvertDate(myDate As String) As String
' Designed to convert date stored as date into format Excel recognises

Dim dateArr As Variant
' Split date into arrays, using "/" as delimiter
dateArr = Split(myDate, "/")

' Reverse array
dateArr = ReverseArray(dateArr)

' Join it back together
ConvertDate = Join(dateArr, "-")
End Function

Public Function Convert2Currency(ByVal inputTextBox As Control) As String
' Convert inputTextBox.value into currency format. Adds "£" on,
'   but if SanitiseReal is called when inputTextBox gets changed, it will appear to not add "£"/.

If inputTextBox.value = "-" Then
    Convert2Currency = ""
    Exit Function
End If

Convert2Currency = Format(inputTextBox.value, "Currency")
End Function

Public Function SplitR1C1(address As String) As Variant
SplitR1C1 = Array("", "") ' Set up as array of strings

' Temporary array which stores address as separate bits
Dim parts As Variant

' Split the R1C1 format into an array of two strings. First is "Rx". Second is "y"
parts = Split(address, "C")
' parts(0) is "Rx". This starts parts(0) from the second character onwards
parts(0) = Mid(parts(0), 2) ' convert to integer

parts(1) = parts(1)
' Set function output to be separated address
SplitR1C1 = parts

End Function
