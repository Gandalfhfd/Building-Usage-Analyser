Attribute VB_Name = "InptValid"
Public Function RmSpecialChars(inputStr As String) As String
' List of chars we want to remove
Const SpecialCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?, ,/,:,',.,£"
Dim char As Variant

RmSpecialChars = inputStr

' Iterate over SpecialCharacters and remove everything that matches
For Each char In Split(SpecialCharacters, ",")
    RmSpecialChars = Replace(RmSpecialChars, char, "")
Next

' Remove commas
RmSpecialChars = Replace(RmSpecialChars, ",", "")

End Function

Public Function RmCommas(inputStr As String) As String
RmCommas = Replace(inputStr, ",", "")
End Function

Public Function RmPound(inputStr As String) As String
RmPound = Replace(inputStr, "£", "")
End Function

Public Function CheckIfNonNegInt(inputStr As String) As Boolean
If inputStr = "" Then ' If blank, ignore
    CheckIfNonNegInt = True
ElseIf IsNumeric(inputStr) = False Then ' Check it can be conveted to a number
    CheckIfNonNegInt = False
ElseIf Round(CDbl(inputStr)) <> CDbl(inputStr) Then ' Check it is an integer
    CheckIfNonNegInt = False
ElseIf CDbl(inputStr) < 0 Then ' Check it is >= 0. CDbl used to prevent overflow.
    CheckIfNonNegInt = False
Else ' Then it must be a non-negative integer
    CheckIfNonNegInt = True
End If
End Function

Public Function CheckIfNum(inputStr As String) As Boolean
If inputStr = "" Then ' If blank, ignore
    CheckIfNum = True
ElseIf inputStr = "-" Then
    ' To allow the user to write negative numbers, despite them not being numbers.
    ' Must be careful to not accept this as input though.
    CheckIfNum = True
ElseIf Len(inputStr) > 1 And Right(inputStr, 1) = "-" Then
    ' If the final character is a "-" and there's more than one character, don't allow this.
    CheckIfNum = False
ElseIf IsNumeric(inputStr) = False Then ' Check it can be converted into a number
    CheckIfNum = False
Else ' Then it must be a number
    CheckIfNum = True
End If
End Function

' select number
' select <, <=, >, >= with string input
' take in original text, modified text
' sanitise percentage, but any greater or lesser than anything

Public Function SanitiseNonNegInt(ByRef TextBoxName As Control, ByRef variableName As String) As Boolean
' Purpose: ensure the user can only enter non-negative integers into the desired text box.
' Input:
' TextBoxName = name of text box whose input we want to sanitise.
' variableName = name of public variable which stores old value of text box
'
' Output:
' True if input was valid
' False if input was invalid

If CheckIfNonNegInt(TextBoxName.value) = False Then
    TextBoxName.value = variableName ' Revert text box to previous valid text
    SanitiseNonNegInt = False ' Return value
Else
    TextBoxName.value = RmSpecialChars(TextBoxName.value) ' Remove commas and full stops/decimal points
    variableName = TextBoxName.value ' Update variable storing valid text
    SanitiseNonNegInt = True ' Return value
End If
End Function

Public Function SanitiseReal(ByRef TextBoxName As Control, ByRef variableName As String) As Boolean
' Purpose: ensure the user can only enter real numbers into the desired text box.
' Input:
' TextBoxName = name of text box whose input we want to sanitise.
' variableName = name of public variable which stores old value of text box
'
' Output:
' True if input was valid
' False if input was invalid

If CheckIfNum(TextBoxName.value) = False Then
    TextBoxName.value = variableName ' Revert text box to previous valid text
    SanitiseReal = False ' Return value
Else
    TextBoxName.value = RmCommas(TextBoxName.value) ' Remove commas
    TextBoxName.value = RmPound(TextBoxName.value) ' Remove £ sign
    variableName = TextBoxName.value ' Update variable storing valid text
    SanitiseReal = True ' Return value
End If
End Function

Public Function SanitisePercentage(ByRef TextBoxName As Control, ByRef variableName As String) As Boolean
' Purpose: sanitise input so that not only is the input real, it's also <= 100
' Input:
' TextBoxName = name of text box whose input we want to sanitise.
' variableName = name of public variable which stores old value of text box
'
' Output:
' True if input was valid
' False if input was invalid

If TextBoxName.value = "" Then ' We don't mind this, but it will confuse the > 100 check
    variableName = TextBoxName.value ' Update variable storing valid text
    SanitisePercentage = True ' Return value
ElseIf CheckIfNum(TextBoxName.value) = False Then
    TextBoxName.value = variableName ' Revert text box to previous valid text
    SanitisePercentage = False ' Return value
ElseIf TextBoxName.value > 100 Then ' Invalid, so revert
    TextBoxName.value = variableName ' Revert text box to previous valid text
    SanitisePercentage = False ' Return value
Else ' So it's <= 100 and we're happy
    variableName = TextBoxName.value ' Update variable storing valid text
    SanitisePercentage = True ' Return value
End If
End Function

Public Function Sanitise24Hr(ByRef TextBoxName As Control, ByRef variableName As String) As Boolean

Dim arr As Variant

' Means the user doesn't have to add the ":" in every time
If Len(TextBoxName.value) = 4 And CheckIfNonNegInt(TextBoxName.value) Then
    ReDim arr(1) As Variant
    arr(0) = Left(TextBoxName.value, 2)
    arr(1) = Right(TextBoxName.value, 2)
ElseIf Len(TextBoxName.value) = 3 And CheckIfNonNegInt(TextBoxName.value) Then
    ReDim arr(1) As Variant
    arr(0) = "0" & Left(TextBoxName.value, 1)
    arr(1) = Right(TextBoxName.value, 2)
ElseIf Len(TextBoxName.value) = 2 And CheckIfNonNegInt(TextBoxName.value) Then
    ReDim arr(1) As Variant
    arr(0) = TextBoxName.value
    arr(1) = "00"
ElseIf Len(TextBoxName.value) = 1 And CheckIfNonNegInt(TextBoxName.value) Then
    ReDim arr(1) As Variant
    arr(0) = "0" & TextBoxName.value
    arr(1) = "00"
Else
    arr = Split(TextBoxName.Text, ":")
End If

If TextBoxName.value = "" Then
    ' This is ok
    variableName = ""
    Exit Function
ElseIf ArrLen(arr) <> 2 Then ' The user has not put in exactly one ":"
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
ElseIf Len(arr(0)) > 2 Or Len(arr(1)) > 2 Then ' More than 2 chars before or after the :
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
ElseIf CheckIfNonNegInt(CStr(arr(0))) = False Then ' not a non negative integer before the :
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
ElseIf CheckIfNonNegInt(CStr(arr(1))) = False Then ' not a non negative integer after the :
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
ElseIf Len(arr(0)) < 2 Or Len(arr(1)) < 2 Then ' less than 2 chars before or after the :
    Dim i As Integer
    Dim j As Integer
    For j = 0 To 1
        For i = 1 To 2 - Len(arr(j))
            arr(j) = "0" & arr(j)
        Next
    Next
    variableName = arr(0) & ":" & arr(1)
    TextBoxName.value = variableName
End If

' Has to be separate in order to be called if Len(arr(0)) < 2 Or Len(arr(1)) < 2.
If CDbl(arr(0)) > 23 Then ' number is bigger than 23 before the :
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
ElseIf CDbl(arr(1)) > 59 Then ' number is bigger than 59 after the :
    MsgBox ("Please enter time in standard 24hr 'hh:mm' format. For example 02:08 for 8 minutes past 2 in the morning")
    TextBoxName.value = variableName
    Exit Function
Else
    ' Nothing further needs to change
    variableName = arr(0) & ":" & arr(1)
    TextBoxName.value = variableName
End If



End Function
