Attribute VB_Name = "Module3"
Function SumNumbersInCell(cell As Range) As Double
    Dim regex As Object              ' RegExp object to handle pattern matching
    Dim matches As Object            ' Collection of all number matches in the text
    Dim match As Variant             ' Temporary variable for iterating matches
    Dim total As Double              ' Sum accumulator

    ' Create and configure the regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True               ' Look for all matches, not just the first
        .Pattern = "\d+(\.\d+)?"     ' Match integers and decimals (e.g., 12, 34.5)
    End With

    ' If the cell contains any matches based on the regex pattern
    If regex.Test(cell.Value) Then
        Set matches = regex.Execute(cell.Value)
        ' Loop through each match and add it to the total sum
        For Each match In matches
            total = total + CDbl(match.Value)
        Next match
    End If

    ' Return the final sum
    SumNumbersInCell = total
End Function
