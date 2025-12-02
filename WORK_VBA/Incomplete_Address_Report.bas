Attribute VB_Name = "Module1"
Sub IncompleteAddressesReport()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim officer As Variant
    Dim newRow As Long
    Dim address As String
    Dim regexDigit As Object, regexLetter As Object
    Dim addressCol As String, officerCol As String
    Dim addressType As String
    Dim ptCache As PivotCache, pt As PivotTable
    
    ' Ask user which columns to use
    addressCol = InputBox("Enter the column letter that contains addresses:", "Address Column", "C")
    officerCol = InputBox("Enter the column letter that contains officer names:", "Officer Column", "L")
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in address column
    lastRow = ws.Cells(ws.Rows.Count, addressCol).End(xlUp).Row
    
    ' Delete old results sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsNew = ThisWorkbook.Sheets("IncompleteAddresses")
    If Not wsNew Is Nothing Then wsNew.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add a new sheet for results
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.name = "IncompleteAddresses"
    
    ' Add headers
    With wsNew
        .Range("A1").Value = "Incomplete Address"
        .Range("B1").Value = "Verification Officer"
        .Range("C1").Value = "Address Type"
        .Range("A1:C1").Font.Bold = True
    End With
    
    ' Start writing from row 2
    newRow = 2
    
    ' Create regex objects
    Set regexDigit = CreateObject("VBScript.RegExp")
    regexDigit.Pattern = "\d"   ' matches any digit
    regexDigit.IgnoreCase = True
    regexDigit.Global = True
    
    Set regexLetter = CreateObject("VBScript.RegExp")
    regexLetter.Pattern = "[A-Za-z]"   ' matches any letter
    regexLetter.IgnoreCase = True
    regexLetter.Global = True
    
    ' Loop through each address (skip header by starting at row 2)
    For Each cell In ws.Range(addressCol & "2:" & addressCol & lastRow)
        If Not IsEmpty(cell.Value) Then
            address = Trim(cell.Value)
            
            ' Case 1: No digits at all (text-only)
            ' Case 2: No letters at all (numbers-only)
            If (Not regexDigit.Test(address)) Or (Not regexLetter.Test(address)) Then
                
                ' Identify address type
                If Not regexDigit.Test(address) Then
                    addressType = "Text-only"
                ElseIf Not regexLetter.Test(address) Then
                    addressType = "Numbers-only"
                Else
                    addressType = "Other"
                End If
                
                ' Get the corresponding officer
                officer = Trim(ws.Cells(cell.Row, officerCol).Value)
                If officer = "" Then officer = "Unassigned"
                
                ' Add to new sheet
                wsNew.Cells(newRow, 1).Value = address
                wsNew.Cells(newRow, 2).Value = officer
                wsNew.Cells(newRow, 3).Value = addressType
                newRow = newRow + 1
            End If
        End If
    Next cell
    
    ' Autofit columns
    wsNew.Columns("A:C").AutoFit
    
    ' Create Pivot Table
    Dim ptStart As Range
    Dim dataRange As Range
    
    If newRow > 2 Then ' Only build pivot if we have data
        Set dataRange = wsNew.Range("A1").CurrentRegion
        Set ptStart = wsNew.Range("E3")
        
        Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
        
        On Error Resume Next
        Set pt = wsNew.PivotTables("IncompletePivot")
        On Error GoTo 0
        
        If pt Is Nothing Then
            Set pt = ptCache.CreatePivotTable(TableDestination:=ptStart, TableName:="IncompletePivot")
        End If
        
        With pt
            .ClearAllFilters
            .PivotFields("Verification Officer").Orientation = xlRowField
            .PivotFields("Address Type").Orientation = xlColumnField
            .AddDataField .PivotFields("Incomplete Address"), "Count of Incomplete Address", xlCount
        End With
    End If
    
    MsgBox "Incomplete addresses report generated with Pivot Table in 'IncompleteAddresses'.", vbInformation
End Sub

