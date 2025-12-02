Attribute VB_Name = "Module2"
Sub CategorizeVerificationOfficers()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim officerName As String
    
    ' Set your Pivot Table sheet
    Set ws = ThisWorkbook.Sheets("PivotTable")
    
    ' Find last row in column A (assuming officers are in column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each officer
    Dim i As Long
    For i = 2 To lastRow ' Start from row 2, assuming row 1 is header
        
        officerName = Trim(ws.Cells(i, "A").Value)
        
        ' Assign category based on officer name
        Select Case officerName
            ' ADHOC Officers
            Case "Abubakar Ishaq", "DEBORAH FAITH OMOSEYE", "Ibrahim Isyaku", "Otukhagua Benjamin", "Ojeikhoa Emmanuel Abiodun", _
                      "Shitu Aliyu", "Sani Abdurahman", "Ubangida Abdurrahman", "PETER STEPHEN", "Sani Aminu", _
                      "Yahaya Ibrahim", "Andrew Vawe", "Shuaibu Saidu", "Yusuf Nasirdeen", "Usman Yahaya", "suleiman abdullahi", _
                      "Mahmud   Ibrahim", "Yakubu Ibrahim"





                ws.Cells(i, "C").Value = "ADHOC"
            
            ' HO Officers
            Case "UDUMA EMMANUEL", "DANLANDI EUGENE", "AMINU AUDU", "PETER DANKARO", "INNOCENT  SIMON", _
                           "BUKAR LAWAL", "TOYOSI ADEYEMI", "SOMI KADIRI", "HARUNA YELWA", "HARUNA ABDULLAHI", _
                           "GABRIEL GEORGE", "GIDEON DANIEL", "RAKIYA MUSA", "SULE UMARU", "HASHIMA ADAM", _
                           "DANLADI EUGENE", "Abdullahi Haruna", "AUSTIN MODI", "ABDULSALAR SULEIMAN", "SULEIMAN AMBALI", _
                           "ABDULSALAM SULEIMAN"
              
              ws.Cells(i, "C").Value = "HO"
            
            
            ' Rider Officers
            Case "ABUBAKAR ISAH", "ABUBAKAR SARKI", "ALIYU MUHAMMED", "OLOCHE NGBEDE", "Yusuf Sunday", _
                       "HASSAN SULEIMAN", "Ibrahim Mohammed", "IBRAHIM UMAR", "IKELLE NAILS", "James Godwin", _
                       "KURAH IBRAHIM", "Mark Misi", "PAUL IKYOOR", "Sadiq Maidugu", "Suleiman Abdullahi", "Sadiq Mijinyawa", _
                       "KUM TIMOTHY", "MATHIAS GBASONGON", "Musa Wada", "RAYMOND SAMUEL", "PHILIP OLORUNSAIYE", _
                       "EMMANUEL DANLADI", "UMAR IBN ISAH", "ABDULLAHI AHMED USMAN", "Yusuf Isah", "Rotimi Elisha", "Christian Odeh", _
                       "ALEXANDER MATHIAS", "BINCHAK BINTUR NANKPAK", "KAPCHANG DANLADI", "Hassan Ali Gambo", "AMANG FRANCIS FELIX", _
                       "SANI ALHAJI BAKARI", "Umoru Odilihi", "ZAHARADDEEN ADAMU", "Usman Buhari", "LAWAL LURWANU", "Kum Timothy"


                ws.Cells(i, "C").Value = "RIDER"
            
            ' If not found
            Case Else
                ws.Cells(i, "C").Value = ""
        End Select
        
    Next i
    
    MsgBox "Officer categorization complete.", vbInformation

End Sub


