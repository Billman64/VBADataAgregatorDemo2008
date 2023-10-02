Sub btnConsolidateData()
'
' Data Consolidation macro
' Coded by: Bill Lugo
' Date: 2008?
'
' Description: Looks through records, collecting certain data points, and displays an aggregate report on a separate sheet.
' Was originally made for a sheet filled with tabular data.
'

Const OUTPUTSHEET As String = "(macro output)"
Const MAXROW As Single = 100 '27000 '98869
Dim sngRow As Single
Dim sngRowS2 As Single
sngRowS2 = 3  ' row # for output shees, also # of unique customers +3
Dim strVal As String
Dim strInputSheet As String
strInputSheet = ActiveSheet.Name


' error trapping - checking for existing output sheet
Application.StatusBar = ""
strVal = "ok"
For sngRow = 1 To Sheets.Count
    If Sheets(sngRow).Name = OUTPUTSHEET Then strVal = "output sheet already exists.  delete '(macro output)' and try again."
Next sngRow
' TODO: concantenate constant to re-use output sheet name

If strVal <> "ok" Then
    Application.StatusBar = "Macro Error:  " & strVal
    Exit Sub
End If  ' end of error trapping

' create a new sheet to output everything
    Sheets(strInputSheet).Select
    Sheets.Add
    ActiveSheet.Name = OUTPUTSHEET
    
' header
'Sheets(OUTPUTSHEET).Range("B3:AI3").Value = Sheets(strInputSheet).Range("B3:AI3").Value
Columns("B:AI").AutoFit
Rows(3).Font.Bold = True
Range("G4").Select
ActiveWindow.FreezePanes = True

' header - column headers for dates going across
Dim intCol
Dim intDate
intDate = 39721
For intCol = 10 To 32 Step 2
    Cells(2, intCol).Value = CStr(intDate)
    Cells(3, intCol).Value = "cost"
    Cells(2, intCol + 1).Value = CStr(intDate)
    Cells(3, intCol + 1).Value = " min.'s"
    intDate = intDate + 30
Next intCol
Range("C2:AG2").NumberFormat = "[$-409]mmm-yy;@"
Range("AH3").Value = "Total Min.'s"
Range("AI3").Value = "Total Cost"
    
' loop through each row to prepare to gather data


sngRow = 5
While sngRow <= MAXROW
        
    ' test if current row is the same customer (id#)
    If Sheets(strInputSheet).Range("F" & CStr(sngRow)).Value = _
        Sheets(strInputSheet).Range("F" & CStr(sngRow - 1)).Value Then
        
    
    
    Else   ' unique customer (same id#)
    
        sngRowS2 = sngRowS2 + 1 ' increment 2nd sheet row counter
    
        ' copy over data (col B-H), single-line qualitative data
        Sheets(OUTPUTSHEET).Range("B" & CStr(sngRowS2) & ":F" & CStr(sngRowS2)).Value = _
            Sheets(strInputSheet).Range("B" & CStr(sngRow) & ":F" & CStr(sngRow)).Value
            
        ' formulas to total mins and amounts
        ' mins
        Sheets(OUTPUTSHEET).Range("AH" & sngRowS2).Formula = _
            "=K" & sngRowS2 & "+M" & sngRowS2 & "+O" & sngRowS2 & _
            "+Q" & sngRowS2 & "+S" & sngRowS2 & "+U" & sngRowS2 & _
            "+W" & sngRowS2 & "+Y" & sngRowS2 & "+AA" & sngRowS2 & _
            "+AC" & sngRowS2 & "+AE" & sngRowS2 & "+AG" & sngRowS2
        
        ' amt's
        Sheets(OUTPUTSHEET).Range("AI" & sngRowS2).Formula = _
            "=J" & sngRowS2 & "+L" & sngRowS2 & "+N" & sngRowS2 & _
            "+P" & sngRowS2 & "+R" & sngRowS2 & "+T" & sngRowS2 & _
            "+V" & sngRowS2 & "+X" & sngRowS2 & "+Z" & sngRowS2 & _
            "+AB" & sngRowS2 & "+AD" & sngRowS2 & "+AF" & sngRowS2
    
    End If


        ' based on month (col I), output each cost & min. (consolidated into same line)
        
        strVal = Sheets(strInputSheet).Range("G" & CStr(sngRow)).Value  ' get the month into a variable
        Select Case strVal  ' case mapping input month to output column
            Case "2008/09"
                Sheets(OUTPUTSHEET).Range("J" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Sept. cost
                Sheets(OUTPUTSHEET).Range("K" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Sept. mins
            Case "2008/10"
                Sheets(OUTPUTSHEET).Range("L" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Oct. cost
                Sheets(OUTPUTSHEET).Range("M" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Oct. mins
            Case "2008/11"
                Sheets(OUTPUTSHEET).Range("N" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Nov. cost
                Sheets(OUTPUTSHEET).Range("O" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Nov. mins
            Case "2008/12"
                Sheets(OUTPUTSHEET).Range("P" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Dec. cost
                Sheets(OUTPUTSHEET).Range("Q" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Dec. mins
            Case "2009/01"
                Sheets(OUTPUTSHEET).Range("R" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Jan. cost
                Sheets(OUTPUTSHEET).Range("S" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Jan. mins
            Case "2009/02"
                Sheets(OUTPUTSHEET).Range("T" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Feb. cost
                Sheets(OUTPUTSHEET).Range("U" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Feb. mins
            Case "2009/03"
                Sheets(OUTPUTSHEET).Range("V" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Mar. cost
                Sheets(OUTPUTSHEET).Range("W" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Mar. mins
            Case "2009/04"
                Sheets(OUTPUTSHEET).Range("X" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Apr. cost
                Sheets(OUTPUTSHEET).Range("Y" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Apr. mins
            Case "2009/05"
                Sheets(OUTPUTSHEET).Range("Z" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' May. cost
                Sheets(OUTPUTSHEET).Range("AA" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' May. mins
            Case "2009/06"
                Sheets(OUTPUTSHEET).Range("AB" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' June cost
                Sheets(OUTPUTSHEET).Range("AC" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' June mins
            Case "2009/07"
                Sheets(OUTPUTSHEET).Range("AD" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Jul. cost
                Sheets(OUTPUTSHEET).Range("AE" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Jul. mins
            Case "2009/08"
                Sheets(OUTPUTSHEET).Range("AF" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AG" & CStr(sngRow)).Value  ' Aug. cost
                Sheets(OUTPUTSHEET).Range("AG" & CStr(sngRowS2)).Value = Sheets(strInputSheet).Range("AF" & CStr(sngRow)).Value  ' Aug. mins
            Case ""
                'Sheets(OUTPUTSHEET).Range("AH" & CStr(sngRowS2)).Value = 0
                'Sheets(OUTPUTSHEET).Range("AI" & CStr(sngRowS2)).Value = 0
            Case Else
                'Sheets(OUTPUTSHEET).Range("AH" & CStr(sngRowS2)).Value = " invalid/unrecognized date"
        End Select
    

    
    If sngRow Mod 1000 = 0 Then Application.StatusBar = Format(sngRow / MAXROW, "###%") ' update on % done

'Next sngRow
    sngRow = sngRow + 1
Wend



'Dress up output a bit
Columns("AH:AI").Font.Size = 14
Columns("A:AI").AutoFit
Range("AH1").Value = "Aggregated Data"
Range("AI3").Select
Range("AH" & CStr(sngRowS2) & ":AI" & CStr(sngRowS2)).Value = ""
Range("A1").Value = CStr(sngRowS2)
Application.StatusBar = "Data consolidation macro complete"
End Sub
