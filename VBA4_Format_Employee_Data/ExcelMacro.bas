Option Explicit

Sub ReadAndFormatData()

    'Declare variables
    Dim dataSheet, resultSheet As Worksheet
    Dim i, n, a, b, currentResultRow, dataNumRows, numberOfGreyRows As Integer
    Dim previousJobNumber, currentJobNumber As String
    Dim uniqueNames, newUniqueNames As Variant
    
    Set dataSheet = Sheets("Data")
    Set resultSheet = Sheets("End Result")
    
    resultSheet.Cells.Clear
    
    resultSheet.Cells.Font.Size = 10
    
    'Create header
    resultSheet.Cells(5, 1) = "NEW JOB NUMBER"
    resultSheet.Cells(5, 2) = "JOB DESCRIPTION"
    resultSheet.Cells(5, 3) = "OLD JOB NUMBER"
    resultSheet.Cells(5, 4) = "ENTRY DATE"
    resultSheet.Cells(5, 5) = "EMPLOYEE NAME"
    resultSheet.Cells(5, 6) = "NARRATIVE"
    resultSheet.Cells(5, 7) = "HOURS"
    resultSheet.Cells(5, 8) = "RATE"
    resultSheet.Cells(5, 9) = "GROSS UBS"
    resultSheet.Cells(5, 10) = "ADJ UBS"
    resultSheet.Cells(3, 12) = "ADMIN"
    resultSheet.Cells(3, 13) = "DNB"
    
    resultSheet.Range("A5:J5").Interior.ColorIndex = 5
    resultSheet.Range("L3:M3").Interior.ColorIndex = 5
    resultSheet.Range("A5:J5").Font.ColorIndex = 4
    resultSheet.Range("A5:J5").Font.Bold = True
    resultSheet.Range("L3:M3").Font.ColorIndex = 4
    resultSheet.Range("L3:M3").Font.Bold = True
    
    'Loop through data rows and update result sheet
    dataNumRows = dataSheet.Range("A1", dataSheet.Range("A1").End(xlDown)).Rows.Count

    currentResultRow = 7
    previousJobNumber = dataSheet.Cells(1, 4).Value

    For i = 1 To dataNumRows
        currentJobNumber = dataSheet.Cells(i, 4).Value
        
        If Right(Trim(currentJobNumber), 5) <> "Total" Then
            resultSheet.Cells(currentResultRow, 1) = dataSheet.Cells(i, 4) '"NEW JOB Number"
            resultSheet.Cells(currentResultRow, 2) = dataSheet.Cells(i, 6) '"JOB Description"
            resultSheet.Cells(currentResultRow, 3) = dataSheet.Cells(i, 5) '"OLD JOB Number"
            resultSheet.Cells(currentResultRow, 4) = dataSheet.Cells(i, 8) '"ENTRY Date"
            If Trim(dataSheet.Cells(i, 9)) = "" Then '"EMPLOYEE Name"
                resultSheet.Cells(currentResultRow, 5) = resultSheet.Cells(currentResultRow - 1, 5)
            Else
                resultSheet.Cells(currentResultRow, 5) = dataSheet.Cells(i, 9)
            End If
            resultSheet.Cells(currentResultRow, 6) = dataSheet.Cells(i, 12) '"NARRATIVE"
            resultSheet.Cells(currentResultRow, 7) = dataSheet.Cells(i, 13) '"HOURS"
            resultSheet.Cells(currentResultRow, 9) = dataSheet.Cells(i, 14) '"GROSS UBS"
            
            resultSheet.Cells(currentResultRow, 8).Formula = "=I" & CStr(currentResultRow) & "/G" & CStr(currentResultRow) '"Rate"
            resultSheet.Cells(currentResultRow, 10).Formula = "=G" & CStr(currentResultRow) & "*H" & CStr(currentResultRow) '"ADJ UBS"
    
            currentResultRow = currentResultRow + 1
            previousJobNumber = dataSheet.Cells(i, 4).Value
        End If
    Next i
    
    'Sort resultSheet data
    resultSheet.Range("A7:J" & CStr(currentResultRow - 1)).Sort Key1:=resultSheet.Range("A7"), _
                                                                Order1:=xlAscending, _
                                                                Key2:=resultSheet.Range("D7"), _
                                                                Order1:=xlAscending
                                                                
    'Add grey row between different job numbers
    previousJobNumber = resultSheet.Cells(7, 1).Value
    numberOfGreyRows = 0
    
    For i = 7 To currentResultRow - 1
        currentJobNumber = resultSheet.Cells(i + numberOfGreyRows, 1).Value
    
        If currentJobNumber <> previousJobNumber Then
            resultSheet.Rows(i + numberOfGreyRows).EntireRow.Insert
            resultSheet.Rows(i + numberOfGreyRows).Interior.ColorIndex = 48
            numberOfGreyRows = numberOfGreyRows + 1
        End If
        
        previousJobNumber = resultSheet.Cells(i + numberOfGreyRows, 1).Value
    Next i
    
    currentResultRow = currentResultRow + numberOfGreyRows
    
    'Create SUM formulas
    resultSheet.Cells(currentResultRow, 7).Formula = "=SUM(G7:G" & CStr(currentResultRow - 1) & ")"
    resultSheet.Cells(currentResultRow, 9).Formula = "=SUM(I7:I" & CStr(currentResultRow - 1) & ")"
    resultSheet.Cells(currentResultRow, 10).Formula = "=SUM(J7:J" & CStr(currentResultRow - 1) & ")"
    resultSheet.Cells(currentResultRow, 12).Formula = "=SUM(L7:L" & CStr(currentResultRow - 1) & ")"
    resultSheet.Cells(currentResultRow, 13).Formula = "=SUM(M7:M" & CStr(currentResultRow - 1) & ")"
    
    'Formatting SUM cells
    resultSheet.Cells(currentResultRow, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow, 7).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow, 7).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow, 9).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow, 9).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow, 10).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow, 10).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow, 12).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow, 12).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow, 12).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow, 12).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow, 13).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow, 13).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow, 13).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow, 13).Borders(xlEdgeTop).ColorIndex = 5
    
    'Adjust column widths and format
    resultSheet.Columns("A:M").AutoFit
    
    resultSheet.Columns("F").ColumnWidth = 63
    resultSheet.Columns("F").WrapText = True
    resultSheet.Columns("K").ColumnWidth = 1
    resultSheet.Columns("N").ColumnWidth = 1
    resultSheet.Columns("L:M").ColumnWidth = 8.5
    
    'Additional sum cells
    resultSheet.Cells(currentResultRow + 2, 9) = "Total Amount Check"
    resultSheet.Cells(currentResultRow + 3, 9) = "Invoiced"
    resultSheet.Cells(currentResultRow + 4, 9) = "Realization"
    resultSheet.Cells(currentResultRow + 5, 9) = "Rate/Hr (Billed)"
    
    resultSheet.Range("I" & CStr(currentResultRow + 2) & ":I" & CStr(currentResultRow + 5)).Font.Bold = True
    
    resultSheet.Cells(currentResultRow + 8, 7) = "Hours"
    resultSheet.Cells(currentResultRow + 8, 8) = "Rate"
    resultSheet.Range("G" & CStr(currentResultRow + 8) & ":H" & CStr(currentResultRow + 8)).Font.Bold = True

    'Populate unique employee names and SUMIFS formulas
    uniqueNames = WorksheetFunction.Unique(resultSheet.Range("E7:E" & CStr(currentResultRow - 1)))
    
    ReDim newUniqueNames(LBound(uniqueNames) To UBound(uniqueNames))
    For a = LBound(uniqueNames) To UBound(uniqueNames)
        If uniqueNames(a, 1) <> "" Then
            b = b + 1
            newUniqueNames(b) = uniqueNames(a, 1)
        End If
    Next a
    ReDim Preserve newUniqueNames(LBound(uniqueNames) To b)
    
    For n = LBound(newUniqueNames) To UBound(newUniqueNames)
        resultSheet.Cells(currentResultRow + 9 + n - 1, 5) = newUniqueNames(n)
        resultSheet.Cells(currentResultRow + 9 + n - 1, 7).Formula = _
                "=SUMIFS(G$7:G$" & CStr(currentResultRow - 1) & ",$E$7:$E$" & CStr(currentResultRow - 1) & ",$E" & CStr(currentResultRow + 9 + n - 1) & ")"
        resultSheet.Cells(currentResultRow + 9 + n - 1, 8) = _
                WorksheetFunction.Index(resultSheet.Range("H$7:H$" & CStr(currentResultRow - 1)), WorksheetFunction.Match(newUniqueNames(n), resultSheet.Range("$E$7:$E$" & CStr(currentResultRow - 1)), 0))
        resultSheet.Cells(currentResultRow + 9 + n - 1, 9).Formula = _
                "=SUMIFS(I$7:I$" & CStr(currentResultRow - 1) & ",$E$7:$E$" & CStr(currentResultRow - 1) & ",$E" & CStr(currentResultRow + 9 + n - 1) & ")/$H" & CStr(currentResultRow + 9 + n - 1)
        resultSheet.Cells(currentResultRow + 9 + n - 1, 10).Formula = _
                "=SUMIFS(J$7:J$" & CStr(currentResultRow - 1) & ",$E$7:$E$" & CStr(currentResultRow - 1) & ",$E" & CStr(currentResultRow + 9 + n - 1) & ")/$H" & CStr(currentResultRow + 9 + n - 1)
    Next n
    
    'Sort employees by rate
    resultSheet.Range("E" & CStr(currentResultRow + 8) & ":M" & CStr(currentResultRow + 9 + UBound(newUniqueNames))).Sort Key1:=resultSheet.Range("H" & CStr(currentResultRow + 8)), _
                                                                                                            Order1:=xlDescending, _
                                                                                                            Header:=xlYes
            
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 7).Formula = "=SUM(G" & CStr(currentResultRow + 9 + LBound(newUniqueNames) - 1) & ":G" & CStr(currentResultRow + 9 + UBound(newUniqueNames) - 1) & ")"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 9).Formula = "=SUM(I" & CStr(currentResultRow + 9 + LBound(newUniqueNames) - 1) & ":I" & CStr(currentResultRow + 9 + UBound(newUniqueNames) - 1) & ")"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 10).Formula = "=SUM(J" & CStr(currentResultRow + 9 + LBound(newUniqueNames) - 1) & ":J" & CStr(currentResultRow + 9 + UBound(newUniqueNames) - 1) & ")"
    
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 7).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 7).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 7).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 7).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 9).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 9).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 9).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 9).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 10).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 10).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames), 10).Borders(xlEdgeTop).ColorIndex = 5
    
    'Populate employee positions SUMIFS formulas
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 3, 5) = "Managing Director"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 4, 5) = "Director"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 5, 5) = "Senior Manager"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 6, 5) = "Manager"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 7, 5) = "Senior Associate"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 8, 5) = "Associate"
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 9, 5) = "Intern"

    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 7).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 7).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 7).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 9).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 9).Borders(xlEdgeTop).ColorIndex = 5
    
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 10).Borders(xlEdgeBottom).ColorIndex = 5
    resultSheet.Cells(currentResultRow + 9 + UBound(newUniqueNames) + 10, 10).Borders(xlEdgeTop).ColorIndex = 5
    
    'Format numeric columns
    resultSheet.Range("G7:J" & CStr(currentResultRow + 9 + UBound(newUniqueNames) + 10)).NumberFormat = "#,##0.00"
    resultSheet.Range("L7:M" & CStr(currentResultRow + 9 + UBound(newUniqueNames) + 10)).NumberFormat = "#,##0.00"
    
End Sub
