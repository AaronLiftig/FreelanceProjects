Option Explicit

Sub Step_1_OrganizeBeforeSheet()

    Call Step_1a_SplitAccountNumber
            
    Call Step_1b_ConvertTradeDateToDate
    
    Call Step_1c_SortBeforeTable
                      
End Sub


Sub Step_1a_SplitAccountNumber()
    
    Dim beforeSheet As Worksheet
    Dim actNum As Object
    Dim i As Variant
    Dim char As String
    
    Set beforeSheet = Sheets("Before")
    
    'Creates Act# Letters and Numbers columns
    For Each actNum In beforeSheet.Range("Table1[Account Number]")
        If IsEmpty(actNum) Then
            Exit For
        End If
        
        actNum = Trim(actNum)
        
        For i = 1 To Len(actNum)
            
            char = Mid(actNum, i, 1)
            
            If char Like "[A-Za-z]" Then
                GoTo Continue1
            ElseIf char Like "[0-9]" Then
                If i = 1 Then
                    beforeSheet.Cells(actNum.Row, 10).Value = CInt(actNum)
                Else
                    beforeSheet.Cells(actNum.Row, 9).Value = CStr(Mid(actNum, 1, i - 1))
                    beforeSheet.Cells(actNum.Row, 10).Value = CInt(Mid(actNum, i, Len(actNum)))
                End If
                GoTo Continue2
            Else
                MsgBox ("Unexpected Account Number character, " & CStr(actNum) & ", in row " & CStr(actNum.Row))
                Exit Sub
            End If
            
Continue1:
            Next i
            
Continue2:
        Next actNum
        
End Sub


Sub Step_1b_ConvertTradeDateToDate()
    
    Dim beforeSheet As Worksheet
    Dim tradeDate As Object
    
    Set beforeSheet = Sheets("Before")
    
    'Make sure Trade Date column is formated as date
    For Each tradeDate In beforeSheet.Range("Table1[Trade Date]")
        'Used if table is too long with blanks at bottom
        If IsEmpty(tradeDate) Then
            Exit For
        End If
        
        beforeSheet.Cells(tradeDate.Row, 5).Value = CDate(tradeDate)
        
        Next tradeDate

End Sub


Sub Step_1c_SortBeforeTable()
    
    Dim beforeSheet As Worksheet
    
    Set beforeSheet = Sheets("Before")
    
    'Sort by Act# Letters, Act# Numbers, and Trade Date
    With beforeSheet.Range("Table1")
        .Cells.Sort Key1:=.Columns(5), Order1:=xlAscending, _
                    Key2:=.Columns(6), Order2:=xlDescending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
        .Cells.Sort Key1:=.Columns(9), Order1:=xlAscending, _
                    Key2:=.Columns(10), Order2:=xlAscending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
    End With

End Sub


Sub Step_2_PopulateMapSheet()

    Dim beforeSheet As Worksheet
    Dim mapSheet As Worksheet
    Dim actNum As Object
    Dim lastActNum As String
    Dim actCounter As Integer
    Dim rowCounter As Integer
    
    Set beforeSheet = Sheets("Before")
    Set mapSheet = Sheets("Map")
    actCounter = 1
    rowCounter = 0
    
    lastActNum = beforeSheet.Range("Table1[Account Number]").Cells(1).Value
    
    For Each actNum In beforeSheet.Range("Table1[Account Number]")
        If actNum <> lastActNum Then
            mapSheet.Range("Table2[Account Number]").Cells(actCounter).Value = lastActNum
            mapSheet.Range("Table2[Account Name]").Cells(actCounter).Value = beforeSheet.Range("Table1[Account Name]").Cells(rowCounter + 1).Value
            mapSheet.Range("Table2[Account Reference Number]").Cells(actCounter).Value = "Account " & CStr(actCounter)
            lastActNum = actNum
            actCounter = actCounter + 1
        End If
        
        rowCounter = rowCounter + 1
        
        Next actNum
    
    mapSheet.Range("Table2[Account Number]").Cells(actCounter).Value = lastActNum
    mapSheet.Range("Table2[Account Name]").Cells(actCounter).Value = beforeSheet.Range("Table1[Account Name]").Cells(rowCounter).Value
    mapSheet.Range("Table2[Account Reference Number]").Cells(actCounter).Value = "Account " & CStr(actCounter)
        
End Sub


Sub Step_3_PopulateAfterSheet()

    Dim beforeSheet As Worksheet
    Dim mapSheet As Worksheet
    Dim afterSheet As Worksheet
    Dim actNum As Object
    Dim lastActNum As String
    Dim beforeRowCount As Integer
    Dim afterRowCount As Integer
    Dim mapRowCount As Integer
    Dim quantity As Long
    
    Set beforeSheet = Sheets("Before")
    Set mapSheet = Sheets("Map")
    Set afterSheet = Sheets("After")
    mapRowCount = 1
    beforeRowCount = 1
    afterRowCount = 1
    
    lastActNum = ""
    
    For Each actNum In beforeSheet.Range("Table1[Account Number]")
        If actNum = lastActNum Then
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Transaction Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
            quantity = beforeSheet.Range("Table1[Quantity]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = quantity
            afterSheet.Range("Table3[Price]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Price]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Amount]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Amount]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount).Value = afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount - 1).Value + quantity
            
            beforeRowCount = beforeRowCount + 1
            afterRowCount = afterRowCount + 1
            
        ElseIf lastActNum = "" Then
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = "Beginning Holdings"
            afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
            quantity = mapSheet.Range("Table2[PreClass]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = quantity
            afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount).Value = quantity
            
            afterRowCount = afterRowCount + 1
            lastActNum = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
        
        Else
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Transaction Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
            quantity = beforeSheet.Range("Table1[Quantity]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = quantity
            afterSheet.Range("Table3[Price]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Price]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Amount]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Amount]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount).Value = afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount - 1).Value + quantity
            
            afterRowCount = afterRowCount + 1
            
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = "End Holdings"
            afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount - 1).Value
            
            afterRowCount = afterRowCount + 1
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = "*"
            afterRowCount = afterRowCount + 1
            beforeRowCount = beforeRowCount + 1
            mapRowCount = mapRowCount + 1
            lastActNum = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            
            afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
            afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = "Beginning Holdings"
            afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
            quantity = mapSheet.Range("Table2[PreClass]").Rows(mapRowCount).Value
            afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = quantity
            afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount).Value = quantity
            
            afterRowCount = afterRowCount + 1
            lastActNum = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
          
        End If

        afterSheet.Range("Table3[Account Number]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Number]").Rows(beforeRowCount).Value
        afterSheet.Range("Table3[Account Name]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Account Name]").Rows(beforeRowCount).Value
        afterSheet.Range("Table3[Account Reference Number]").Rows(afterRowCount).Value = mapSheet.Range("Table2[Account Reference Number]").Rows(mapRowCount).Value
        afterSheet.Range("Table3[Security Type]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Security Type]").Rows(beforeRowCount).Value
        afterSheet.Range("Table3[Transaction Type]").Rows(afterRowCount).Value = "End Holdings"
        afterSheet.Range("Table3[Trade Date]").Rows(afterRowCount).Value = beforeSheet.Range("Table1[Trade Date]").Rows(beforeRowCount).Value
        afterSheet.Range("Table3[Quantity]").Rows(afterRowCount).Value = afterSheet.Range("Table3[PreClass + Quantity]").Rows(afterRowCount - 1).Value
        
        Next actNum
    
End Sub
