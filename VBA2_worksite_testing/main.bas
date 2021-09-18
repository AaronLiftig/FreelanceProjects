Option Explicit

Sub AlarmMacro()

Dim wb As Workbook
Dim sh, sh2 As Worksheet
Dim DN_val, portId_val As String
Dim descr_val, inUse_val, polarity_val, severity_val As String
Dim r, r2 As Integer
Dim okay1, okay2, okay3, okay4, okay5, okay8, okay9, okay16 As Integer 'For checking inUse values

Set wb = Workbooks("New Data.xlsx")
Set sh = wb.Sheets("Dataset")

'Create "DN Sort" column
sh.Cells(2, 20) = "DN Sort"

r = 3
DN_val = Trim(sh.Cells(r, 2))

Do Until DN_val = ""
    If Mid(Right(DN_val, 2), 1, 1) = "-" Then
        sh.Cells(r, 20) = Mid(DN_val, 1, Len(DN_val) - 1)
    Else
        sh.Cells(r, 20) = Mid(DN_val, 1, Len(DN_val) - 2)
    End If

    r = r + 1
    DN_val = Trim(sh.Cells(r, 2))
Loop


'Sort by "DN Sort" and "Port ID" columns
With sh.Range("A3:T" & CStr(r))
    .Sort _
    Key1:=sh.Range("T3"), Order1:=xlAscending, _
    Key2:=sh.Range("Q3"), Order2:=xlAscending, DataOption2:=xlSortTextAsNumbers
End With


'Create "Check" column
sh.Cells(2, 21) = "All Check"
sh.Cells(2, 22) = "Non-Required Check"

'Check each row
r = 3
DN_val = Trim(sh.Cells(r, 20))
descr_val = Trim(sh.Cells(r, 14))
inUse_val = Trim(sh.Cells(r, 15))
polarity_val = Trim(sh.Cells(r, 16))
portId_val = Trim(sh.Cells(r, 17))
severity_val = Trim(sh.Cells(r, 18))

Do Until DN_val = ""
    If portId_val = 1 Then
        If descr_val = "DOOR_OPEN" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 2 Then
        If descr_val = "TECH_ON" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 3 Then
        If descr_val = "PWR_AC_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 4 Then
        If descr_val = "ENV_HIGH_LOW_TEMP" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 5 Then
        If descr_val = "ENV_HVAC_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 6 Then
        If descr_val = "ENV_SMOKE" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "ENV_SMOKE" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 7 Then
        If descr_val = "PWR_AC_SURGEPROT" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "PWR_AC_SURGEPROT" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 8 Then
        If descr_val = "PWR_BAY_MJ" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 9 Then
        If descr_val = "PWR_BAY_MN" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 10 Then
        If descr_val = "TWR_BCN_STRB" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "TWR_BCN_STRB" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 11 Then
        If descr_val = "TWR_SIDELIGHT" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "TWR_SIDELIGHT" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 12 Then
        If descr_val = "GEN_RUN" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "GEN_RUN" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 13 Then
        If descr_val = "GEN_FAIL" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "GEN_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 14 Then
        If descr_val = "GEN_XFER" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "GEN_XFER" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 15 Then
        If descr_val = "PWR_BAY_LOW_VOLTS" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "PWR_BAY_LOW_VOLTS" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 16 Then
        If descr_val = "PWR_BREAKER_DC" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        sh.Cells(r, 22) = "Required"
    ElseIf portId_val = 17 Then
        If descr_val = "MW_DEHYDRATOR_FAIL" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "MW_DEHYDRATOR_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 18 Then
        If descr_val = "COPPER_THEFT" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "COPPER_THEFT" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    ElseIf portId_val = 19 Then
        If descr_val = "MW_MJ" And inUse_val = "false" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 21) = "okay"
        Else
            sh.Cells(r, 21) = "error"
        End If
        If descr_val = "MW_MJ" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            sh.Cells(r, 22) = "okay"
        Else
            sh.Cells(r, 22) = "error"
        End If
    Else
        sh.Cells(r, 21) = "ID out of range"
        sh.Cells(r, 22) = "ID out of range"
    End If
    
    r = r + 1
    DN_val = Trim(sh.Cells(r, 2))
    descr_val = Trim(sh.Cells(r, 14))
    inUse_val = Trim(sh.Cells(r, 15))
    polarity_val = Trim(sh.Cells(r, 16))
    portId_val = Trim(sh.Cells(r, 17))
    severity_val = Trim(sh.Cells(r, 18))
Loop


'Create new sheet for "inUse" verification
On Error GoTo sheetExists
    wb.Sheets.Add.Name = "inUse Check"

Set sh2 = wb.Sheets("inUse Check")

sh2.Cells(1, 1) = "DN"
sh2.Cells(1, 2) = "inUse okay?"
sh2.Cells(1, 3) = "ID1"
sh2.Cells(1, 4) = "ID2"
sh2.Cells(1, 5) = "ID3"
sh2.Cells(1, 6) = "ID4"
sh2.Cells(1, 7) = "ID5"
sh2.Cells(1, 8) = "ID8"
sh2.Cells(1, 9) = "ID9"
sh2.Cells(1, 10) = "ID16"

r2 = 2

r = 3
DN_val = Trim(sh.Cells(r, 20))
descr_val = Trim(sh.Cells(r, 14))
inUse_val = Trim(sh.Cells(r, 15))
polarity_val = Trim(sh.Cells(r, 16))
portId_val = Trim(sh.Cells(r, 17))
severity_val = Trim(sh.Cells(r, 18))

okay1 = 0
okay2 = 0
okay3 = 0
okay4 = 0
okay5 = 0
okay8 = 0
okay9 = 0
okay16 = 0

Do Until DN_val = ""
    If portId_val = 1 Then
        If descr_val = "DOOR_OPEN" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            okay1 = 1
        Else
            okay1 = 2
        End If
    ElseIf portId_val = 2 Then
        If descr_val = "TECH_ON" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            okay2 = 1
        Else
            okay2 = 2
        End If
    ElseIf portId_val = 3 Then
        If descr_val = "PWR_AC_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            okay3 = 1
        Else
            okay3 = 2
        End If
    ElseIf portId_val = 4 Then
        If descr_val = "ENV_HIGH_LOW_TEMP" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            okay4 = 1
        Else
            okay4 = 2
        End If
    ElseIf portId_val = 5 Then
        If descr_val = "ENV_HVAC_FAIL" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            okay5 = 1
        Else
            okay5 = 2
        End If
    ElseIf portId_val = 8 Then
        If descr_val = "PWR_BAY_MJ" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Major" Then
            okay8 = 1
        Else
            okay8 = 2
        End If
    ElseIf portId_val = 9 Then
        If descr_val = "PWR_BAY_MN" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Minor" Then
            okay9 = 1
        Else
            okay9 = 2
        End If
    ElseIf portId_val = 16 Then
        If descr_val = "PWR_BREAKER_DC" And inUse_val = "true" And polarity_val = "Normally_closed" And severity_val = "Critical" Then
            okay16 = 1
        Else
            okay16 = 2
        End If
    End If
    
    If DN_val <> Trim(sh.Cells(r + 1, 20)) Then
        sh2.Cells(r2, 1) = DN_val
        If okay1 = 1 And okay2 = 1 And okay3 = 1 And okay4 = 1 And okay5 = 1 And okay8 = 1 And okay9 = 1 And okay16 = 1 Then
            sh2.Cells(r2, 2) = "okay"
        Else
            sh2.Cells(r2, 2) = "error"
        End If
        
        If okay1 = 1 Then
            sh2.Cells(r2, 3) = "okay"
        ElseIf okay1 = 2 Then
            sh2.Cells(r2, 3) = "error"
        Else
            sh2.Cells(r2, 3) = "DNE"
        End If
        If okay2 = 1 Then
            sh2.Cells(r2, 4) = "okay"
        ElseIf okay2 = 2 Then
            sh2.Cells(r2, 4) = "error"
        Else
            sh2.Cells(r2, 4) = "DNE"
        End If
        If okay3 = 1 Then
            sh2.Cells(r2, 5) = "okay"
        ElseIf okay3 = 2 Then
            sh2.Cells(r2, 5) = "error"
        Else
            sh2.Cells(r2, 5) = "DNE"
        End If
        If okay4 = 1 Then
            sh2.Cells(r2, 6) = "okay"
        ElseIf okay4 = 2 Then
            sh2.Cells(r2, 6) = "error"
        Else
            sh2.Cells(r2, 6) = "DNE"
        End If
        If okay5 = 1 Then
            sh2.Cells(r2, 7) = "okay"
        ElseIf okay5 = 2 Then
            sh2.Cells(r2, 7) = "error"
        Else
            sh2.Cells(r2, 7) = "DNE"
        End If
        If okay8 = 1 Then
            sh2.Cells(r2, 8) = "okay"
        ElseIf okay8 = 2 Then
            sh2.Cells(r2, 8) = "error"
        Else
            sh2.Cells(r2, 8) = "DNE"
        End If
        If okay9 = 1 Then
            sh2.Cells(r2, 9) = "okay"
        ElseIf okay9 = 2 Then
            sh2.Cells(r2, 9) = "error"
        Else
            sh2.Cells(r2, 9) = "DNE"
        End If
        If okay16 = 1 Then
            sh2.Cells(r2, 10) = "okay"
        ElseIf okay16 = 2 Then
            sh2.Cells(r2, 10) = "error"
        Else
            sh2.Cells(r2, 10) = "DNE"
        End If
        
        okay1 = 0
        okay2 = 0
        okay3 = 0
        okay4 = 0
        okay5 = 0
        okay8 = 0
        okay9 = 0
        okay16 = 0
        
        r2 = r2 + 1
    End If
    
    r = r + 1
    DN_val = Trim(sh.Cells(r, 20))
    descr_val = Trim(sh.Cells(r, 14))
    inUse_val = Trim(sh.Cells(r, 15))
    polarity_val = Trim(sh.Cells(r, 16))
    portId_val = Trim(sh.Cells(r, 17))
    severity_val = Trim(sh.Cells(r, 18))
Loop


sheetExists:
    Select Case Err.Number
        Case 1004
            MsgBox ("The 'inUse Check' sheet already exists. Please delete that sheet and the additional sheet created and rerun the macro.")
    End Select
    Exit Sub
    
End Sub
