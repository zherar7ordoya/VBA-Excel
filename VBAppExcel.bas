Attribute VB_Name = "VBAppExcel"


'*----------------------------------------*'
' Desarrollado con ♡ por © Gerardo Tordoya '
'*----------------------------------------*'


Option Explicit


#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)            'MS Office 32 Bit
#End If


Option Private Module
    Dim E_DIMENSION As Variant
    Dim E_SUBROUTINE As Variant
    Dim EDITIONDATE As String
    
Sub AutoOpen(Optional Hider As Byte)
    Dim ctrl As IRibbonControl
End Sub
    
'Callback for buttonGenerate onAction
Sub mainSubroutine(control As IRibbonControl)
On Error GoTo crashedParty

    'Err.Raise vbObjectError + 513, "Module1::Test()", "My custom error."
    E_DIMENSION = InputBox("Dimension?", "Subroutine", 0)
    If Val(E_DIMENSION) = 0 Or Val(E_DIMENSION) Mod 4 <> 0 Then Exit Sub
    
    'E_SUBROUTINE = InputBox("Subroutine", , 0)
    'If Val(Val(Left(E_SUBROUTINE, 4)) & Val(Right(E_SUBROUTINE, 2))) <> Val(Year(Now()) & Month(Now())) Then Exit Sub
    
    Call sweepPrevious
    Call recoverReport
    Call makeTable
    Call evaluateLocation
    Call roundHeight
    Call getMeasurement
    Call makeFinals
    
    MsgBox "Okay!", vbOKOnly + vbInformation, "OK"
    
    Exit Sub
    
crashedParty:
    
    MsgBox "Knockout...", vbOKOnly + vbCritical, "KO"
    
End Sub
    
    
Sub sweepPrevious(Optional Hider As Byte)
On Error GoTo crashedParty
    
    Sheets("InputSheet").Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Sub recoverReport(Optional Hider As Byte)
On Error GoTo crashedParty
    
    Dim inputFile As String
    inputFile = Application.GetOpenFilename(FileFilter:="Soft-Data,*.xls*")
    
    Workbooks.Open inputFile
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy Before:=Workbooks(ThisWorkbook.Name).Sheets(1)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "InputSheet"
    Windows(Right(inputFile, Len(inputFile) - InStrRev(inputFile, "\"))).Activate
    ActiveWindow.Close
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Sub makeTable(Optional Hider As Byte)
On Error GoTo crashedParty
    
    Dim cellText As String
    Dim totalRows
    Dim i
    
    Sheets("InputSheet").Select
    
    'Insert empty column
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'Fill clasification column
    Range("B2").Select
    Selection.End(xlDown).Select
    totalRows = ActiveCell.Row
    Range("B2").Select
    
    For i = 2 To totalRows
        If IsDate(ActiveCell) Then
            ActiveCell.Offset(0, 1).Select
            cellText = Selection.Value
            ActiveCell.Offset(0, -2).Select
            Selection.Value = cellText
            ActiveCell.Offset(1, 1).Select
            Sleep (25)
        Else
            ActiveCell.Offset(0, -1).Select
            Selection.Value = cellText
            ActiveCell.Offset(1, 1).Select
            Sleep (25)
        End If
    Next i
    
    Range("B2").Select
    EDITIONDATE = Selection.Value
    Range("F2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Cut
    Range("B2").Select
    ActiveSheet.Paste
    Range("A2").Select
    Selection.FormulaR1C1 = "Section"
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection.Address), , xlYes).Name = "InputData"
    
    'UnnecessaryData
    ActiveSheet.ListObjects("InputData").Range.AutoFilter Field:=16, Criteria1:=Array("Ubic.Obtenida", "="), Operator:=xlFilterValues
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    'RemoveFilter
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Selection.AutoFilter
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Sub evaluateLocation(Optional Hider As Byte)
On Error GoTo crashedParty
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Evaluate Location"
    Selection.AutoFilter
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=EVALUATE_LOCATION([@Section],[@[Forma de Pago]],[@[Ubic.Pretendida]],[@[Nro pag]])"
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Function EVALUATE_LOCATION(editionSection, _
                           paymentMethod, _
                           intendedLocation, _
                           pageNumber)
On Error GoTo crashedParty
    
    EVALUATE_LOCATION = " "
    If editionSection <> "Cuerpo Central" _
    Or paymentMethod <> "Cuenta Corriente" _
    Or intendedLocation = "SER" _
    Then Exit Function
    If pageNumber = 0 Then EVALUATE_LOCATION = "Error"
    
    Select Case intendedLocation
        Case "PES":     Exit Function
        Case "SU":      Exit Function
        Case "OBI":     Exit Function
        Case "TAPA":    If pageNumber = 1 Then Exit Function
        Case "PAG3":    If pageNumber = 3 Then Exit Function
        Case "PAG5":    If pageNumber = 5 Then Exit Function
        Case "PAG7":    If pageNumber = 7 Then Exit Function
        Case "PAG9":    If pageNumber = 9 Then Exit Function
        Case "CTAPA":   If pageNumber = (E_DIMENSION * 1) Then Exit Function
        
        Case "SUDC":    If Val(pageNumber) > (E_DIMENSION / 2) Then Exit Function
        Case "IAC":     If Val(pageNumber) < (E_DIMENSION / 2) And WorksheetFunction.IsOdd(pageNumber) Then Exit Function
        Case "IDC":     If Val(pageNumber) > (E_DIMENSION / 2) And WorksheetFunction.IsOdd(pageNumber) Then Exit Function
        Case "PDC":     If Val(pageNumber) > (E_DIMENSION / 2) And WorksheetFunction.IsEven(pageNumber) Then Exit Function
        Case "PAC":     If Val(pageNumber) <= (E_DIMENSION / 2) And WorksheetFunction.IsEven(pageNumber) Then Exit Function
    End Select
    
    EVALUATE_LOCATION = "Error"
    
    Exit Function
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Function


Sub roundHeight()
On Error GoTo crashedParty
    
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Alto"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=ROUND([@AltoCm],0)"
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Sub getMeasurement()
On Error GoTo crashedParty
    
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Medida"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=[@Col]&""x""&IF(LEN([@Alto])=2,[@Alto],""0""&[@Alto])"
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub


Sub makeFinals(Optional Hider As Byte)
On Error GoTo crashedParty
    
    Dim totalRows As Integer
    totalRows = Range("A1").CurrentRegion.Rows.Count - 1
    
    Columns("T:T").Select
    Selection.Delete Shift:=xlToLeft
    
    Sheets("OutputSheet").Select
    
    With ActiveSheet.PageSetup
        .LeftHeader = "&""Ubuntu Mono,Normal" & Chr(34) & E_DIMENSION & " PAGES"
        .CenterHeader = "&""Ubuntu Mono,Normal" & Chr(34) & totalRows & " ADS"
        .RightHeader = "&""Ubuntu Mono,Normal" & Chr(34) & UCase(Format(EDITIONDATE, "dddd dd mmmm yyyy"))
    End With
    
    ActiveWorkbook.RefreshAll
    
    Cells.Select
    Selection.RowHeight = 32
    
    ActiveSheet.PivotTables("DataPivotTable").PivotSelect "Page[All;Total]", xlDataAndLabel, True
    Selection.Rows.AutoFit
    
    Rows("1:1").Select
    Selection.Rows.AutoFit
    Range("M1").Select
    
    Exit Sub
    
crashedParty:
    
    MsgBox Error.Description & vbNewLine & _
           Error.Number, _
           vbOKOnly + vbCritical + vbMsgBoxRight, _
           "Error"
    End
    
End Sub