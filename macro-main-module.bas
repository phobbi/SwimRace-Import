Attribute VB_Name = "Main"

Sub Auto()
    
'
' Makro pro úpravu dat k importu do programu SwimRace
'
    
'
    Count = Range("A1").End(xlDown).Row
    sheetExists = False
    Dim ActiveRange As String
    Dim CopySource As String
    Dim PasteRange As String
    Dim DistCell As String
    Dim BirthYear As String
    Dim SexCell As String
    Row = 2
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Prosím èekejte..."
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:H").Select
    Selection.Merge True
    Selection.UnMerge
    Columns("G:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("Q:Q").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("H1:L1").Select
    Selection.ClearContents
    Columns("G:L").Select
    Selection.Merge True
    Selection.UnMerge
    Columns("H:L").Select
    Selection.Delete Shift:=xlToLeft
    Range("I1:N1").Select
    Selection.ClearContents
    Columns("H:N").Select
    Selection.Merge True
    Columns("H:N").Select
    Range("H111").Activate
    Selection.UnMerge
    Columns("I:P").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("D:D").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Oddíl"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Zkratka oddílu"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Typ"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Rok narození"
    
    For Each Sheet In Worksheets
        If "pøihlášky" = Sheet.Name Then
            sheetExists = True
        End If
        Next Sheet
        
        If sheetExists = False Then
            With ThisWorkbook
                Set WS = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
                WS.Name = "pøihlášky"
            End With
        End If
        
        Worksheets("pøihlášky").Activate
        
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "Oddíl"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "zk# Oddíl"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Pøíjmení"
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Jméno"
        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Rok nar"
        Range("F1").Select
        ActiveCell.FormulaR1C1 = "M/Ž"
        Range("G1").Select
        ActiveCell.FormulaR1C1 = "Disc"
        Range("H1").Select
        ActiveCell.FormulaR1C1 = "èas"
        
        Worksheets(1).Activate
        
        For i = 2 To Count
            Let ActiveRange = "A" & i
            Range(ActiveRange).Select
            If ActiveCell.FormulaR1C1 = "POFM - závodník" Or ActiveCell.FormulaR1C1 = "POFM - nezávodník" Then
                ActiveCell.FormulaR1C1 = "Plavecký oddíl Frýdek-Místek"
                Let ActiveRange = "B" & i
                Range(ActiveRange).Select
                ActiveCell.FormulaR1C1 = "POFM"
            End If
            Let ActiveRange = "F" & i
            Range(ActiveRange).Select
            If ActiveCell.FormulaR1C1 = "Chlapec" Then
                ActiveCell.FormulaR1C1 = "M"
            Else
                ActiveCell.FormulaR1C1 = "Ž"
            End If
            Next i
            
            
            For i = 2 To Count
                Let ActiveRange = "G" & i
                Range(ActiveRange).Select
                If ActiveCell.FormulaR1C1 = "Pøedplavec" Then
                    Let CopySource = "A" & i & ":" & "F" & i
                    Let PasteRange = "A" & Row
                    Let DistCell = "G" & Row
                    Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                    Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "26"
                    Row = Row + 1
                End If
                If ActiveCell.FormulaR1C1 = "Hendikepovaný" Then
                    Let CopySource = "A" & i & ":" & "F" & i
                    Let PasteRange = "A" & Row
                    Let DistCell = "G" & Row
                    Let BirthYear = "E" & i
                    Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                    Range(BirthYear).Select
                    If Year(Date) - CInt(ActiveCell.FormulaR1C1) > 16 Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "12"
                    Else
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "11"
                    End If
                    Row = Row + 1
                End If
                If ActiveCell.FormulaR1C1 = "Ostatní" Or ActiveCell.FormulaR1C1 = "B&#283;žný plavec" Then
                    Let CopySource = "A" & i & ":" & "F" & i
                    Let PasteRange = "A" & Row
                    Let DistCell = "G" & Row
                    Let BirthYear = "E" & i
                    Let SexCell = "F" & i
                    Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                    Let ActiveRange = "H" & i
                    Range(ActiveRange).Select
                    
                    If ActiveCell.FormulaR1C1 = "" Then
                        If CInt(Year(Date)) - CInt(Range(BirthYear).FormulaR1C1) < 6 Then
                            Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "26"
                            Row = Row + 1
                        End If
                    End If
                    
                    If ActiveCell.FormulaR1C1 = "Prsa 16 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "2"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 16 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "6"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "10"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "10"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Prsa 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "4"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "8"
                        Row = Row + 1
                    End If
                    
                    If ActiveCell.FormulaR1C1 = "Prsa 16 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "1"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 16 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "5"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "9"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "9"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Prsa 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "3"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "7"
                        Row = Row + 1
                    End If
                    
                    Let ActiveRange = "I" & i
                    Range(ActiveRange).Select
                    Let PasteRange = "A" & Row
                    Let DistCell = "G" & Row
                    
                    If ActiveCell.FormulaR1C1 = "Prsa 16 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "2"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 16 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "6"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "10"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "10"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Prsa 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "4"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 33 m" And Range(SexCell).FormulaR1C1 = "M" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "8"
                        Row = Row + 1
                    End If
                    
                    If ActiveCell.FormulaR1C1 = "Prsa 16 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "1"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 16 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "5"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "9"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Volný zpùsob 33m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "9"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Prsa 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "3"
                        Row = Row + 1
                    End If
                    If ActiveCell.FormulaR1C1 = "Znak 33 m" And Range(SexCell).FormulaR1C1 = "Ž" Then
                        Worksheets(1).Range(CopySource).Copy Worksheets("pøihlášky").Range(PasteRange)
                        Worksheets("pøihlášky").Range(DistCell).FormulaR1C1 = "7"
                        Row = Row + 1
                    End If
                End If
                Next i
                
                Application.StatusBar = False
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Worksheets("pøihlášky").Activate
                i = MsgBox("Data jsou pøipravena k importu do aplikace SwimRace.", vbOKOnly + vbInformation, "Dokonèeno")
                
            End Sub
            
