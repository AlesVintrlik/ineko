Attribute VB_Name = "Advanced"

Option Explicit
Dim x As Integer

Sub AktualizaceKontrolyKapacit()
   
    x = 1 ' Nastavení hodnoty, aby se v tomto pøípadì nespouštìl ukazatel prùbìhu
    
    Call VymazaniSouctuPodGrafem
    Call SumarizaceBoduProVsechnySloupce

End Sub

Sub VymazaniSouctuPodGrafem()

    Dim wsPlan As Worksheet
    Set wsPlan = ThisWorkbook.Sheets("Gantt")
       
    ' Najít první prázdný øádek ve sloupci B (pod poslední zakázkou)
    Dim nextRow As Long
    nextRow = wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp).row + 1
    
    ' Kontrola, zda jsou øádky k odstranìní v rámci rozsahu listu
    If nextRow > 0 And nextRow + 3 <= wsPlan.Rows.Count Then
        ' Smazat øádky nextRow až nextRow + 3 jednorázovì
        wsPlan.Rows(nextRow & ":" & nextRow + 3).Delete Shift:=xlUp
    Else
        MsgBox "Øádky k odstranìní jsou mimo rozsah listu.", vbCritical
        Exit Sub
    End If
    
End Sub

Sub SumarizaceBoduProVsechnySloupce()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")
    
    Dim cell As Range
    Dim progress As Double
    
    Dim totalPointsPriprava As Long
    Dim totalPointsSvarovani As Long
    Dim totalPointsMontaz As Long
    Dim totalPointsElektro As Long
    
    Dim LidiPriprava As Integer
    Dim LidiSvarovani As Integer
    Dim LidiMontaz As Integer
    Dim LidiElektro As Integer
    
    Dim holidays As Variant
    Dim holidayCell As Range
    Dim holidayRange As Range
    
     ' Potlaèení pøekreslování a aktualizace
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Sheets("Gantt").Range("B1").Select
    
    ' Naètení svátkù do pole
    Set holidayRange = Sheets("Svátky").Range("B2", Sheets("Svátky").Cells(Sheets("Svátky").Rows.Count, "B").End(xlUp))
    holidays = holidayRange.Value

    LidiPriprava = 1
    LidiSvarovani = 2
    LidiMontaz = 2
    LidiElektro = 1
    
    If x <> 1 Then
        ' Vytvoøení instance formuláøe frmProgress
        Load frmProgress
        frmProgress.Show vbModeless
    Else
    End If

    ' Najít poslední sloupec s daty ve druhém øádku
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Debug.Print "Poslední sloupec: " & lastCol

    ' Najít první prázdný øádek ve sloupci B (pod poslední zakázkou)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row + 1
    Debug.Print "NextRow: " & nextRow
    
    ' Kontrola, zda jsou øádky k odstranìní v rámci rozsahu listu
    If nextRow > 0 And nextRow + 3 <= ws.Rows.Count Then
        ' Smazat øádky nextRow až nextRow + 3 jednorázovì
        Debug.Print "Mazání øádkù: " & nextRow & " až " & nextRow + 3
        ws.Rows(nextRow & ":" & nextRow + 3).Delete Shift:=xlUp
    Else
        MsgBox "Øádky k odstranìní jsou mimo rozsah listu.", vbCritical
        Exit Sub
    End If

    Dim col As Long
    For col = 16 To lastCol ' Sloupce od O (èíslo 16) do posledního relevantního sloupce
        Debug.Print "Zpracování sloupce: " & col
        Dim compareValue As Double
        compareValue = ws.Cells(2, col).Value ' Uložíme hodnotu z buòky v druhém øádku pro aktuální sloupec
        Debug.Print "Hodnota: " & compareValue
        
        totalPointsPriprava = 0
        totalPointsSvarovani = 0
        totalPointsMontaz = 0
        totalPointsElektro = 0
        
        If x <> 1 Then
        
            ' Aktualizace prùbìhu
            progress = (col - 16) / (lastCol - 16 + 1)
            frmProgress.UpdateProgressBar progress
            frmProgress.lblProgressText.Caption = Round(progress * 100) & " %"
            
        Else
        End If
        
        ' Kontrola, zda compareValue je víkend nebo svátek
        If Weekday(compareValue, vbMonday) <= 5 Then
            Dim isHoliday As Boolean
            isHoliday = False

            For Each holidayCell In holidayRange
                If compareValue = holidayCell.Value Then
                    isHoliday = True
                    Exit For
                End If
            Next holidayCell

            ' Pokud compareValue není svátek, pokraèovat ve zpracování
            If Not isHoliday Then
                ' Projít všechny buòky ve sloupci od øádku 4 dolù
                For Each cell In ws.Range(ws.Cells(4, col), ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).row, col))
                    Dim row As Long
                    row = cell.row
'                    Debug.Print "Zpracování buòky: (" & row & ", " & col & ")"
                    
                    ' Kontrola podmínky a pøidìlení bodù pro sloupce E-F (Pøíprava)
                    If compareValue >= ws.Cells(row, "E").Value And compareValue <= ws.Cells(row, "F").Value Then
                        totalPointsPriprava = totalPointsPriprava + LidiPriprava
                    End If
                    
                    ' Kontrola podmínky a pøidìlení bodù pro sloupce G-H (Svaøování)
                    If compareValue >= ws.Cells(row, "G").Value And compareValue <= ws.Cells(row, "H").Value Then
                        totalPointsSvarovani = totalPointsSvarovani + LidiSvarovani
                    End If
                    
                    ' Kontrola podmínky a pøidìlení bodù pro sloupce I-J (Montáž)
                    If compareValue >= ws.Cells(row, "I").Value And compareValue <= ws.Cells(row, "J").Value Then
                        totalPointsMontaz = totalPointsMontaz + LidiMontaz
                    End If
                    
                    ' Kontrola podmínky a pøidìlení bodù pro sloupce K-L (Elektro)
                    If compareValue >= ws.Cells(row, "K").Value And compareValue <= ws.Cells(row, "L").Value Then
                        totalPointsElektro = totalPointsElektro + LidiElektro
                    End If
                Next cell
            End If
        End If
        
        ' Pokud compareValue není víkend ani svátek, zapsat hodnoty do aktuálního sloupce
        If Weekday(compareValue, vbMonday) <= 5 And Not isHoliday Then
            ws.Cells(nextRow, col).Value = totalPointsPriprava
            ws.Cells(nextRow + 1, col).Value = totalPointsSvarovani
            ws.Cells(nextRow + 2, col).Value = totalPointsMontaz
            ws.Cells(nextRow + 3, col).Value = totalPointsElektro
        End If
    Next col

    ' Nastavit formátování pro zapsané hodnoty
    With ws.Range(ws.Cells(nextRow, 16), ws.Cells(nextRow + 3, lastCol))
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0"
        .Orientation = 0
    End With
    
    ' Vypne podmínìné formátování listu
    ws.EnableFormatConditionsCalculation = False

    ' Pøidání podmínìného formátování
    With ws.Range(ws.Cells(nextRow, 16), ws.Cells(nextRow, lastCol))
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=LidiPriprava"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 217, 217)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=LidiPriprava"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 225, 129)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=LidiPriprava"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(218, 242, 208)
    End With

    With ws.Range(ws.Cells(nextRow + 1, 16), ws.Cells(nextRow + 1, lastCol))
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=LidiSvarovani"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 217, 217)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=LidiSvarovani"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 225, 129)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=LidiSvarovani"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(218, 242, 208)
    End With

    With ws.Range(ws.Cells(nextRow + 2, 16), ws.Cells(nextRow + 2, lastCol))
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=LidiMontaz"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 217, 217)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=LidiMontaz"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 225, 129)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=LidiMontaz"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(218, 242, 208)
    End With

    With ws.Range(ws.Cells(nextRow + 3, 16), ws.Cells(nextRow + 3, lastCol))
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=LidiElektro"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 217, 217)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=LidiElektro"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 225, 129)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=LidiElektro"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(218, 242, 208)
    End With
    
    ' Zapne podmínìné formátování listu
    ws.EnableFormatConditionsCalculation = True
    
    ' Zapnutí pøekreslování a aktualizace
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If x <> 1 Then
    
        ' Zavøít frmProgress
        Unload frmProgress
    Else
    End If

End Sub

Sub ZrusitVsechnyFiltry()
    Dim ws As Worksheet
    Set ws = Worksheets("Gantt")
    
    ' Zkontroluje, zda je na listu AutoFilter a zda je nìjaký filtr aktivní
    If ws.AutoFilterMode Then
        ' Pokud jsou aktivní filtry, zruší je
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
End Sub

Sub ZobrazitMinimalistickyGantt()
    ' Aktivace listu Gantt
    Sheets("Gantt").Activate
    
    ' Skrýt møížku, záhlaví a panel vzorcù
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
End Sub

